Attribute VB_Name = "modHttpServer"
Option Explicit

' =============================================================================
' modHttpServer
' WindowsのWinsock APIを直接呼び出す、学習専用の簡易HTTPサーバーです。
'
' 重要:
'   - 実運用のWebサーバーではありません。
'   - 127.0.0.1だけで待ち受けるため、同じPCからだけ接続できます。
'   - 1接続ずつ処理する、HTTP/1.1の最小実装です。
'   - Server_StartInNewExcelで、同じブックを別Excelプロセスに開けます。
'   - ping.exe、PowerShell、curlなどの外部コマンドは起動しません。
'
' Application.OnTimeで約1秒ごとに非ブロッキングソケットを確認します。
' 長いDo WhileループでExcelを占有しないため、停止やブック保存が行いやすい構成です。
' =============================================================================

Public Const HTTP_SERVER_PORT As Long = 18080

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const SOMAXCONN As Long = &H7FFFFFFF
Private Const FIONBIO As Long = &H8004667E
Private Const SOCKET_ERROR As Long = -1
Private Const INVALID_SOCKET As Long = -1
Private Const WSAEWOULDBLOCK As Long = 10035
Private Const SOL_SOCKET As Long = &HFFFF&
' WindowsではSO_REUSEADDRを安易に使わず、学習用ポートを排他的に確保します。
Private Const SO_EXCLUSIVEADDRUSE As Long = -5
Private Const RECEIVE_CHUNK_SIZE As Long = 4096
Private Const MAX_REQUEST_BYTES As Long = 65536
Private Const MAX_DELAY_MILLISECONDS As Long = 10000

' Windows SDKのWSADATAは、64ビット版と32ビット版でフィールド順が異なります。
' ポインター型だけをLongPtrへ変えるのでは不十分なため、Win64で構造体そのものを
' 分けています。ここがずれるとWSAStartupが隣接メモリを書き換える危険があります。
#If Win64 Then
    Private Type WSADATA
        wVersion As Integer
        wHighVersion As Integer
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpVendorInfo As LongPtr
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
    End Type
#Else
    Private Type WSADATA
        wVersion As Integer
        wHighVersion As Integer
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpVendorInfo As LongPtr
    End Type
#End If

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal wVersionRequired As Integer, _
    ByRef lpWSAData As WSADATA) As Long

Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare PtrSafe Function ws_socket Lib "ws2_32.dll" Alias "socket" ( _
    ByVal af As Long, _
    ByVal socketType As Long, _
    ByVal protocol As Long) As LongPtr

Private Declare PtrSafe Function ws_bind Lib "ws2_32.dll" Alias "bind" ( _
    ByVal socketHandle As LongPtr, _
    ByRef socketAddress As SOCKADDR_IN, _
    ByVal addressLength As Long) As Long

Private Declare PtrSafe Function ws_listen Lib "ws2_32.dll" Alias "listen" ( _
    ByVal socketHandle As LongPtr, _
    ByVal backlog As Long) As Long

Private Declare PtrSafe Function ws_accept Lib "ws2_32.dll" Alias "accept" ( _
    ByVal socketHandle As LongPtr, _
    ByRef clientAddress As SOCKADDR_IN, _
    ByRef addressLength As Long) As LongPtr

Private Declare PtrSafe Function ws_recv Lib "ws2_32.dll" Alias "recv" ( _
    ByVal socketHandle As LongPtr, _
    ByRef receiveBuffer As Any, _
    ByVal bufferLength As Long, _
    ByVal flags As Long) As Long

Private Declare PtrSafe Function ws_send Lib "ws2_32.dll" Alias "send" ( _
    ByVal socketHandle As LongPtr, _
    ByRef sendBuffer As Any, _
    ByVal bufferLength As Long, _
    ByVal flags As Long) As Long

Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, _
    ByVal command As Long, _
    ByRef argumentValue As Long) As Long

Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, _
    ByVal level As Long, _
    ByVal optionName As Long, _
    ByRef optionValue As Any, _
    ByVal optionLength As Long) As Long

Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr) As Long

Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" ( _
    ByVal addressText As String) As Long

Private Declare PtrSafe Function htons Lib "ws2_32.dll" ( _
    ByVal hostShort As Integer) As Integer

Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
    ByVal milliseconds As Long)

Private m_listenSocket As LongPtr
Private m_clientSocket As LongPtr
Private m_serverRunning As Boolean
Private m_winsockStarted As Boolean
Private m_pollScheduled As Boolean
Private m_nextPollTime As Date
Private m_pendingBytes() As Byte
Private m_pendingLength As Long
Private m_acceptStartedAt As Double
Private m_currentTraceId As String
Private m_stopAfterResponse As Boolean
Private m_serverExcelInstance As Excel.Application
Private m_serverWorkbookInstance As Excel.Workbook

' 現在のマクロブックを、別のExcelプロセスで読み取り専用として開き、
' そのコピー側でServer_Startを実行します。
'
' CreateObjectによるExcelのCOMオートメーションであり、Shell、PowerShell、curlは
' 起動しません。元ブックとサーバー側コピーは別プロセスなので、/delayのSleep中も
' 元ブックは操作できます。
Public Sub Server_StartInNewExcel()
    Dim traceId As String
    Dim startedAt As Double
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String
    Dim serverMacroName As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "先にこのブックをマクロ有効ブック（.xlsm）として保存してください。", _
            vbInformation
        Exit Sub
    End If

    If Not (m_serverExcelInstance Is Nothing) Then
        MsgBox "このブックから起動したサーバー用Excelは、すでに存在します。", _
            vbInformation
        Exit Sub
    End If

    On Error GoTo LaunchError

    traceId = Trace_NewId("LAUNCH")
    startedAt = Trace_StartTimer()
    Trace_Write traceId, "CLIENT", "SERVER_LAUNCH", "LOCAL", _
        "別Excelプロセスを作成", "STARTING", startedAt

    Set m_serverExcelInstance = CreateObject("Excel.Application")
    m_serverExcelInstance.Visible = True

    ' 同じファイルは元ブックで編集中のため、サーバー側は読み取り専用で開きます。
    ' サーバーのTraceはそのExcelプロセスのメモリ上で確認できます。
    Set m_serverWorkbookInstance = m_serverExcelInstance.Workbooks.Open( _
        Filename:=ThisWorkbook.FullName, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False)

    serverMacroName = "'" _
        & Replace(m_serverWorkbookInstance.Name, "'", "''") _
        & "'!Server_Start"

    ' サーバー側の過去ログを消し、今回の待受けだけを観察しやすくします。
    m_serverExcelInstance.Run "'" _
        & Replace(m_serverWorkbookInstance.Name, "'", "''") _
        & "'!Trace_Clear"
    m_serverExcelInstance.Run serverMacroName

    Trace_Write traceId, "CLIENT", "SERVER_LAUNCH", "LOCAL", _
        "別ExcelプロセスでServer_Startを実行", "OK", startedAt
    Exit Sub

LaunchError:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description

    Trace_Write traceId, "CLIENT", "SERVER_LAUNCH", "LOCAL", _
        "Err=" & CStr(savedErrorNumber) & " / " & savedErrorDescription, _
        "ERROR", startedAt

    On Error Resume Next
    If Not (m_serverWorkbookInstance Is Nothing) Then
        m_serverWorkbookInstance.Close SaveChanges:=False
    End If
    If Not (m_serverExcelInstance Is Nothing) Then
        m_serverExcelInstance.Quit
    End If
    Set m_serverWorkbookInstance = Nothing
    Set m_serverExcelInstance = Nothing
    On Error GoTo 0

    MsgBox "サーバー用Excelを起動できませんでした。" & vbCrLf _
        & savedErrorDescription, vbExclamation
End Sub

' Server_StartInNewExcelで起動したExcelへServer_Stopを依頼し、
' 読み取り専用ブックを保存せず閉じて、そのExcelプロセスも終了します。
Public Sub Server_StopInNewExcel()
    Dim traceId As String
    Dim serverMacroName As String

    traceId = Trace_NewId("LAUNCH")

    If m_serverExcelInstance Is Nothing Then
        MsgBox "このブックから起動したサーバー用Excelは見つかりません。", _
            vbInformation
        Exit Sub
    End If

    On Error Resume Next

    If Not (m_serverWorkbookInstance Is Nothing) Then
        serverMacroName = "'" _
            & Replace(m_serverWorkbookInstance.Name, "'", "''") _
            & "'!Server_Stop"
        m_serverExcelInstance.Run serverMacroName
        m_serverWorkbookInstance.Close SaveChanges:=False
    End If

    m_serverExcelInstance.Quit
    Set m_serverWorkbookInstance = Nothing
    Set m_serverExcelInstance = Nothing

    If Err.Number = 0 Then
        Trace_Write traceId, "CLIENT", "SERVER_CLOSE", "LOCAL", _
            "サーバー用Excelプロセスを終了", "OK"
    Else
        Trace_Write traceId, "CLIENT", "SERVER_CLOSE", "LOCAL", _
            "終了処理の一部でErr=" & CStr(Err.Number), "WARNING"
    End If

    On Error GoTo 0
End Sub

' サーバーを開始します。
' 同じPCだけから到達できる127.0.0.1:18080へbindします。
Public Sub Server_Start()
    Dim wsaDataValue As WSADATA
    Dim serverAddress As SOCKADDR_IN
    Dim nonBlockingMode As Long
    Dim exclusiveAddressUse As Long
    Dim resultCode As Long
    Dim startTraceId As String
    Dim startedAt As Double
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    If m_serverRunning Then
        MsgBox "簡易HTTPサーバーはすでに起動しています。", vbInformation
        Exit Sub
    End If

    On Error GoTo StartError

    startTraceId = Trace_NewId("SERVER")
    startedAt = Trace_StartTimer()
    m_listenSocket = INVALID_SOCKET
    m_clientSocket = INVALID_SOCKET

    ' Winsock 2.2を初期化します。&H0202はバージョン2.2を表します。
    resultCode = WSAStartup(&H202, wsaDataValue)
    If resultCode <> 0 Then
        Err.Raise vbObjectError + 2200, "Server_Start", _
            "WSAStartupに失敗しました。コード=" & CStr(resultCode)
    End If
    m_winsockStarted = True
    Trace_Write startTraceId, "SERVER", "WSA_STARTUP", "LOCAL", _
        "Winsock 2.2を初期化", "OK", startedAt

    ' IPv4/TCPの待受けソケットを作ります。
    m_listenSocket = ws_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If m_listenSocket = INVALID_SOCKET Then
        RaiseWinsockError "socket"
    End If

    ' 2つの教材サーバーが同じポートへ同時にbindし、応答先が不定になることを
    ' 防ぎます。使用中なら後から起動した側を明確なエラーにします。
    exclusiveAddressUse = 1
    If setsockopt(m_listenSocket, SOL_SOCKET, SO_EXCLUSIVEADDRUSE, _
                  exclusiveAddressUse, LenB(exclusiveAddressUse)) = SOCKET_ERROR Then
        RaiseWinsockError "setsockopt(SO_EXCLUSIVEADDRUSE)"
    End If

    serverAddress.sin_family = AF_INET
    serverAddress.sin_port = htons(CInt(HTTP_SERVER_PORT))
    serverAddress.sin_addr = inet_addr("127.0.0.1")

    If ws_bind(m_listenSocket, serverAddress, LenB(serverAddress)) = SOCKET_ERROR Then
        RaiseWinsockError "bind"
    End If

    If ws_listen(m_listenSocket, SOMAXCONN) = SOCKET_ERROR Then
        RaiseWinsockError "listen"
    End If

    ' acceptでExcelを待たせないよう、待受けソケットを非ブロッキングにします。
    nonBlockingMode = 1
    If ioctlsocket(m_listenSocket, FIONBIO, nonBlockingMode) = SOCKET_ERROR Then
        RaiseWinsockError "ioctlsocket"
    End If

    m_serverRunning = True
    Trace_Write startTraceId, "SERVER", "LISTEN", "LOCAL", _
        "http://127.0.0.1:" & CStr(HTTP_SERVER_PORT), "READY", startedAt

    Server_ScheduleNextPoll
    Exit Sub

StartError:
    ' 後始末やTrace_Writeの途中でErrの内容が変わる可能性があるため、
    ' 最初に番号と説明を退避します。
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    Trace_Write startTraceId, "SERVER", "START", "LOCAL", _
        "Err=" & CStr(savedErrorNumber) & " / " & savedErrorDescription, _
        "ERROR", startedAt
    Server_CloseSocketsAndWinsock
    MsgBox "簡易HTTPサーバーを開始できませんでした。" & vbCrLf _
        & savedErrorDescription, vbExclamation
End Sub

' サーバーを停止します。
' OnTime予約、接続ソケット、待受けソケット、Winsockの順に片付けます。
Public Sub Server_Stop()
    Dim traceId As String

    traceId = m_currentTraceId
    If Len(traceId) = 0 Then traceId = Trace_NewId("SERVER")

    m_serverRunning = False
    m_stopAfterResponse = False

    On Error Resume Next
    If m_pollScheduled Then
        Application.OnTime EarliestTime:=m_nextPollTime, _
            Procedure:=Server_PollProcedureName(), Schedule:=False
    End If
    m_pollScheduled = False
    On Error GoTo 0

    Server_CloseSocketsAndWinsock
    Trace_Write traceId, "SERVER", "STOP", "LOCAL", _
        "待受けとWinsockを終了", "STOPPED"
End Sub

' Application.OnTimeから約1秒ごとに呼ばれるポーリング処理です。
' 接続受付と受信を少量ずつ進め、処理後に次回実行を予約します。
Public Sub Server_Poll()
    Dim requestHeader As String
    Dim requestBody() As Byte
    Dim requestBodyLength As Long

    m_pollScheduled = False
    If Not m_serverRunning Then Exit Sub

    On Error GoTo PollError

    If Not Server_HasClient() Then
        Server_TryAcceptClient
    End If

    If Server_HasClient() Then
        Server_ReceiveAvailableBytes

        If Server_HasClient() Then
            If Server_TryBuildRequest( _
                requestHeader, requestBody, requestBodyLength) Then

                Server_HandleRequest requestHeader, requestBody, requestBodyLength

                If m_stopAfterResponse Then
                    Server_Stop
                    Exit Sub
                End If
            End If
        End If
    End If

    Server_ScheduleNextPoll
    Exit Sub

PollError:
    Trace_Write m_currentTraceId, "SERVER", "POLL", "LOCAL", _
        "Err=" & CStr(Err.Number) & " / " & Err.Description, "ERROR"
    Server_CloseClient

    If m_serverRunning Then Server_ScheduleNextPoll
End Sub

Private Sub Server_TryAcceptClient()
    Dim clientAddress As SOCKADDR_IN
    Dim addressLength As Long
    Dim nonBlockingMode As Long
    Dim errorCode As Long

    addressLength = LenB(clientAddress)
    m_clientSocket = ws_accept(m_listenSocket, clientAddress, addressLength)

    If m_clientSocket = INVALID_SOCKET Then
        errorCode = WSAGetLastError()

        ' WSAEWOULDBLOCKは「今は接続要求がない」という通常状態です。
        ' エラーとして記録するとTraceが埋め尽くされるため、そのまま戻ります。
        If errorCode = WSAEWOULDBLOCK Then Exit Sub
        RaiseWinsockError "accept"
    End If

    nonBlockingMode = 1
    If ioctlsocket(m_clientSocket, FIONBIO, nonBlockingMode) = SOCKET_ERROR Then
        RaiseWinsockError "ioctlsocket(client)"
    End If

    m_pendingLength = 0
    Erase m_pendingBytes
    m_acceptStartedAt = Trace_StartTimer()
    ' 新しい接続が前回のTraceIdを誤って引き継がないよう、
    ' HTTPヘッダーを解析するまでは空に戻します。
    m_currentTraceId = vbNullString
End Sub

Private Sub Server_ReceiveAvailableBytes()
    Dim chunk(0 To RECEIVE_CHUNK_SIZE - 1) As Byte
    Dim receivedBytes As Long
    Dim errorCode As Long

    Do
        receivedBytes = ws_recv( _
            m_clientSocket, chunk(0), RECEIVE_CHUNK_SIZE, 0)

        If receivedBytes > 0 Then
            Server_AppendBytes chunk, receivedBytes

            If m_pendingLength > MAX_REQUEST_BYTES Then
                Err.Raise vbObjectError + 2201, "Server_ReceiveAvailableBytes", _
                    "要求サイズが上限（65536バイト）を超えました。"
            End If

        ElseIf receivedBytes = 0 Then
            ' recv=0は相手側が接続を閉じたことを表します。
            Trace_Write m_currentTraceId, "SERVER", "RECEIVE", "IN", _
                "クライアントが接続を閉じました。", "CLOSED", m_acceptStartedAt
            Server_CloseClient
            Exit Do

        Else
            errorCode = WSAGetLastError()

            If errorCode = WSAEWOULDBLOCK Then
                Exit Do
            Else
                RaiseWinsockError "recv"
            End If
        End If
    Loop
End Sub

Private Sub Server_AppendBytes( _
    ByRef sourceBytes() As Byte, _
    ByVal sourceLength As Long)

    Dim sourceIndex As Long
    Dim oldLength As Long

    oldLength = m_pendingLength
    m_pendingLength = m_pendingLength + sourceLength

    ' 未初期化配列へのReDim Preserveは環境差を避けるため、
    ' 最初の受信だけ通常のReDimを使います。
    If oldLength = 0 Then
        ReDim m_pendingBytes(0 To m_pendingLength - 1)
    Else
        ReDim Preserve m_pendingBytes(0 To m_pendingLength - 1)
    End If

    For sourceIndex = 0 To sourceLength - 1
        m_pendingBytes(oldLength + sourceIndex) = sourceBytes(sourceIndex)
    Next sourceIndex
End Sub

' HTTPヘッダー末尾（CRLF CRLF）とContent-Lengthを確認し、
' 1要求分がそろったときだけTrueを返します。
Private Function Server_TryBuildRequest( _
    ByRef requestHeader As String, _
    ByRef requestBody() As Byte, _
    ByRef requestBodyLength As Long) As Boolean

    Dim headerEndPosition As Long
    Dim contentLength As Long
    Dim bodyIndex As Long

    headerEndPosition = Server_FindHeaderEnd()
    If headerEndPosition = 0 Then Exit Function

    requestHeader = Server_AsciiFromPendingBytes(headerEndPosition)
    contentLength = Server_GetContentLength(requestHeader)

    If contentLength < 0 Then
        Err.Raise vbObjectError + 2202, "Server_TryBuildRequest", _
            "Content-Lengthが数値ではありません。"
    End If

    If contentLength > MAX_REQUEST_BYTES Then
        Err.Raise vbObjectError + 2203, "Server_TryBuildRequest", _
            "Content-Lengthが許容上限を超えています。"
    End If

    If m_pendingLength < headerEndPosition + contentLength Then
        Exit Function
    End If

    requestBodyLength = contentLength

    If contentLength > 0 Then
        ReDim requestBody(0 To contentLength - 1)

        For bodyIndex = 0 To contentLength - 1
            requestBody(bodyIndex) = m_pendingBytes(headerEndPosition + bodyIndex)
        Next bodyIndex
    Else
        Erase requestBody
    End If

    Server_TryBuildRequest = True
End Function

' 戻り値はヘッダー終了直後の「バイト位置」です。
' 0はまだCRLF CRLFが見つからないことを表します。
Private Function Server_FindHeaderEnd() As Long
    Dim byteIndex As Long

    If m_pendingLength < 4 Then Exit Function

    For byteIndex = 0 To m_pendingLength - 4
        If m_pendingBytes(byteIndex) = 13 _
            And m_pendingBytes(byteIndex + 1) = 10 _
            And m_pendingBytes(byteIndex + 2) = 13 _
            And m_pendingBytes(byteIndex + 3) = 10 Then

            Server_FindHeaderEnd = byteIndex + 4
            Exit Function
        End If
    Next byteIndex
End Function

' HTTPヘッダー名はASCIIなので、1バイトずつChr$で文字列化できます。
' 本文は日本語を含む可能性があるため、この関数では変換しません。
Private Function Server_AsciiFromPendingBytes(ByVal byteCount As Long) As String
    Dim byteIndex As Long
    Dim result As String

    For byteIndex = 0 To byteCount - 1
        result = result & Chr$(m_pendingBytes(byteIndex))
    Next byteIndex

    Server_AsciiFromPendingBytes = result
End Function

Private Function Server_GetContentLength(ByVal requestHeader As String) As Long
    Dim headerValue As String

    headerValue = Server_GetHeaderValue(requestHeader, "Content-Length")

    If Len(headerValue) = 0 Then
        Server_GetContentLength = 0
    ElseIf IsNumeric(headerValue) Then
        Server_GetContentLength = CLng(headerValue)
    Else
        Server_GetContentLength = -1
    End If
End Function

Private Sub Server_HandleRequest( _
    ByVal requestHeader As String, _
    ByRef requestBody() As Byte, _
    ByVal requestBodyLength As Long)

    Dim requestLines As Variant
    Dim requestParts As Variant
    Dim requestMethod As String
    Dim requestTarget As String
    Dim routePath As String
    Dim queryPosition As Long
    Dim responseBody() As Byte
    Dim responseContentType As String
    Dim delayMilliseconds As Long
    Dim statusCode As Long
    Dim reasonPhrase As String

    requestLines = Split(requestHeader, vbCrLf)
    requestParts = Split(CStr(requestLines(0)), " ")

    If UBound(requestParts) < 1 Then
        Err.Raise vbObjectError + 2204, "Server_HandleRequest", _
            "HTTPリクエスト行を解析できません。"
    End If

    requestMethod = UCase$(CStr(requestParts(0)))
    requestTarget = CStr(requestParts(1))
    m_currentTraceId = Server_GetHeaderValue(requestHeader, "X-Trace-Id")

    ' 値を応答ヘッダーへ載せるため、改行などを含まない限定文字列だけを採用します。
    ' 生のTCPクライアントから不正ヘッダーを送られても、ヘッダー注入を起こしません。
    If Not Server_IsSafeTraceId(m_currentTraceId) Then
        m_currentTraceId = Trace_NewId("SERVER")
    End If

    Trace_Write m_currentTraceId, "SERVER", "ACCEPT", "IN", _
        "クライアント接続を受付", "OK", m_acceptStartedAt
    Trace_Write m_currentTraceId, "SERVER", "RECEIVE", "IN", _
        "受信=" & CStr(m_pendingLength) & "バイト", "OK", m_acceptStartedAt
    Trace_Write m_currentTraceId, "SERVER", "PARSE", "LOCAL", _
        requestMethod & " " & requestTarget, "OK", m_acceptStartedAt

    queryPosition = InStr(1, requestTarget, "?", vbBinaryCompare)
    If queryPosition > 0 Then
        routePath = Left$(requestTarget, queryPosition - 1)
    Else
        routePath = requestTarget
    End If

    statusCode = 200
    reasonPhrase = "OK"
    responseContentType = "application/json; charset=utf-8"

    Select Case routePath
        Case "/hello"
            responseBody = Utf8_Encode( _
                "{""message"":""Hello from VBA HTTP server""}")

        Case "/echo"
            If requestBodyLength > 0 Then
                responseBody = requestBody
                responseContentType = Server_GetHeaderValue( _
                    requestHeader, "Content-Type")
                If Len(responseContentType) = 0 Then
                    responseContentType = "application/octet-stream"
                End If
            Else
                responseBody = Utf8_Encode("{""echo"":null}")
            End If

        Case "/delay"
            delayMilliseconds = Server_GetDelayMilliseconds(requestTarget)

            Trace_Write m_currentTraceId, "SERVER", "DELAY", "LOCAL", _
                CStr(delayMilliseconds) & "ミリ秒待機", "WAITING", m_acceptStartedAt

            ' Sleepで停止するのはサーバー用の別Excelプロセスだけです。
            ' クライアント側Excelが操作できることが非同期学習のポイントです。
            Sleep delayMilliseconds

            responseBody = Utf8_Encode( _
                "{""delayedMilliseconds"":" & CStr(delayMilliseconds) & "}")

        Case "/status/500"
            statusCode = 500
            reasonPhrase = "Internal Server Error"
            responseBody = Utf8_Encode( _
                "{""error"":""Intentional test error""}")

        Case "/shutdown"
            responseBody = Utf8_Encode( _
                "{""message"":""Server will stop""}")
            m_stopAfterResponse = True

        Case Else
            statusCode = 404
            reasonPhrase = "Not Found"
            responseBody = Utf8_Encode( _
                "{""error"":""Route not found""}")
    End Select

    Server_SendResponse statusCode, reasonPhrase, _
        responseContentType, responseBody

    Server_CloseClient
End Sub

Private Sub Server_SendResponse( _
    ByVal statusCode As Long, _
    ByVal reasonPhrase As String, _
    ByVal contentType As String, _
    ByRef responseBody() As Byte)

    Dim responseHeader As String
    Dim headerBytes() As Byte
    Dim allBytes() As Byte
    Dim headerLength As Long
    Dim bodyLength As Long
    Dim byteIndex As Long

    bodyLength = Bytes_Length(responseBody)

    responseHeader = "HTTP/1.1 " & CStr(statusCode) & " " & reasonPhrase & vbCrLf _
        & "Content-Type: " & contentType & vbCrLf _
        & "Content-Length: " & CStr(bodyLength) & vbCrLf _
        & "Connection: close" & vbCrLf _
        & "X-Trace-Id: " & m_currentTraceId & vbCrLf _
        & vbCrLf

    headerBytes = Utf8_Encode(responseHeader)
    headerLength = Bytes_Length(headerBytes)
    ReDim allBytes(0 To headerLength + bodyLength - 1)

    For byteIndex = 0 To headerLength - 1
        allBytes(byteIndex) = headerBytes(byteIndex)
    Next byteIndex

    For byteIndex = 0 To bodyLength - 1
        allBytes(headerLength + byteIndex) = responseBody(byteIndex)
    Next byteIndex

    Server_SendAll allBytes

    Trace_Write m_currentTraceId, "SERVER", "RESPONSE", "OUT", _
        "HTTP " & CStr(statusCode) & " / 本文=" & CStr(bodyLength) & "バイト", _
        "SENT", m_acceptStartedAt
End Sub

' sendは1回で全バイトを送るとは限らないため、送信済み位置を進めながら繰り返します。
' 非ブロッキングソケットが一時的に送れない場合は10ミリ秒待ち、最大5秒で打ち切ります。
Private Sub Server_SendAll(ByRef sendBytes() As Byte)
    Dim totalLength As Long
    Dim sentTotal As Long
    Dim sentNow As Long
    Dim retryCount As Long
    Dim errorCode As Long

    totalLength = Bytes_Length(sendBytes)

    Do While sentTotal < totalLength
        sentNow = ws_send( _
            m_clientSocket, sendBytes(sentTotal), totalLength - sentTotal, 0)

        If sentNow > 0 Then
            sentTotal = sentTotal + sentNow
            retryCount = 0
        Else
            errorCode = WSAGetLastError()

            If errorCode <> WSAEWOULDBLOCK Then
                RaiseWinsockError "send"
            End If

            retryCount = retryCount + 1
            If retryCount > 500 Then
                Err.Raise vbObjectError + 2205, "Server_SendAll", _
                    "送信可能になるまでの待機が5秒を超えました。"
            End If

            DoEvents
            Sleep 10
        End If
    Loop
End Sub

Private Function Server_GetHeaderValue( _
    ByVal requestHeader As String, _
    ByVal headerName As String) As String

    Dim requestLines As Variant
    Dim lineIndex As Long
    Dim separatorPosition As Long
    Dim currentName As String

    requestLines = Split(requestHeader, vbCrLf)

    For lineIndex = LBound(requestLines) To UBound(requestLines)
        separatorPosition = InStr(1, CStr(requestLines(lineIndex)), ":", vbBinaryCompare)

        If separatorPosition > 0 Then
            currentName = Trim$(Left$( _
                CStr(requestLines(lineIndex)), separatorPosition - 1))

            If StrComp(currentName, headerName, vbTextCompare) = 0 Then
                Server_GetHeaderValue = Trim$(Mid$( _
                    CStr(requestLines(lineIndex)), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next lineIndex
End Function

' 応答ヘッダーへ安全に載せられるTraceIdか確認します。
' この教材が生成するIDで使う英数字とハイフンだけを許可します。
Private Function Server_IsSafeTraceId(ByVal value As String) As Boolean
    Dim characterIndex As Long
    Dim currentCharacter As String

    If Len(value) = 0 Or Len(value) > 128 Then Exit Function

    For characterIndex = 1 To Len(value)
        currentCharacter = Mid$(value, characterIndex, 1)
        If Not ((currentCharacter >= "0" And currentCharacter <= "9") Or _
                (currentCharacter >= "A" And currentCharacter <= "Z") Or _
                (currentCharacter >= "a" And currentCharacter <= "z") Or _
                currentCharacter = "-") Then
            Exit Function
        End If
    Next characterIndex

    Server_IsSafeTraceId = True
End Function

Private Function Server_GetDelayMilliseconds( _
    ByVal requestTarget As String) As Long

    Dim parameterPosition As Long
    Dim parameterText As String
    Dim ampersandPosition As Long
    Dim parsedValue As Double

    parameterPosition = InStr(1, requestTarget, "ms=", vbTextCompare)

    If parameterPosition = 0 Then
        Server_GetDelayMilliseconds = 1000
        Exit Function
    End If

    parameterText = Mid$(requestTarget, parameterPosition + 3)
    ampersandPosition = InStr(1, parameterText, "&", vbBinaryCompare)
    If ampersandPosition > 0 Then
        parameterText = Left$(parameterText, ampersandPosition - 1)
    End If

    If Not IsNumeric(parameterText) Then
        Server_GetDelayMilliseconds = 1000
        Exit Function
    End If

    parsedValue = CDbl(parameterText)
    If parsedValue < 0# Then parsedValue = 0#
    If parsedValue > MAX_DELAY_MILLISECONDS Then
        parsedValue = MAX_DELAY_MILLISECONDS
    End If

    Server_GetDelayMilliseconds = CLng(parsedValue)
End Function

Private Function Server_HasClient() As Boolean
    Server_HasClient = (m_clientSocket <> 0 _
        And m_clientSocket <> INVALID_SOCKET)
End Function

Private Sub Server_ScheduleNextPoll()
    If Not m_serverRunning Then Exit Sub

    m_nextPollTime = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=m_nextPollTime, _
        Procedure:=Server_PollProcedureName(), Schedule:=True
    m_pollScheduled = True
End Sub

Private Function Server_PollProcedureName() As String
    ' ブック名を付け、別ブックに同名のマクロがあっても取り違えないようにします。
    Server_PollProcedureName = "'" _
        & Replace(ThisWorkbook.Name, "'", "''") _
        & "'!Server_Poll"
End Function

Private Sub Server_CloseClient()
    On Error Resume Next

    If Server_HasClient() Then
        closesocket m_clientSocket
    End If

    m_clientSocket = INVALID_SOCKET
    m_pendingLength = 0
    Erase m_pendingBytes
    On Error GoTo 0
End Sub

Private Sub Server_CloseSocketsAndWinsock()
    On Error Resume Next

    Server_CloseClient

    If m_listenSocket <> 0 And m_listenSocket <> INVALID_SOCKET Then
        closesocket m_listenSocket
    End If
    m_listenSocket = INVALID_SOCKET

    If m_winsockStarted Then
        WSACleanup
        m_winsockStarted = False
    End If

    On Error GoTo 0
End Sub

Private Sub RaiseWinsockError(ByVal operationName As String)
    Dim errorCode As Long

    errorCode = WSAGetLastError()
    Err.Raise vbObjectError + 2299, operationName, _
        operationName & "に失敗しました。Winsockエラー=" & CStr(errorCode)
End Sub
