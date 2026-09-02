VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' CAsyncHttpClient
' MSXML2.XMLHTTP60をイベント付きで保持する、非同期HTTPクライアントです。
'
' m_http.Openの第3引数をTrueにすると、Sendの直後にVBAへ制御が戻ります。
' 応答の進行はonreadystatechangeイベントで通知されるため、待機のための
' Do WhileループやDoEventsの連続実行は不要です。
'
' 使用前にVBEの「ツール」→「参照設定」で、
' 「Microsoft XML, v6.0」を有効にしてください。
' =============================================================================

Private WithEvents m_http As MSXML2.XMLHTTP60
Attribute m_http.VB_VarHelpID = -1

Private m_traceId As String
Private m_startedAt As Double
Private m_lastReadyState As Long
Private m_isCompleted As Boolean
Private m_requestMethod As String
Private m_requestUrl As String

Public Property Get TraceId() As String
    TraceId = m_traceId
End Property

Public Property Get IsCompleted() As Boolean
    IsCompleted = m_isCompleted
End Property

' 非同期HTTP要求を開始します。
' X-Trace-Idをヘッダーに載せることで、サーバー側でも同じ識別子を使えます。
Public Sub SendAsync( _
    ByVal requestMethod As String, _
    ByVal requestUrl As String, _
    Optional ByVal requestBody As String = vbNullString, _
    Optional ByVal contentType As String = "application/json; charset=utf-8")

    On Error GoTo SendError

    m_traceId = Trace_NewId("HTTP")
    m_startedAt = Trace_StartTimer()
    m_lastReadyState = -1
    m_isCompleted = False
    m_requestMethod = UCase$(Trim$(requestMethod))
    m_requestUrl = requestUrl

    Trace_Write m_traceId, "CLIENT", "CREATE", "LOCAL", _
        m_requestMethod & " " & m_requestUrl, "OK", m_startedAt

    Set m_http = New MSXML2.XMLHTTP60

    ' 第3引数のTrueが「非同期」を表します。
    ' FalseにするとSendが応答完了まで戻らない同期通信になります。
    m_http.Open m_requestMethod, m_requestUrl, True

    m_http.setRequestHeader "Accept", "application/json"
    m_http.setRequestHeader "X-Trace-Id", m_traceId

    If Len(requestBody) > 0 Then
        m_http.setRequestHeader "Content-Type", contentType
    End If

    Trace_Write m_traceId, "CLIENT", "SEND", "OUT", _
        "本文文字数=" & CStr(Len(requestBody)), "SENDING", m_startedAt

    m_http.send requestBody
    Exit Sub

SendError:
    m_isCompleted = True
    Trace_Write m_traceId, "CLIENT", "SEND", "LOCAL", _
        "Err=" & CStr(Err.Number) & " / " & Err.Description, _
        "ERROR", m_startedAt
    Client_RequestCompleted m_traceId
End Sub

' 利用者が通信を中止したい場合に呼び出します。
' abort後も環境によってreadyStateイベントが発生することがあるため、
' m_isCompletedを先にTrueにして二重完了を防ぎます。
Public Sub Cancel()
    If m_isCompleted Then Exit Sub

    On Error Resume Next
    m_isCompleted = True
    m_http.abort
    Trace_Write m_traceId, "CLIENT", "CANCEL", "LOCAL", _
        "利用者が要求を中止しました。", "CANCELLED", m_startedAt
    On Error GoTo 0
End Sub

' readyStateが変化するたびにMSXMLから呼び出されます。
' 同じ状態が複数回通知される実装もあるため、直前の状態と同じ場合は記録しません。
Private Sub m_http_onreadystatechange()
    Dim currentState As Long
    Dim statusCode As Long
    Dim responseText As String
    Dim responseTraceId As String

    On Error GoTo EventError

    If m_http Is Nothing Then Exit Sub

    currentState = m_http.readyState

    If currentState <> m_lastReadyState Then
        m_lastReadyState = currentState
        Trace_Write m_traceId, "CLIENT", "READY_STATE", "IN", _
            CStr(currentState) & " = " & ReadyStateName(currentState), _
            "WAITING", m_startedAt
    End If

    If currentState <> 4 Then Exit Sub
    If m_isCompleted Then Exit Sub

    m_isCompleted = True
    statusCode = SafeStatusCode()
    responseText = SafeResponseText()
    responseTraceId = SafeResponseHeader("X-Trace-Id")

    If statusCode > 0 Then
        Trace_Write m_traceId, "CLIENT", "TRACE_ID", "IN", _
            "Response X-Trace-Id=" & responseTraceId, _
            IIf(StrComp(responseTraceId, m_traceId, vbBinaryCompare) = 0, _
                "MATCH", "MISMATCH"), _
            m_startedAt
    End If

    Trace_Write m_traceId, "CLIENT", "RESPONSE", "IN", _
        responseText, "HTTP " & CStr(statusCode), m_startedAt

    Trace_Write m_traceId, "CLIENT", "COMPLETE", "LOCAL", _
        m_requestMethod & " " & m_requestUrl, _
        IIf(statusCode >= 200 And statusCode < 300, "OK", "HTTP_ERROR"), _
        m_startedAt

    ' modClientDemo側のCollectionから完了済みオブジェクトを外します。
    ' これを行わないと、テストを繰り返すたびにオブジェクトが残り続けます。
    Client_RequestCompleted m_traceId
    Exit Sub

EventError:
    If Not m_isCompleted Then
        m_isCompleted = True
        Trace_Write m_traceId, "CLIENT", "EVENT", "LOCAL", _
            "Err=" & CStr(Err.Number) & " / " & Err.Description, _
            "ERROR", m_startedAt
        Client_RequestCompleted m_traceId
    End If
End Sub

Private Function SafeStatusCode() As Long
    On Error GoTo StatusUnavailable

    SafeStatusCode = m_http.Status
    Exit Function

StatusUnavailable:
    ' サーバーに接続できない場合など、HTTP応答自体が存在しないケースです。
    SafeStatusCode = 0
End Function

' 指定した応答ヘッダーを安全に取得します。通信エラーでHTTP応答自体がない場合は
' 空文字を返し、元のエラー経路を別の例外で隠さないようにします。
Private Function SafeResponseHeader(ByVal headerName As String) As String
    On Error GoTo HeaderUnavailable

    SafeResponseHeader = m_http.getResponseHeader(headerName)
    Exit Function

HeaderUnavailable:
    SafeResponseHeader = vbNullString
End Function

Private Function SafeResponseText() As String
    On Error GoTo ResponseUnavailable

    SafeResponseText = m_http.responseText
    Exit Function

ResponseUnavailable:
    SafeResponseText = "（応答本文を取得できませんでした）"
End Function

Private Function ReadyStateName(ByVal stateValue As Long) As String
    Select Case stateValue
        Case 0
            ReadyStateName = "UNSENT（未初期化）"
        Case 1
            ReadyStateName = "OPENED（接続準備済み）"
        Case 2
            ReadyStateName = "HEADERS_RECEIVED（ヘッダー受信済み）"
        Case 3
            ReadyStateName = "LOADING（本文受信中）"
        Case 4
            ReadyStateName = "DONE（完了）"
        Case Else
            ReadyStateName = "UNKNOWN（不明）"
    End Select
End Function
