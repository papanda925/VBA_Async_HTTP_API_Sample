Attribute VB_Name = "modClientDemo"
Option Explicit

' =============================================================================
' modClientDemo
' CAsyncHttpClientを使うための、実行しやすい入口をまとめたモジュールです。
'
' 非同期通信では、SendAsyncを呼び出したSubが終了した後もHTTP処理が続きます。
' そのためCAsyncHttpClientをモジュール変数のCollectionに保持し、応答が来る前に
' オブジェクトが破棄されないようにしています。
' =============================================================================

Private Const SERVER_BASE_URL As String = "http://127.0.0.1:18080"

Private m_activeRequests As Collection

' GET /hello
' 固定のJSON応答を受信する、最初に試すサンプルです。
Public Sub Client_Hello()
    Client_SendRequest "GET", "/hello"
End Sub

' POST /echo
' 送信したJSON本文が、そのまま応答本文として返ります。
Public Sub Client_Echo()
    Client_SendRequest "POST", "/echo", _
        "{""message"":""Hello from Excel VBA"",""number"":123}"
End Sub

' GET /delay?ms=3000
' サーバーが約3秒待ってから応答します。
' SendAsyncの直後に制御がExcelへ戻るため、待機中にも別セルを選択できることを
' 確認すると、同期通信との違いを体験できます。
Public Sub Client_Delay3Seconds()
    Client_SendRequest "GET", "/delay?ms=3000"
End Sub

' GET /status/500
' HTTP 500を返すテストです。通信そのものは完了していても、HTTPの処理結果が
' 成功とは限らないことをTraceシートで確認できます。
Public Sub Client_Status500()
    Client_SendRequest "GET", "/status/500"
End Sub

' POST /shutdown
' サーバーが応答を返した後、安全に待受けを終了します。
Public Sub Client_ShutdownServer()
    Client_SendRequest "POST", "/shutdown"
End Sub

' 複数の要求を続けて開始し、TraceIdで区別できることを確認します。
Public Sub Client_RunBasicSequence()
    Client_Hello
    Client_Echo
    Client_Status500
End Sub

' 実行中のすべての要求を中止します。
Public Sub Client_CancelAll()
    Dim requestIndex As Long
    Dim request As CAsyncHttpClient

    If m_activeRequests Is Nothing Then Exit Sub

    For requestIndex = m_activeRequests.Count To 1 Step -1
        Set request = m_activeRequests(requestIndex)
        request.Cancel
        m_activeRequests.Remove requestIndex
    Next requestIndex
End Sub

' CAsyncHttpClientから完了通知を受け、Collectionを整理します。
' Publicにしているのは、クラスモジュールから呼び出せるようにするためです。
Public Sub Client_RequestCompleted(ByVal completedTraceId As String)
    Dim requestIndex As Long
    Dim request As CAsyncHttpClient

    If m_activeRequests Is Nothing Then Exit Sub

    For requestIndex = m_activeRequests.Count To 1 Step -1
        Set request = m_activeRequests(requestIndex)

        If request.TraceId = completedTraceId Then
            m_activeRequests.Remove requestIndex
            Exit For
        End If
    Next requestIndex
End Sub

Private Sub Client_SendRequest( _
    ByVal requestMethod As String, _
    ByVal relativePath As String, _
    Optional ByVal requestBody As String = vbNullString)

    Dim request As CAsyncHttpClient

    If m_activeRequests Is Nothing Then
        Set m_activeRequests = New Collection
    End If

    Set request = New CAsyncHttpClient

    ' SendAsyncより先にCollectionへ追加します。
    ' 非常に短時間で失敗通知が来た場合にも、完了処理で正しく見つけられます。
    m_activeRequests.Add request
    request.SendAsync requestMethod, SERVER_BASE_URL & relativePath, requestBody
End Sub
