Attribute VB_Name = "modTrace"
Option Explicit

' =============================================================================
' modTrace
' クライアントとサーバーで共通利用する、学習用のトレース出力モジュールです。
'
' このサンプルでは、単に「成功した／失敗した」だけを表示しません。
' 1つのHTTP要求に同じTraceIdを付け、要求の開始から応答完了までを追跡します。
' クライアント用ブックとサーバー用ブックのTraceシートをTraceIdで照合すると、
' 通信の両側で何が起きたかを時系列で確認できます。
' =============================================================================

Private Const TRACE_SHEET_NAME As String = "Trace"
Private Const TRACE_DETAIL_LIMIT As Long = 1000

Private m_traceSequence As Long

' 新しい要求を識別するTraceIdを作ります。
' 時刻だけでは同じミリ秒に複数の要求を開始した場合に重複する可能性があるため、
' ブック内の連番も付加しています。
Public Function Trace_NewId(Optional ByVal prefix As String = "REQ") As String
    m_traceSequence = m_traceSequence + 1

    Trace_NewId = prefix & "-" _
        & Format$(Now, "yyyymmdd-hhnnss") & "-" _
        & Format$(CLng(Int((Timer - Fix(Timer)) * 1000#)), "000") & "-" _
        & Format$(m_traceSequence, "0000")
End Function

' Timer関数の値を記録します。
' 経過時間の測定開始時に呼び出し、戻り値をTrace_Writeへ渡してください。
Public Function Trace_StartTimer() As Double
    Trace_StartTimer = Timer
End Function

' Timer関数は午前0時に0へ戻ります。
' そのため、日付をまたいだ場合は1日分の秒数を加えて補正します。
Public Function Trace_ElapsedMilliseconds(ByVal startedAt As Double) As Long
    Dim currentValue As Double
    Dim elapsedSeconds As Double

    currentValue = Timer
    elapsedSeconds = currentValue - startedAt

    If elapsedSeconds < 0# Then
        elapsedSeconds = elapsedSeconds + 86400#
    End If

    Trace_ElapsedMilliseconds = CLng(elapsedSeconds * 1000#)
End Function

' Traceシートへ1行追記します。
'
' 引数の意味:
'   traceId  : 1つの要求を両側で結び付ける識別子
'   side     : CLIENT または SERVER
'   stepName : OPEN、SEND、RECEIVE、PARSE、COMPLETEなどの処理段階
'   direction: LOCAL、OUT、INのいずれか
'   detail   : 実行内容。秘密情報はTrace_Sanitizeで伏せ字にします
'   result   : OK、WAITING、HTTP 200、ERRORなどの結果
'   startedAt: Trace_StartTimerの戻り値。0の場合は経過時間を空欄にします
Public Sub Trace_Write( _
    ByVal traceId As String, _
    ByVal side As String, _
    ByVal stepName As String, _
    ByVal direction As String, _
    ByVal detail As String, _
    ByVal result As String, _
    Optional ByVal startedAt As Double = 0#)

    Dim ws As Worksheet
    Dim nextRow As Long
    Dim elapsedValue As Variant
    Dim timestampValue As String
    Dim safeDetail As String

    If startedAt > 0# Then
        elapsedValue = Trace_ElapsedMilliseconds(startedAt)
    Else
        elapsedValue = Empty
    End If

    timestampValue = Trace_Timestamp()
    safeDetail = Trace_Sanitize(detail)

    ' シートの取得・作成より先に出力します。ブック構造やシートが保護されて
    ' いても、通信イベントそのものはイミディエイトウィンドウへ残ります。
    Debug.Print timestampValue & vbTab _
        & traceId & vbTab _
        & side & vbTab _
        & stepName & vbTab _
        & direction & vbTab _
        & result & vbTab _
        & safeDetail

    On Error GoTo TraceError

    Set ws = Trace_GetWorksheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Value2は、日付や通貨への自動変換を避け、与えた値をそのまま記録しやすい
    ' プロパティです。A:Gは文字列、Hだけを数値として扱います。
    ' 既存のTraceシートを再利用する場合にも崩れないよう、追記行にも形式を設定します。
    ws.Range("A" & CStr(nextRow) & ":G" & CStr(nextRow)).NumberFormat = "@"
    ws.Cells(nextRow, 8).NumberFormat = "0"
    ws.Cells(nextRow, 1).Value2 = timestampValue
    ws.Cells(nextRow, 2).Value2 = traceId
    ws.Cells(nextRow, 3).Value2 = side
    ws.Cells(nextRow, 4).Value2 = stepName
    ws.Cells(nextRow, 5).Value2 = direction
    ws.Cells(nextRow, 6).Value2 = safeDetail
    ws.Cells(nextRow, 7).Value2 = result
    ws.Cells(nextRow, 8).Value2 = elapsedValue

    Exit Sub

TraceError:
    ' トレース出力の失敗によって本来の通信処理まで停止させないため、
    ' ここではエラーを再送出せず、イミディエイトウィンドウへ退避します。
    Debug.Print "Trace_Write failed: " & Err.Number & " / " & Err.Description
End Sub

' Traceシートの記録を消去します。
' 見出しは再作成されるため、そのまま次のテストを開始できます。
Public Sub Trace_Clear()
    Dim ws As Worksheet

    Set ws = Trace_GetWorksheet()
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    Trace_InitializeWorksheet ws
End Sub

' Traceシートがなければ作成し、見出しと表示形式を設定します。
Private Function Trace_GetWorksheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TRACE_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = TRACE_SHEET_NAME
        Trace_InitializeWorksheet ws
    ElseIf Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        Trace_InitializeWorksheet ws
    End If

    Set Trace_GetWorksheet = ws
End Function

Private Sub Trace_InitializeWorksheet(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim columnIndex As Long

    headers = Array( _
        "日時", "TraceId", "側", "処理段階", _
        "方向", "詳細", "結果", "経過時間(ms)")

    For columnIndex = LBound(headers) To UBound(headers)
        ws.Cells(1, columnIndex + 1).Value2 = headers(columnIndex)
    Next columnIndex

    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        If Not ws.AutoFilterMode Then
            .AutoFilter
        End If
    End With

    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C:E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 70
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("F").WrapText = True

    ' TraceIdや日時をExcelが日付・指数などへ自動変換しないよう、
    ' A:Gを文字列形式に固定します。経過時間のH列だけは数値形式です。
    ws.Columns("A:G").NumberFormat = "@"
    ws.Columns("H").NumberFormat = "0"
End Sub

' Authorizationヘッダーなどを、そのまま記録しないための安全対策です。
' この教材ではAPIキーを使いませんが、別のAPIへ応用したときの事故を防ぐため、
' 最初からトレース共通処理にマスキングを入れています。
Private Function Trace_Sanitize(ByVal value As String) As String
    Dim sanitized As String
    Dim lines As Variant
    Dim lineIndex As Long

    sanitized = value
    lines = Split(sanitized, vbCrLf)

    For lineIndex = LBound(lines) To UBound(lines)
        If InStr(1, lines(lineIndex), "Authorization:", vbTextCompare) > 0 Then
            lines(lineIndex) = "Authorization: ***REDACTED***"
        ElseIf InStr(1, lines(lineIndex), "Api-Key:", vbTextCompare) > 0 Then
            lines(lineIndex) = "Api-Key: ***REDACTED***"
        End If
    Next lineIndex

    sanitized = Join(lines, vbCrLf)

    If Len(sanitized) > TRACE_DETAIL_LIMIT Then
        sanitized = Left$(sanitized, TRACE_DETAIL_LIMIT) _
            & " ...（長いため省略）"
    End If

    Trace_Sanitize = sanitized
End Function

Private Function Trace_Timestamp() As String
    Trace_Timestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") _
        & "." & Format$( _
            CLng(Int((Timer - Fix(Timer)) * 1000#)), "000")
End Function
