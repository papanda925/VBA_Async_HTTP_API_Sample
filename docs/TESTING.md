# 実機テスト手順

この文書は、Windows版Excelで実際にコンパイルし、クライアントとサーバーの通信を確認するためのチェックリストです。

作成環境にはWindows版Excel/VBEがないため、リポジトリ公開前または利用開始前に実機テストを行ってください。結果が異なる場合は「想定結果に合わせて無理に操作する」のではなく、Traceとエラー番号を保存してください。

## 0. テスト環境を記録する

最初に次を記録します。

| 項目 | 記入欄 |
|---|---|
| Windowsバージョン | |
| Excelバージョン | |
| Officeのビット数 | 32ビット / 64ビット |
| Microsoft XML 6.0参照 | 有効 / 無効 |
| 実施日 | |

Officeのビット数は、Excelの`ファイル`→`アカウント`→`Excelのバージョン情報`で確認できます。

## 1. コンパイル確認

1. 1冊の`.xlsm`へ5モジュールすべてを取り込みます。
2. `Microsoft XML, v6.0`を参照設定します。
3. VBEの`デバッグ`→`VBAProjectのコンパイル`を実行します。
4. コンパイルエラーが出ないことを確認します。

期待結果: 32ビット版・64ビット版ともにコンパイル完了。

## 2. サーバー起動と停止

1. 保存済みブックから`Server_StartInNewExcel`を実行します。
2. 別Excelプロセスが開き、同じブックが読み取り専用になることを確認します。
3. サーバー側ExcelのTraceシートで`WSA_STARTUP / OK`を確認します。
4. サーバー側ExcelのTraceシートで`LISTEN / READY`を確認します。
5. 元のブックから`Server_StopInNewExcel`を実行します。
6. サーバー側Excelプロセスが終了することを確認します。
7. 再び`Server_StartInNewExcel`を実行できることを確認します。

期待結果: 2回目もWinsockエラー10048にならず起動できる。

自動起動が組織ポリシーで許可されない場合は、`excel.exe /x`で別Excelを手動起動し、読み取り専用コピー側で`Server_Start`を実行します。この場合もセキュリティ設定を回避しないでください。

### 手動のServer_Start/Stop確認

サーバー側の読み取り専用コピーで、次も個別に確認します。

1. `Server_Start`を実行します。
2. Traceシートで`WSA_STARTUP / OK`を確認します。
3. Traceシートで`LISTEN / READY`を確認します。
4. `Server_Stop`を実行します。
5. Traceシートで`STOP / STOPPED`を確認します。
6. 再び`Server_Start`を実行できることを確認します。

期待結果: 2回目もWinsockエラー10048にならず起動できる。

## 3. `/hello`正常系

1. `Server_StartInNewExcel`でサーバーを起動します。
2. 元のブックで`Client_Hello`を実行します。
3. クライアントTraceをTraceIdで絞り込みます。
4. サーバーTraceで同じTraceIdを探します。

期待結果:

- サーバー: `ACCEPT`→`RECEIVE`→`PARSE`→`RESPONSE`
- クライアント: `READY_STATE`の最後が`4 = DONE`
- クライアント: `TRACE_ID / MATCH`
- クライアント: `RESPONSE / HTTP 200`
- 応答本文: `{"message":"Hello from VBA HTTP server"}`
- 両側のTraceIdが一致

## 4. `/echo`本文往復

1. `Client_Echo`を実行します。
2. クライアントの`SEND`行で本文文字数を確認します。
3. クライアントの`RESPONSE`行で、送信したJSONが返ったことを確認します。

期待結果:

```json
{"message":"Hello from Excel VBA","number":123}
```

追加確認として、`Client_Echo`の本文を日本語へ変更します。

```vb
"{""message"":""こんにちは""}"
```

期待結果: 日本語が文字化けせず往復する。文字化けした場合は、ソース取込み時の文字コードと`Content-Type`を記録する。

## 5. `/delay`非同期確認

1. `Client_Delay3Seconds`を実行します。
2. 実行直後、クライアント側Excelで別セルを選択します。
3. 約3～5秒後に`HTTP 200`となることを確認します。
4. サーバーTraceの`DELAY / WAITING`を確認します。

期待結果: サーバー用ExcelはSleep中に一時停止するが、別プロセスのクライアント用Excelは操作できる。

ポーリング間隔があるため、経過時間は指定値ちょうどにはなりません。性能試験ではなく、非同期の動作確認として評価します。

## 6. HTTP 500確認

1. `Client_Status500`を実行します。
2. readyStateが4まで進むことを確認します。
3. `RESPONSE / HTTP 500`を確認します。
4. `COMPLETE / HTTP_ERROR`を確認します。

期待結果: VBAの実行時エラーで停止せず、HTTPエラーとしてTraceへ記録される。

## 7. 404確認

`modClientDemo`へ一時的に次のテスト用Subを追加し、実行します。

```vb
Public Sub Client_NotFoundForTest()
    ' Client_SendRequestはPrivateなので、テスト時だけ同じモジュール内へ追加します。
    Client_SendRequest "GET", "/not-found"
End Sub
```

期待結果: `HTTP 404`、`HTTP_ERROR`となる。

## 8. 複数要求のTraceId

1. `Client_RunBasicSequence`を実行します。
2. TraceIdが3つ作成されることを確認します。
3. それぞれの要求が別のTraceIdで完了することを確認します。

注意: サーバーは1接続ずつ処理する教材実装です。複数要求は順次処理され、実運用Webサーバーの同時処理性能を再現しません。

## 9. 接続失敗

1. `Server_Stop`でサーバーを停止します。
2. `Client_Hello`を実行します。

期待結果:

- Excelが異常終了しない
- クライアントTraceに`Status=0`相当またはイベントエラーが記録される
- 完了済み要求がCollectionへ残り続けない

MSXMLのエラー通知内容はWindows/Officeの版で異なる可能性があります。実際のTraceを結果として保存してください。

## 10. キャンセル

1. サーバーを起動します。
2. `Client_Delay3Seconds`を実行します。
3. すぐに`Client_CancelAll`を実行します。

期待結果: `CANCEL / CANCELLED`が記録され、Excelが停止しない。

## 11. `/shutdown`

1. サーバーを起動します。
2. `Client_ShutdownServer`を実行します。
3. クライアントがHTTP 200を受け取ることを確認します。
4. サーバーTraceで`STOP / STOPPED`を確認します。
5. その後の`Client_Hello`が接続失敗になることを確認します。
6. サーバー側Traceを確認後、元のブックで`Server_StopInNewExcel`を実行します。

期待結果: 応答を送信してからサーバーが停止する。

## 12. 入力上限と安全対策

次を静的に確認します。

- `MAX_REQUEST_BYTES = 65536`
- `MAX_DELAY_MILLISECONDS = 10000`
- bind先が`127.0.0.1`
- `#If Win64 Then`でWSADATAの64/32ビット配置を分離
- Traceに`Authorization:`が渡されたとき伏せ字になる
- `Trace_DETAIL_LIMIT`相当が1000文字
- TraceのA:Gが文字列形式、Hが数値形式
- Traceが`Value2`と`Debug.Print`の両方へ出力される

マスキング確認は、イミディエイトウィンドウから次を実行できます。

```vb
Trace_Write "TEST", "CLIENT", "MASK", "LOCAL", _
    "Authorization: Bearer secret-value", "CHECK"
```

期待結果: 詳細列が`Authorization: ***REDACTED***`となる。

## 13. ブック終了時の確認

サーバーを起動したままブックを閉じるテストは、保存確認やOnTime予約の影響を受けます。安全のため、通常は先に`Server_Stop`を実行してください。

より確実にする場合は、Server.xlsmの`ThisWorkbook`へ次を追加できます。

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Server_Stop
End Sub
```

このイベントコードは利用者のブック構成に関わるため、配布ソースへ自動挿入していません。

## 結果の報告に必要な情報

不具合を報告するときは、次を添えてください。

- 再現したテスト番号
- Officeの32/64ビット
- ExcelとWindowsのバージョン
- VBEで選択中の参照設定
- エラー番号とメッセージ
- クライアントとサーバー双方のTrace（秘密情報を削除）
- 毎回発生するか、時々発生するか
