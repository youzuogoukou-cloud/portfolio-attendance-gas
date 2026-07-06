# Google Apps Script デプロイ前チェックリスト

クラウド版を公開する前に、このチェックだけは必ず確認してください。

## 1. ローカル確認

プロジェクトフォルダで次を実行します。

```powershell
node validate.mjs
```

成功すると次のように表示されます。

```text
cloud-apps-script validation passed
```

この確認では、Apps Scriptの公開設定が `自分のみ` になっていること、Gmail/Drive/Sheets権限が入っていること、構文エラーがないことを確認します。

## 2. Apps Scriptへコピーするファイル

Apps Scriptプロジェクトへコピーするのは次の3つだけです。

```text
Code.gs
Index.html
appsscript.json
```

`validate.mjs` とこのチェックリストは、ローカル確認用なのでApps Scriptへコピーしません。

## 3. Webアプリのデプロイ設定

デプロイ時は必ず次にしてください。

```text
種類: ウェブアプリ
実行ユーザー: 自分
アクセスできるユーザー: 自分のみ
```

次は選ばないでください。

```text
全員
Anyone
Anyone with Google account
Anyone, even anonymous
```

これらを選ぶと、URLを知っている人がアクセスできる可能性があります。

## 4. テンプレートID

既存テンプレートをGoogle Driveへアップロードし、Google Sheets形式で保存してから、クラウド版画面の初期設定へIDを入力します。

```text
営業用テンプレートID: attendance-template.xlsx をGoogle Sheets化したファイルID
客先用テンプレートID: sales-attendance-template.xlsx をGoogle Sheets化したファイルID
```

IDはGoogle SheetsのURLのこの部分です。

```text
https://docs.google.com/spreadsheets/d/ここがID/edit
```

## 5. 初回の実機確認

1. スマホでWebアプリURLを開く
2. 設定した送信元アカウントでログインする
3. 初期設定にテンプレートIDを保存する
4. テスト日付で勤怠を保存する
5. 月次作成で営業用/客先用ExcelがDriveに作られることを確認する
6. Gmail下書きを作成し、Gmail上で宛先・CC・添付を確認する
7. 実送信は内容確認後にだけ行う
