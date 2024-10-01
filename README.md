# MailSenderApp

## 概要
MailSenderAppは、ユーザーが使いやすいグラフィカルインターフェースを使用して複数の受信者にメールを送信できるPythonベースのアプリケーションです。このアプリケーションは、認証のために送信者のメールアドレスとパスワードを必要とし、CSVファイルから受信者のメールアドレスを読み込むことができます。

## 特徴
- 簡単なメール作成と送信のためのユーザーフレンドリーなGUI。
- 件名とメッセージのための入力フィールド。
- 最大5つまでの個別のメールアドレスを追加する機能。
- CSVファイルから受信者のメールアドレスを読み込む機能。
- 送信者のメール資格情報を安全に保存。
- 送信されたメールやエラーのログ記録。
- GUIがフリーズしないようにするためのマルチスレッドメール送信。

## 始め方
### 前提条件
- Python 3.x
- 必要なライブラリ：`pandas`、`tkinter`、`smtplib`、`email`、`logging`

### インストール
1. リポジトリをクローンまたはダウンロードします。
2. pipを使用して必要なライブラリをインストールします。
   ```bash
   pip install pandas
### 初回起動時
アプリを起動すると、メールアドレスとGoogleアプリパスワードを入力するセットアップ画面が表示されます。

### メイン画面の設定
- 件名: メールの件名を入力します。
- 本文: メールの内容を入力します。
- 個別メールアドレスの入力: 最大5つの個別メールアドレスを追加できます。
- CSVファイルの読み込み: 「受信者CSVを読み込む」ボタンをクリックして、宛先リストをCSVファイルから取得します。
- メール送信: 「送信」ボタンを押すと、指定された宛先にメールが送信されます。


## CSV File Format

The application requires a CSV file containing the email addresses of the recipients. The file should have the following format:

```csv
email
example1@example.com
example2@example.com


重要な注意事項
Gmailアカウントの設定: メールを送信するには、GmailアカウントとGoogleアプリパスワードが必要です。
アプリパスワードの生成: Googleアカウントのセキュリティ設定でアプリパスワードを生成してください。詳細はこちらを参照してください。
ロギング
アプリケーションは、メール送信プロセスを email_log.txt というファイルに記録します。このファイルには送信されたメールの詳細とエラーメッセージが含まれます。





