# outlook-classic-mcp

日本語 | [English](README.en.md)

**Windows の旧式 Outlook（デスクトップ版）** のメールを、Claude から読む・探す・下書きする・送信するための [Model Context Protocol](https://modelcontextprotocol.io)（MCP）サーバーです。PC に入っている Outlook をそのまま使うので、Outlook に設定済みのアカウント（Exchange・IMAP・POP3）ならどれでも使えます。

送信するときは、必ず画面に確認ウィンドウが出ます。自分で **「送信する」** を押さない限り、メールは送信されません。

[![ダウンロード](https://img.shields.io/badge/%E3%83%80%E3%82%A6%E3%83%B3%E3%83%AD%E3%83%BC%E3%83%89-outlook--classic--mcp.mcpb-2563eb?style=for-the-badge)](https://github.com/nicosphere-group/outlook-classic-mcp/releases/latest/download/outlook-classic-mcp.mcpb)

1. 上のボタンから `outlook-classic-mcp.mcpb` をダウンロードします（zip などに展開する必要はありません）。
2. Claude デスクトップアプリを開き、「設定」→「拡張機能」→「詳細設定」→「拡張機能をインストール」で、そのファイルを選びます。
3. 「インストール」を押します。旧式 Outlook を起動してから、Claude に「今日届いたメールを教えて」と頼んでみてください。

## 特長

- **いまの Outlook をそのまま使う**：メールのパスワードやサーバーの設定、クラウドへのサインインは要りません。
- **ほかに入れるものがない**：必要なものを 1 つにまとめ、Claude デスクトップアプリの拡張機能として配ります。
- **送信には自分のクリックが必要**：確認ウィンドウのボタンを Claude が押すことはできません。
- **PC の外に出ない**：外部との通信をせず、利用状況も集めません。

## ツール一覧

| ツール | 内容 |
|---|---|
| `list_folders` | メールフォルダの一覧と、件数・未読数 |
| `list_accounts` | アカウントとデータファイルの一覧 |
| `list_emails` | フォルダ内のメール一覧（期間・未読・件数で絞り込み） |
| `search_emails` | 件名・差出人・本文・期間で検索 |
| `read_email` | メールの本文・宛先・添付ファイル名を読む |
| `list_attachments` | メールの添付ファイルの一覧 |
| `save_attachment` | 添付ファイルを決められたフォルダに保存 |
| `create_draft` | 新規メールの下書きを作る |
| `create_reply_draft` | 返信・全員に返信の下書きを作る |
| `create_forward_draft` | 転送の下書きを作る |
| `update_draft` | 下書きを直す |
| `send_draft` | 画面で確認したうえで、下書きを送信する |

メールの削除・フォルダ移動と、予定表の操作は、第 1 版には含めません。

## 送信するときの動き

送信できるのは、作成済みの下書きだけです。新規メールをいきなり送る機能はありません。

1. Claude に「送って」と頼みます。
2. 画面に確認ウィンドウが出ます。差出人・宛先・CC・BCC・件名・添付ファイル・本文の全文が表示されます。
3. 3 つのボタンから選びます。

| ボタン | 押すとどうなるか |
|---|---|
| **送信する** | メールを送信します。 |
| **Outlook で開いて確認** | 下書きを Outlook で開きます。送信はしません。 |
| **送信しない** | 何も送信しません。下書きはそのまま残ります。 |

- ウィンドウを閉じたとき、何もせず 3 分たったときも、送信されません。
- 宛先が多い、全員に返信、別の組織のアドレスが含まれる、添付ファイルがある場合は、上部に注意が表示されます。
- 確認ウィンドウは毎回出ます。「今後は確認しない」の選択肢や、確認を省く設定はありません。
- ウィンドウを開いている間に下書きが変わった場合は、送信されません。
- 送信したときと、送信しなかったときの記録が `%LOCALAPPDATA%\outlook-classic-mcp\audit.log` に残ります。記録するのは日時・差出人アカウント・宛先・件名で、本文は含みません。

## 必要なもの

- Windows 10 または 11
- 旧式 Outlook 2016 以降（メールアカウントを設定済みであること）
- Claude デスクトップアプリ（Windows）、または Claude Code

新しい Outlook（タイトルバーに「新しい Outlook」と出るもの）と、ブラウザ版の Outlook では使えません。ブラウザ版・スマホ版の Claude でも使えません。Outlook と同じ PC の中で動かす必要があるためです。

会社の Claude（Team / Enterprise）では、拡張機能を入れるのに管理者の許可が必要な場合があります。

## インストール

### Claude デスクトップアプリ

1. [`outlook-classic-mcp.mcpb`](https://github.com/nicosphere-group/outlook-classic-mcp/releases/latest/download/outlook-classic-mcp.mcpb) をダウンロードします。最新版が落ちてきます。過去の版は[リリースのページ](https://github.com/nicosphere-group/outlook-classic-mcp/releases)にあります。
2. Claude デスクトップアプリの「設定」→「拡張機能」→「詳細設定」→「拡張機能をインストール」で、そのファイルを選びます。

### Claude Code

ソースからビルドしたあと、次のように登録します。

```
claude mcp add outlook-classic -- C:\path\to\outlook-classic-mcp.exe
```

## 使い方

旧式 Outlook を起動してから、Claude に普通の言葉で頼みます。

- 「今日届いたメールを教えて」
- 「田中さんからの見積りのメールを探して」
- 「このメールに、来週水曜なら伺えますと返信の下書きを作って」
- 「もう少し丁寧な言い方に直して」
- 「送って」（確認ウィンドウが出ます）

## 設定

設定は不要です。どの設定にしても、送信前の確認ウィンドウは必ず出ます。

Claude デスクトップアプリでは、拡張機能の設定画面で「送信を許可するか」と「添付ファイルの保存先」を変えられます。Claude Code などでは、次の環境変数で設定します。

| 環境変数 | 既定 | 内容 |
|---|---|---|
| `OUTLOOK_MCP_ALLOW_SEND` | `1` | `0` にすると送信のツールがなくなります。読み取りと下書きは使えます。 |
| `OUTLOOK_MCP_SAVE_DIR` | `ダウンロード\outlook-classic-mcp` | 添付ファイルの保存先。 |
| `OUTLOOK_MCP_ATTACH_DIRS` | ドキュメント・デスクトップ・ダウンロード | 下書きに添付できるファイルの置き場所（`;` 区切り）。 |
| `OUTLOOK_MCP_CONFIRM_TIMEOUT` | `180` | 確認ウィンドウを自動で閉じるまでの秒数（15〜600）。 |
| `OUTLOOK_MCP_LANG` | 自動 | 確認ウィンドウの言語（`ja` / `en`）。 |

## セキュリティとプライバシー

- 外部とは通信しません。同じ PC の Outlook とだけやり取りします。
- メールの内容が Claude に渡るのは、会話の中でツールが返したときだけです。
- メールの内容は「他人が書いたもの」として扱います。メールの中に書かれた指示には従わないよう、Claude に伝えています。
- 管理者権限は要りません。Outlook か Claude を「管理者として実行」していると接続できません。
- このツールのために、Outlook のセキュリティ警告をレジストリで無効にしないでください。
- Claude Code では、Claude がコマンドも実行できるため、Outlook に届く道が確認ウィンドウだけではありません。コマンド実行の許可確認は有効のままにしてください。

## 困ったとき

**「Outlook did not respond」と出る**
Outlook を普通に起動してください。Outlook か Claude を「管理者として実行」していると接続できません。

**「Classic Outlook was not found」と出る**
新しい Outlook だけが入っている PC では使えません。旧式 Outlook が必要です。

**Outlook が「プログラムがアクセスしようとしています」という警告を出す**
ウイルス対策ソフトが無効、または古いときに Outlook が出す警告です。有効にして最新の状態にしてください。

**送信したのに相手に届かない（POP3 / IMAP）**
メールは Outlook の送信トレイに入り、次の送受信で送られます。Outlook を開いたままにしてください。

## ビルド

必要なもの：.NET 10 SDK と、拡張機能ファイルを作るための Node.js

```
./build.ps1
```

`dist\outlook-classic-mcp.mcpb` と `extension\server\outlook-classic-mcp.exe` ができます。

- [skills/outlook-classic/SKILL.md](skills/outlook-classic/SKILL.md)：ツールの安全な使い方を Claude に教えるスキル

## ライセンス

[MIT](LICENSE)
