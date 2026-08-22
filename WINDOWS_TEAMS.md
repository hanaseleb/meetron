# Windows 11とTeams Webの開発者向け検証

Meetronには、Windows 11上でChromeまたはEdgeを専用プロファイルとして起動し、Microsoft Teams Webの参加前画面を操作するコードがあります。
自動テストは通過していますが、Windows実機では検証中です。
通常利用のブラウザープロファイルは使わないため、既存のタブ、拡張、ログイン状態を分離できます。

現時点では、仮想音声デバイス、Native Messaging Host、ChatGPT Voiceとの統合起動は未対応です。
以下のコマンドはブラウザー制御の検証だけを対象とします。

## 前提

- Windows 11
- Node.js 22または24 LTS
- Google ChromeまたはMicrosoft Edge
- `npm install`を実行済みのMeetronリポジトリ
- 機密情報を含まないTeamsテスト会議URL

PowerShellを開き、Meetronリポジトリへ移動してから実行します。

```powershell
Set-Location 'C:\path\to\meetron'
```

このコマンドは、後続のコマンドがMeetronのファイルを参照できるように作業場所を変更します。
システム設定やファイル内容は変更しません。

次に、Teams会議URLをPowerShellの一時変数へ入力します。

```powershell
$TeamsMeetingUrl = Read-Host 'Teams会議URLを貼り付けてください'
```

このコマンドは、入力したURLを現在のPowerShellセッション内だけに保持します。
ファイルやWindowsの設定には保存しません。

## ブラウザー起動内容の確認

最初に`--dry-run`を付け、Edgeの検出結果と起動引数を確認します。

```powershell
node .\scripts\open-gpt-participant.mjs `
  --browser edge `
  --dry-run `
  $TeamsMeetingUrl
```

このコマンドはブラウザーを起動せず、実行予定の内容だけを表示します。
出力には、Edgeの実行ファイル、専用プロファイルの保存先、`127.0.0.1`だけで待ち受ける自動操作ポートが含まれます。

Chromeで検証する場合は、`--browser edge`を`--browser chrome`へ変更します。
EdgeとChromeはどちらもChromium系ですが、実行ファイルの場所と企業ポリシーは別に判定されるため、両方で同じ結果になるとは限りません。

## Teams Webを専用ブラウザーで開く

`--dry-run`を外すと、Teams Webを専用プロファイルで開きます。

```powershell
node .\scripts\open-gpt-participant.mjs `
  --browser edge `
  $TeamsMeetingUrl
```

このコマンドは`$env:LOCALAPPDATA\Meetron\GPTParticipantChromium`を作成し、Edgeを起動します。
Windowsの既定ブラウザー、通常のEdgeプロファイル、音声設定は変更しません。
カメラとマイクの権限はTeamsの画面または自動準備処理から付与し、Edgeが非対応として警告するメディア自動許可フラグは使用しません。

成功すると、PowerShellに`Dedicated meeting browser is ready.`と表示されます。
初回は専用ウィンドウでTeamsへのサインインやCookie同意が必要になる場合があります。

## 参加前画面を自動設定する

Teamsに表示される音声デバイス名を確認し、マイクとスピーカーを明示して実行します。

```powershell
node .\scripts\open-gpt-participant.mjs `
  --browser edge `
  --auto-prepare `
  --name 'GPT-Live' `
  --microphone-device 'VIRTUAL_MIC_NAME' `
  --speaker-device 'VIRTUAL_SPEAKER_NAME' `
  $TeamsMeetingUrl
```

このコマンドはTeamsのブラウザー継続、表示名入力、音声デバイス選択、マイクのミュート、カメラの停止を行います。
`--auto-prepare`では参加ボタンを押しません。

参加リクエストまで検証する場合だけ、ホストと参加者へ事前に知らせたテスト会議で`--auto-prepare`を`--join`へ変更します。
`--join`は実際に参加ボタンを押すため、本番会議では使用しないでください。

## 結果として共有する情報

問題が起きた場合は、次の情報を共有すると原因を切り分けられます。

- 実行したコマンド（会議URLの識別部分は伏せる）
- PowerShellに表示されたエラー全文
- ChromeとEdgeのどちらを使ったか
- Teamsの画面が止まった箇所
- Teamsに表示されたマイクとスピーカーの名前

パスワード、認証コード、会議URLの識別部分は共有しないでください。
