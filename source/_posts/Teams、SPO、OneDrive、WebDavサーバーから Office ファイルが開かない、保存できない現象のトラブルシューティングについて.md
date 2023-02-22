---
title: Teams、SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象のトラブルシューティングについて
date: '2021-06-21'
lastupdate: '2022-05-30'
id: cl0m4xkk8002oeovs8ggucayv
tags:
  - ファイル I/O
  - クラウド ストレージ

---

こんにちは、Office サポートの西川 (直) です。  

今回の投稿では、Teams、SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象のトラブルシューティングについて説明します。
以下の手順を上から順にご実施ください。

<br>

1\. Office への再サインイン
--
1\. Office からサインアウトし、Office を全て終了します。
2\. OS のスタートメニュー - 歯車 - "アカウント" - "職場または学校にアクセスする" をクリックし、"職場または学校アカウント" として表示されている項目がある場合、"切断" をクリックします
※ Azure AD や AD に接続済みの項目を "切断" しますと、ドメインから外れてしまいますので、Azure AD や AD に接続済み の項目は切断しないようにご注意ください。

![](image1.png)

3\. Office を起動してサインインし、事象が解消するかをお試しください。

<br>

2\.  Office のファイルキャッシュクリア
--
以下の "対処方法" - "方法 2" - "ユーザーインターフェースから削除する場合" をご実施ください。
[SharePoint Server など WebDAV が有効なサーバーから Office ファイルが開かない、保存できない現象について](https://officesupportjp.github.io/blog/cl0m82yjj002mvovs9cil5fgy/index.html)


<br>

3\. Office のファイルキャッシュの完全クリア
--
**事前準備**
・ボリュームライセンス版の Office 2013、2016、または、[サポートされていない古いバージョンの Microsoft 365 Apps](https://learn.microsoft.com/ja-jp/officeupdates/update-history-microsoft365-apps-by-date) の場合、
タスク マネージャーを起動し、\[プロセス\] タブで MSOSYNC.EXE が存在したら "プロセスの終了" をクリックしてください。
・"2. Office のファイルキャッシュクリア"の [ファイルを閉じたときにOfficeドキュメントキャッシュから削除する] もご実施ください
・以下の手順について、Office 2013 の場合は 15.0、Office 2016 以降は 16.0 となります。


- **手順 (UI から実施する場合)**
1. `%localappdata%\Microsoft\Office\16.0` に移動し、OfficeFileCache と冠するフォルダーを全て削除して下さい。
※ OfficeFileCache2 や OfficeFileCache3 等が存在する場合、全て削除してください

2. スタートメニューを右クリックして \[ファイル名を指定して実行\] から regedit と入力し、レジストリ エディタを起動します。
※ ログオンユーザーの権限で起動してください。管理者権限は不要です。
※ UAC プロンプトで他の管理者ユーザーの UAC を入力すると、正しくキャッシュが削除出来ません。

3. 以下のキーを展開し、Server Cache キーを削除します。
`HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Internet\Server Cache`

- **手順 (コマンドで実施する場合)**
1. コマンドプロンプトを起動します。
※ ログオンユーザーの権限で起動してください。管理者権限は不要です。
※ UAC プロンプトで他の管理者ユーザーの UAC を入力すると、正しくキャッシュが削除出来ません。

2. 以下のコマンドを実施します。
```
cd %localappdata%\Microsoft\Office\16.0
for /f %i in ('dir /a:d /b OfficeFileCache*') do rd /s /q %i
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\Internet\Server Cache" /f
```

3. 念のため、`%localappdata%\Microsoft\Office\16.0` に移動し、OfficeFileCache と名前が付いているフォルダーが存在しないことをご確認ください。

<br>

4\. Office のサインインキャッシュのクリア
--
以下の記事の手順により、サインインキャッシュのクリアをご実施ください。
[Office のサインインのトラブルシュートについて](https://officesupportjp.github.io/blog/cl0m6umvg001gi4vsgppfcmk0/index.html)

<br>

**\- 関連情報**
[MSI 版の Office 2016 で SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象について](https://officesupportjp.github.io/blog/cl0m69xu000124cvsclmc5vis/index.html)

<span style="color:#ff0000">**2023/2/22 Update**</span>  
<span style="color:#339966">体裁を整え、対処方法を簡略化しました</span>

<br>
**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**