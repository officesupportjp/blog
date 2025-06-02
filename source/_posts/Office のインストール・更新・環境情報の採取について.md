---
title: Office のインストール・更新・環境情報の採取について
date: '2024-11-28'
lastupdate: '2025-06-02'
id: cm3e5mpvi0000cgkj2sfvb55j
tags:
  - 更新
  - インストール
---

こんにちは、Office サポートの西川 (直) です。 
本記事では Office のインストール・更新・環境情報の採取について解説します。


環境情報ログ
---
1. [MSOfficeinfo.zip](MSOfficeinfo_v3.2.zip) をダウンロードします。

2. MSOfficeinfo.zip のプロパティを開き、以下で "許可する"にチェックをつけます。
![](image1.png)
3. MSOfficeinfo.zip を解凍し、MSOfficeinfo.txt の拡張子を .cmd に変更します。

4. MSOfficeinfo.cmd をダブルクリックで実行します。結果は、C:\MS_DATA 配下に出力されます。
※ 管理者を持つ別のユーザーで実行すると正しく取得できませんので、ログオンユーザーでダブルクリックで取得ください。

5. 取得した以下の情報を弊社にお寄せください
・C:\MS_DATA に出力された MOI から始まる zip ファイル

<br>

インストール、更新ログ
---
0. 事前に環境情報ログを採取してください

1. [Microsoft 365 Apps Deployment Log Collector](https://www.microsoft.com/en-us/download/details.aspx?id=103236) から **officedeploymentlogcollector.exe** をダウンロードします。

2. **officedeploymentlogcollector.exe** をダブルクリックし、任意のフォルダに展開します。

3. 展開したフォルダに配置された **OfficeC2RLogCollector.exe** を実行します。(端末の管理者権限が必要となります)

4. C:\Users\\<ユーザー名>\AppData\Local\Temp\C2RLogs 配下に出力された c2rlogs_ から始まる zip ファイルを取得ください

5. 取得した以下の情報を弊社にお寄せください
・環境情報ログ(C:\MS_DATA に出力された MOI から始まる zip ファイル)
・c2rlogs_ から始まる zip ファイル
・インストール・更新の事象の場合、事象発生日時
・インストールの事象の場合、利用した Configuration.xml ファイル
・エラーが生じる事象の場合、スクリーンショット

<br>

ライセンス認証ログ
---
0. 事前に環境情報ログを採取してください

1. [Office Licensing Diagnostic Tool](https://www.microsoft.com/en-us/download/details.aspx?id=55948) から **OfficeLicenseDiagnostic.zip** をダウンロードし展開してください。
2. **licenseInfo.cmd** をダブルクリックで実行してください。
※ 途中、数回 Enter キーの入力が求められます。また、結果は %temp% に licensingInfo から始まる名前で出力されます。
※ 管理者を持つ別のユーザーで実行すると正しく取得できませんので、ログオンユーザーでダブルクリックで取得ください。

3. 取得した以下の情報を弊社にお寄せください
・環境情報ログ(C:\MS_DATA に出力された MOI から始まる zip ファイル)
・licensingInfo から始まる zip ファイル

<br>

<span style="color:#ff0000">**2024/11/28  Update**</span>  
<span style="color:#339966">一部記載を変更しました</span>

<span style="color:#ff0000">**2024/1/29  Update**</span>  
<span style="color:#339966">スクリプトを更新しました</span>

<span style="color:#ff0000">**2024/2/5  Update**</span>  
<span style="color:#339966">採取方法を簡素化しました</span>

<span style="color:#ff0000">**2024/6/2  Update**</span>  
<span style="color:#339966">スクリプトを更新しました</span>

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**