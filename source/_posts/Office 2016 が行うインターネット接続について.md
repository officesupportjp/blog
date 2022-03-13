---
title: Office 2016 が行うインターネット接続について
date: 2019-02-07
id: cl0n10de0001a90vs28153y7h
alias: /Office 2016 が行うインターネット接続について/
---

(※ 2017 年 2 月 8 日に Office Support Team Blog JAPAN に公開した情報のアーカイブです。)

  
  

こんにちは、Office サポートです。  
本記事では Office 2016 が行うインターネット接続について説明します。

  
  

**説明**
======

Office 2016 では起動時や一部の機能でインターネットへ接続する動作になっています。

この動作は想定された動作です。

具体的な接続先は、以下の資料で案内しています。

  

 Office 365 ProPlus でのネットワーク要求

 [https://support.office.com/ja-jp/article/Office-365-ProPlus-%E3%81%A7%E3%81%AE%E3%83%8D%E3%83%83%E3%83%88%E3%83%AF%E3%83%BC%E3%82%AF%E8%A6%81%E6%B1%82-eb73fcd1-ca88-4d02-a74b-2dd3a9f3364d](https://support.office.com/ja-jp/article/Office-365-ProPlus-%E3%81%A7%E3%81%AE%E3%83%8D%E3%83%83%E3%83%88%E3%83%AF%E3%83%BC%E3%82%AF%E8%A6%81%E6%B1%82-eb73fcd1-ca88-4d02-a74b-2dd3a9f3364d)

  

上記資料は、Office 365 ProPlus (クイック実行版) を主とした説明資料ですが、\[展開してOffice 365 ProPlus のFQDN を表示する\] の項番1 と9 以外は、Windows インストーラー版のOffice 2016 においても同様です。

  
  
  

**参考情報**
========

インターネットが接続できる環境では、以下の機能を使うことができます。

・ヘルプ  
・オンライン テンプレート  
・\[挿入\] タブにある\[オンライン画像\]、\[オンライン ビデオ\]、\[オンライン オーディオ\]  
・\[挿入\] タブにある\[ストア\]  
・\[校閲\] タブの\[翻訳\] で、Bing 等のオンラインのサービスでの検索

  

上記のようなオンラインを前提とした機能があるため、インターネットへ接続できる環境での利用が理想的ではありますが、中にはご運用上、インターネット接続が禁止されている環境もあることを認識しています。

  

Office 2016 が行うインターネット接続を完全に抑止する方法は現状ありませんが、以下のレジストリまたはOffice 2016 のポリシーで、後述の一部の通信(\*I) を除いて抑止することは可能です。

  

<レジストリの場合\>

キー: HKEY\_CURRENT\_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Common\\Internet  
種類: REG\_DWORD  
名前: UseOnlineContent  
値: 0

値の説明)  
0 : ユーザーがインターネット上のOffice 2016 リソースにアクセスすることを禁止します。1 : ユーザーがインターネット上のOffice 2016 リソースにアクセスできるかどうかを自身で選択できるようにします。(以下の設定項目がOFF の状態)  
2 : ユーザーがインターネット上のOffice 2016 リソースにアクセスすることを許可します。(以下の設定項目がON の状態、既定の設定です)

\[Excel のオプション\] - \[セキュリティセンター\] - \[セキュリティセンターの設定\] - \[プライバシーオプション\] 内の "Office をMicrosoft のオンライン サービスに接続して、使用状況や環境設定に関連する機能を提供できるようにしますか？"

<グループ ポリシーの場合\>

\[ユーザーの構成\]  
\- \[Microsoft Office 2016\]  
\- \[ツール| オプション| 全般| サービス オプション...\]  
  \- \[オンライン コンテンツ\]  
  \- \[オンライン コンテンツのオプション\]  
  \- 有効: "Office にインターネットへの接続を許可しない" に設定します

  

補足)  
Office 2016 の管理用テンプレートはこちらからダウンロードしてください。

Office 2016 Administrative Template files (ADMX/ADML) and Office Customization Tool  
[http://www.microsoft.com/en-us/download/details.aspx?id=49030&WT](http://www.microsoft.com/en-us/download/details.aspx?id=49030&WT)

  

\*I :  
\~\~\~\~\~

上記レジストリまたはグループ ポリシーを設定しても、次のインターネット接続は抑止できません。

これは想定された動作です。

  

 [https://officeclient.microsoft.com/config16/](https://officeclient.microsoft.com/config16/?lcid=1041&syslcid=1041&uilcid=1041&build=16.0.4266&crev=3)

  

上記は、Office アプリケーションから利用される各種のネットワーク サービスへのエントリー ポイントになるURL などの情報を取得するために読み取りのみで利用されます。

Office 2016 では、プロセスの起動時にこのXML の情報を参照する処理が行われます。

一度参照に成功すると、取得された情報はキャッシュされるため、次回起動時に参照しませんが、Office の更新プログラム適用などによりバージョンが更新されると、新しい情報を取得するために再度通信が行われます。

Office 製品では得られた情報をもとに各種ネットワーク サービスを利用するため、この通信自体はOffice 2016 として必要な動作です。

  
  
  

**補足**
======

ネットワーク環境の構成によって、前述の資料「Office 365 ProPlus でのネットワーク要求」に記載された通信に時間がかかる、あるいは通信ができない場合は、ネットワークの応答待ちが発生し、その影響でOffice 2016 アプリケーションが一時的に応答なしになる場合があります。

この場合は、上記資料の通信が遅延などなく正常に行われているか確認してください。

  
  
  

**その他 - Office 2013 について -**


================================

本記事では Office 2016 について説明をしましたが、Office 2013 においても同様の動作となっております。  
適宜、Office 2016 を Office 2013 に読み換えてご参照ください。  
なお、Office 2013 で設定する場合は以下の変更点にご注意ください。

  

<レジストリの設定\>  
キー: HKEY\_CURRENT\_USER\\SOFTWARE\\Microsoft\\Office\\<span style="color:#ff0000">15.0</span>\\Common\\Internet  
種類: REG\_DWORD  
名前: UseOnlineContent  
値: 0

  

<グループポリシーの設定\>  
\[ユーザーの構成\]  
\- \[Microsoft Office <span style="color:#ff0000">2013</span>\]

  
管理用テンプレートのダウンロード URL は以下となります。  
[https://www.microsoft.com/en-us/download/details.aspx?id=35554](https://www.microsoft.com/en-us/download/details.aspx?id=35554)

  
  

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**