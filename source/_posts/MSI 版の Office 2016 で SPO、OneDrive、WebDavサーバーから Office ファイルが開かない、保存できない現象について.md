---
title: MSI 版の Office 2016 で SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象について
date: 2020-09-23
lastupdate: 2021-04-23
---

こんにちは、Office サポートの西川 (直) です。

今回の投稿では、MSI 版の Office 2016 で SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象について説明します。

[MSI 版の Office 2016 とは？](https://social.msdn.microsoft.com/Forums/ja-JP/57e5d81e-ef69-4c1f-9ef0-932d03d0e7ce/1246312452124831246323455348922441824335-c2r-12392-windows?forum=officesupportteamja "クイック実行形式 (C2R) と Windows インストーラー形式 (MSI) を見分ける方法")

以下の Blog 記事のとおり、「SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない」現象についてお問い合わせいただく場合があります。

Title : SharePoint Server など WebDAV が有効なサーバーから Office ファイルが開かない、保存できない現象について  
URL : [https://social.msdn.microsoft.com/Forums/ja-JP/ea318b53-5f30-4a35-b21c-64ac83be977e/sharepoint-server-1239412393-webdav?forum=officesupportteamja](https://social.msdn.microsoft.com/Forums/ja-JP/ea318b53-5f30-4a35-b21c-64ac83be977e/sharepoint-server-1239412393-webdav?forum=officesupportteamja)  

上記の Blog 記事の方法では改善せず、ご利用されているのが MSI 版の Office 2016 の場合、セキュリティ更新のみ等、限定的に更新プログラムを適用していることに起因している可能性があります。

この場合、Microsoft Update を適用するか、少なくとも、以下のモジュールの更新プログラムを適用し、事象が回避するかをご確認いただくと解消する可能性があります。

**<実施方法>**

URL を開き、以下に列挙するモジュールの "Non-security" と "Security release" の日付を比較し、新しい方の KB の更新プログラムを適用してください。

なお、"Security KB superseded" は "置き換えられたセキュリティ更新の KB" を意味いたしますので、適用する必要はありません。

Title : List of the most current .msp files for Office 2016 products  
URL : [https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016](https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016)

・ace-x-none  
・csi-x-none  
・mso-x-none  
・msodll20-x-none  
・msodll30-x-none  
・msodll40ui-x-none  
・msodll99l-x-none  
・orgidcrl-x-none  
・msmipc-x-none

今回の投稿は以上です。  
  
**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**