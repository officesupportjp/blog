---
title: MSI 版の Office 2016 で SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象について
date: 2020-09-23
lastupdate: 2021-04-23
---

こんにちは、Office サポートの西川 (直) です。

今回の投稿では、MSI 版の Office 2016 で SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象について説明します。

[MSI 版の Office 2016 とは？](https://officesupportjp.github.io/blog/%E3%82%AF%E3%82%A4%E3%83%83%E3%82%AF%E5%AE%9F%E8%A1%8C%E5%BD%A2%E5%BC%8F%20(C2R)%20%E3%81%A8%20Windows%20%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%A9%E3%83%BC%E5%BD%A2%E5%BC%8F%20(MSI)%20%E3%82%92%E8%A6%8B%E5%88%86%E3%81%91%E3%82%8B%E6%96%B9%E6%B3%95/)

以下の Blog 記事のとおり、「SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない」現象についてお問い合わせいただく場合があります。

Title : SharePoint Server など WebDAV が有効なサーバーから Office ファイルが開かない、保存できない現象について  
URL : [https://officesupportjp.github.io/blog/SharePoint Server など WebDAV が有効なサーバーから Office ファイルが開かない、保存できない現象について](https://officesupportjp.github.io/blog/SharePoint%20Server%20%E3%81%AA%E3%81%A9%20WebDAV%20%E3%81%8C%E6%9C%89%E5%8A%B9%E3%81%AA%E3%82%B5%E3%83%BC%E3%83%90%E3%83%BC%E3%81%8B%E3%82%89%20Office%20%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%8C%E9%96%8B%E3%81%8B%E3%81%AA%E3%81%84%E3%80%81%E4%BF%9D%E5%AD%98%E3%81%A7%E3%81%8D%E3%81%AA%E3%81%84%E7%8F%BE%E8%B1%A1%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6/)  

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