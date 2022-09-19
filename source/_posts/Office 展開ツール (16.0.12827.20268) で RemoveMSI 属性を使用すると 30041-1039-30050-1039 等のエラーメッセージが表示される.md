---
title: >-
  Office 展開ツール (16.0.12827.20268) で RemoveMSI 属性を使用すると 30041-1039-30050-1039
  等のエラーメッセージが表示される
date: '2020-09-15'
lastupdate: '2020-10-30'
id: cl0m69xux00254cvs8a1nb7qn
tags:
  - インストール

---

こんにちは、Office サポートの西川 (直) です。

表題の件について、既知の事象が報告されておりますのでご案内いたします。

**【事象】**  

Office 展開ツール (16.0.12827.20268) で 30041-1039 や 30050-1039 のエラーメッセージが表示される

**【説明】**  

本事象は 2020/6/9 にリリースされた Office 展開ツール (Version 16.0.12827.20268) に起因した事象である可能性がございます。

当該バージョンの Office 展開ツールでは、RemoveMSI 属性 で Office 2007 製品もアンインストールが可能となりましたが、  
バージョン番号が "12" から始まる一部の MSI インストール形式の製品がインストールされている環境で、アンインストールが失敗しエラーとなる事象が発生しております。

<div style="color:#ff0000">
本事象は、最新の Office 展開ツールで改善されておりますので、最新の Office 展開ツールをご利用いただきますようお願いいたします。

Title : Release history for Office Deployment Tool  
URL : [https://docs.microsoft.com/en-us/officeupdates/odt-release-history](https://docs.microsoft.com/en-us/officeupdates/odt-release-history)  
\----------------抜粋開始----------------  

October 29, 2020  
Version 16.0.13328.20292 (setup.exe version 16.0.13328.20290)  
• Resolves an issue where certain Office 2007 products may unexpectedly block installation when using RemoveMSI  

\----------------抜粋終了----------------  
  
Title : Office Deployment Tool  
URL : [https://www.microsoft.com/en-us/download/details.aspx?id=49117](https://www.microsoft.com/en-us/download/details.aspx?id=49117)  

</div>

~~本不具合については、リリース時期はまだ未定ですが、次期バージョンの ODT で修正される見込みとなっております~~

~~**【回避策】**~~  

~~以下の何れかの回避策をご実施いただけますようお願いいたします。~~

~~・Office 2007 製品をコントロールパネルからアンインストールするか、以下の資料の OffScrub07.vbs により Office 2007 製品をアンインストールしてください。~~  

~~[https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls)~~  

~~・RemoveMSI 属性を削除してください~~

~~**【参考】**~~  

~~Release history for Office Deployment Tool~~

~~[](https://docs.microsoft.com/en-us/officeupdates/odt-release-history)[https://docs.microsoft.com/en-us/officeupdates/odt-release-history](https://docs.microsoft.com/en-us/officeupdates/odt-release-history)~~  

**2020/10/15 :** **回避策について説明を追記しました。**

**2020/10/30 :** **回避策について対応バージョンを変更しました。**

   
今回の投稿は以上です。  
  
**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**