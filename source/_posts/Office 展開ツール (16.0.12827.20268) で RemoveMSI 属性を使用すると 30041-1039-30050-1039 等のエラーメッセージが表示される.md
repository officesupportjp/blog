---
title: Office 展開ツール (16.0.12827.20268) で RemoveMSI 属性を使用すると 30041-1039-30050-1039 等のエラーメッセージが表示される
date: 2020-09-15
---

こんにちは、Office サポートの西川 (直) です。

表題の件について、既知の事象が報告されておりますのでご案内いたします。

**【事象】**  

Office 展開ツール (16.0.12827.20268) で 30041-1039 や 30050-1039 のエラーメッセージが表示される

**【説明】**  

本事象は 2020/6/9 にリリースされた Office 展開ツール (Version 16.0.12827.20268) に起因した事象である可能性がございます。

当該バージョンの Office 展開ツールでは、RemoveMSI 属性 で Office 2007 製品もアンインストールが可能となりましたが、  
バージョン番号が "12" から始まる一部の MSI インストール形式の製品がインストールされている環境で、アンインストールが失敗しエラーとなる事象が発生しております。

本不具合については、リリース時期はまだ未定ですが、次期バージョンの ODT で修正される見込みとなっております

**【回避策】**  

以下の何れかの回避策をご実施いただけますようお願いいたします。

・Office 2007 製品をコントロールパネルからアンインストールするか、以下の資料の OffScrub07.vbs により Office 2007 製品をアンインストールしてください。  

[https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls)  

・RemoveMSI 属性を削除してください

**【参考】**  

Release history for Office Deployment Tool

[](https://docs.microsoft.com/en-us/officeupdates/odt-release-history)[https://docs.microsoft.com/en-us/officeupdates/odt-release-history](https://docs.microsoft.com/en-us/officeupdates/odt-release-history)