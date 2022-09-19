---
title: クイック実行形式 Office 環境で、Office 以外のアプリケーションから Office の OLEDB-ODBC を利用できます
date: '2020-08-07'
lastupdate: '2021-05-21'
id: cl0m6umwp0037i4vsflq54qjb
tags:
  - オートメーション

---

こんにちは、Office サポート チームです。  
  
<div style="color:#ff0000">
2020/12/01 Update  
Office ODBC が利用できるようになりましたので、タイトルと Office ODBC に関する箇所を修正/加筆いたしました。  
</div>
  
クイック実行形式 Office 製品をインストールした環境では、Office 以外のアプリケーションから Office の OLEDB プロバイダを利用できませんでしたが、Office を以下のバージョン以降に更新すると、Office 以外のアプリケーションから Office の OLEDB プロバイダ (Microsoft.ACE.OLEDB.12.0) を利用できるようになります。  
  
カレント チャネル : バージョン 2001 以降  
半期エンタープライズ チャネル : バージョン 2002 以降  
  
  
また、Office ODBC ドライバは、バージョン 2009 以降に更新することで、Office 以外のアプリケーションから利用できるようになります。  
(ユーザー DSN を作成いただけます)  
  
　カレント チャネル : バージョン 2009 以降  
　半期エンタープライズ チャネル : バージョン 2008 以降  

  
**更新手順**  
1\. Excel 等の Office アプリケーションを起動します  
2\. \[ファイル\] - \[アカウント\] を選択します。 \[更新オプション\] ボタンをクリックし、\[今すぐ更新\] をクリックします。  
ご利用の環境にて Office 製品の更新を管理されている場合は、更新方法についてシステム管理者にご確認ください。  

  

以下弊社サイトの「解決方法」 に、追加のコンポーネントの要否について一覧表でまとめておりますのでご一読ください。

Office クイック実行アプリケーションの外部にある Access ODBC、OLEDB、または DAO インターフェイスを使用できない  
[https://docs.microsoft.com/ja-jp/office/troubleshoot/access/cannot-use-odbc-or-oledb](https://docs.microsoft.com/ja-jp/office/troubleshoot/access/cannot-use-odbc-or-oledb)  
なお、上記記事の「ODBC 接続の作成に関する追加情報」に記載がありますように、バージョン 2009 以降もシステム DSN は作成できませんのでご留意ください。  
  

  
**補足**  
・以下の弊社記事にありますように、今まではクイック実行形式 Office 製品をインストールした環境では、Office 以外のアプリケーションからは Office の OLEDB/ODBC を利用することができませんでした。  
  
クイック実行形式の Office をインストールすると Office 以外のアプリケーションから ODBC - OLEDB が利用できない  
[https://officesupportjp.github.io/blog/クイック実行形式の Office をインストールすると Office 以外のアプリケーションから ODBC - OLEDB が利用できない](https://officesupportjp.github.io/blog/%E3%82%AF%E3%82%A4%E3%83%83%E3%82%AF%E5%AE%9F%E8%A1%8C%E5%BD%A2%E5%BC%8F%E3%81%AE%20Office%20%E3%82%92%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB%E3%81%99%E3%82%8B%E3%81%A8%20Office%20%E4%BB%A5%E5%A4%96%E3%81%AE%E3%82%A2%E3%83%97%E3%83%AA%E3%82%B1%E3%83%BC%E3%82%B7%E3%83%A7%E3%83%B3%E3%81%8B%E3%82%89%20ODBC%20-%20OLEDB%20%E3%81%8C%E5%88%A9%E7%94%A8%E3%81%A7%E3%81%8D%E3%81%AA%E3%81%84/)

  
  
本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。