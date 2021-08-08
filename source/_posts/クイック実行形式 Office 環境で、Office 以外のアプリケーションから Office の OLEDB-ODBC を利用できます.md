---
title: クイック実行形式 Office 環境で、Office 以外のアプリケーションから Office の OLEDB-ODBC を利用できます
date: 2020-08-07
lastupdate: 2021-05-21
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
  
クイック実行形式の Office をインストールすると Office 以外のアプリケーションから ODBC / OLEDB が利用できない  
[https://social.msdn.microsoft.com/Forums/ja-JP/32778472-58f7-4840-8d90-18c2bdd6cb6b/124631245212483124632345534892244182433512398-office?forum=officesupportteamja](https://social.msdn.microsoft.com/Forums/ja-JP/32778472-58f7-4840-8d90-18c2bdd6cb6b/124631245212483124632345534892244182433512398-office?forum=officesupportteamja)

  
  
本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。