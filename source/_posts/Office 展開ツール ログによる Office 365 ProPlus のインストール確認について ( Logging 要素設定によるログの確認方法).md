---
title: Office 展開ツール ログによる Office 365 ProPlus のインストール確認について ( Logging 要素設定によるログの確認方法)
date: '2019-03-09'
id: cl0m7opcg002adcvshfqe4ujr
tags:
  - インストール

---

(※ 2016 年 8 月 2 日に Japan Office Support Blog に公開した情報のアーカイブです。)

こんにちは、日本マイクロソフト Office サポート チームです。

  

今回は、Office 365 ProPlus インストール時の Office 展開ツール ログの確認方法についてご案内します。

  

Office 365 ProPlus インストール時の完了の確認方法は、イベント ログやレジストリなどいくつかありますが、今回は Office 展開ツールで出力されるログについてご案内します。

  

  

なお、 Office 製品は Office 365 ProPlus でインストールする Office 2013 および Office 2016 を対象としていますが、

  

今後、動作が変更される可能性があります。

  

  

**目次**

1.Office 展開ツールとは？  
2.Office 展開ツールのログの出力について  
3.ログの確認方法について  

  

  

**1.Office 展開ツールとは？**  

Office 展開ツールは、クイック実行版の Office を組織内に展開する際に、管理者にてインストールを管理できるツールです。

  

  

機能紹介  
  
\-----------

  

IT 管理者が、実行ファイル (setup.exe) と構成ファイル (XML ファイル形式) を利用して、配布用データのダウンロードや、クライアントへのインストールや更新を行うことができます。

  

本ツールを利用することで、例えば以下のような作業を実現することができます。

  

例)  

・Office のカスタマイズインストール  
 \+ インストール製品の指定  
 \+ 言語の指定  
 \+ 更新チャネルの指定  
 \+ インストールバージョンの指定  
 \+ インストールしないアプリケーションの指定  
 \+ ログの出力

  

・更新のダウンロードおよび管理  
 \+ 更新ファイルのダウンロード  
 \+ 更新を適用  
 \+ 更新チャネルの変更

  

・Office の再インストール  

  

タイトル :クイック実行用 Office 展開ツール  
アドレス : [https://technet.microsoft.com/ja-jp/library/jj219422.aspx](https://technet.microsoft.com/ja-jp/library/jj219422.aspx)

  

入手方法  
\--------------  

投稿時点で、2013 版と 2016 版の 2 種類があります。  
以下のリンクから入手することができます。  

  

2013 版の入手方法 : [http://www.microsoft.com/en-us/download/details.aspx?id=36778  
](http://www.microsoft.com/en-us/download/details.aspx?id=36778)
2016 版の入手方法 : [https://www.microsoft.com/en-us/download/details.aspx?id=49117  
](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

  

設定方法  
\-----------  

設定方法については、以下の情報をご参考いただけますと幸いです。

  

タイトル : クイック実行 configuration.xml ファイルのリファレンス  
アドレス : [http://technet.microsoft.com/ja-jp/library/jj219426.aspx](http://technet.microsoft.com/ja-jp/library/jj219426.aspx)

  
**2.Office 展開ツールのログの出力について**  


Office 展開ツールでは、Logging 要素を構成ファイル (Configurartion.xml) に記述することで、ログを出力することが可能です。  

  

構文  
\---------  

<Logging Level = Off | Standard Path = UNC or local path />

  

Logging 要素の属性と値  

|**属性**|**値**|**説明**|
| :--- | :---: | :--- |
|Level|Off<br/>Standard|省略可能。<br/>クイック実行 セットアップによって<br/>実行されるログ記録のオプションを指定します。<br/>既定のレベルは Standard です。|
|Path|%temp%<br/>\\\\server\\share\\userlogs\\ |省略可能。<br/>ログ ファイル用のフォルダーの<br/>完全修飾パスを指定します。<br/>既定値は %temp% です。|


  

  

例 :  
\<Logging Level="Standard" Path="%temp%" \/\>

  

タイトル :クイック実行 configuration.xml ファイルのリファレンス  
アドレス : [https://technet.microsoft.com/ja-jp/library/jj219426.aspx#BKMK\_LoggingElement](https://technet.microsoft.com/ja-jp/library/jj219426.aspx#BKMK_LoggingElement)

  
**3. ログの確認方法について**  


Office 365 ProPlus を、setup.exe /configure コマンドでインストール実行時に Standard レベルのログにて出力されるログでは、以下の Telemetry / General Telemetry の結果にて確認可能です。

  
  

Office  2013 の場合  
\-------------------------  

(サンプル ログここから)  
MM/DD/YYYY hh:mm:ss.SSS SETUP (0x1690)        0xac0                Click-To-Run Telemetry        aoh9z        Medium        AdminBootstrapper::Main: Installation came back with 0.  
(サンプル ログここまで)  

  

Office 2016 の場合  
\-------------------------  

(サンプル ログここから)  
MM/DD/YYYY hh:mm:ss.SSS SETUP (0xb2c) 0xa1c Click-To-Run General Telemetry arqpm Medium AdminBootstrapper::Main {"MachineId": "xxxxxxxxxxxxxxxxxxxxx", "SessionID": "6a424b9e-a112-47c8-8e8d-16e5d461b39b", "GeoID": 122, "Ver": "16.0.7070.2028", "C2RClientVer": "16.1.7070.2028", "ContextData": "AdminBootstrapper::Main: Installation came back with 0."} 63F743D9-0216-4508-82B8-CC0BA72985E1  
(サンプル ログここまで)

  

ログ内で、Telemetry / General Telemetry は複数箇所ありますが、最後に出力された Telemetry / General Telemetry の結果の "AdminBootstrapper::Main: Installation came back with 0." にて、エラーコード 0 で正常終了していることを確認可能です。

  

setup.exe /download コマンドでダウンロード実行時に出力されるログについても、同様の方法で確認可能です。

  

本情報の内容（添付文書、リンク先などを含む）は、作成日時点でのものであり、予告なく変更される場合があります。