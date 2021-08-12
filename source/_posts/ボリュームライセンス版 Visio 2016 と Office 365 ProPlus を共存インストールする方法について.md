---
title: ボリュームライセンス版 Visio 2016 と Office 365 ProPlus を共存インストールする方法について
date: 2019-03-09
---

(※ 2017 年 4 月 7 日に Japan Office Support Blog に公開した情報のアーカイブです。)

こんにちは、Office サポートの山本です。  
本記事では、ボリュームライセンス版の Visio 2016 と、Office 365 ProPlus (Office 2016) を共存インストールする方法について、ご案内します。

ボリュームライセンス版の Visio 2016 と、Office 365 ProPlus (Office 2016) は、共存してインストールできません。  
これは、ボリュームライセンス版の Visio 2016 は、MSI 形式で実行され、Office 365 ProPlus (Office 2016) はクイック実行形式で実行されるという、インストール テクノロジーの差異によるものです。

[同一コンピューター上に異なるバージョンの Office、Visio、Project をインストールするためのサポート対象シナリオ](https://technet.microsoft.com/ja-jp/library/mt712177(v=office.16).aspx)  

しかし、ボリュームライセンス版の Visio 2016 に関しては、Office 2016 用の Office 展開ツールにて、ボリュームライセンス版の Visio 2016 の Product ID を指定しクイック実行形式でインストールすることで、Office 365 ProPlus (Office 2016) と共存してインストールすることが可能です。

[Office 展開ツールを使用して Visio 2016 および Project 2016 のボリューム ライセンス エディションをインストールする](https://technet.microsoft.com/ja-jp/library/mt703272.aspx)  

以下に詳細手順をご案内いたします。

### <span style="color:#0000ff">**インストール手順**</span>

1\. インストール関連ファイルを保存するためのフォルダを用意します。  
 ※ この手順では "C:Visio2016" を使用する例でご説明します。

2\. 以下のサイトの \[Download\] から、Office 展開ツールを入手して、手順 1) で作成したフォルダに保存します。

Office 2016 Deployment Tool  
[https://www.microsoft.com/en-us/download/details.aspx?id=49117](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

3\. ダウンロードしたファイルをダブルクリックします。  
 マイクロソフト ソフトウェア ライセンス条項の確認画面が表示されるので、チェックボックスにチェックを付け、\[Continue\] をクリックします。その後、手順 1 で作成したフォルダを指定してファイルを展開します。  
※ この手順では、"C:Visio2016" フォルダに保存する例でご説明します。

ダブルクリック後に展開されるファイル : configuration.xml, setup.exe

4\. ダウンロードした configuration.xml の内容を変更して、ボリュームライセンス版 Visio 2016 のダウンロード用構成ファイルを作成します。  
 configuration.xml をメモ帳で開き、全ての内容を削除した後、以下のサンプル コードの \<Configuration\> から \</Configuration\>までをコピーして、貼り付けます。編集したファイルを、任意のファイル名で .xml ファイルとして保存します。  
※ 下記のサンプルは、Visio Professional 2016 の 32 ビット版、Deferred Channel 日本語版をダウンロードするための構成ファイルとなります。  
ご利用の Visio のエディションによって、以下のように Product ID を変更してください。

<span style="color:#0000ff">**Product ID**</span>  
<span style="color:#0000ff">Visio Standard 2016</span> : VisioStdXVolume  
<span style="color:#0000ff">Visio Professional 2016</span> : VisioProXVolume  

※ 冒頭の「Add SourcePath=」以下のパスには、手順 3 で指定した setup.exe の保存パスを指定します。  
※ この手順では、"Visio-configuration.xml" という名前で保存する例でご説明します。

```
<Configuration>
  <Add SourcePath="C:\Visio2016" OfficeClientEdition="32" Channel="Deferred">
    <Product ID="VisioProXVolume">
      <Language ID="ja-jp" />
    </Product>
  </Add>
</Configuration>
```

5\. 以下のコマンドを実行して、インストール実行モジュールのダウンロードを行います。  
 \[FilePath\] 部分には、setup.exe の保存パス、および、Visio-configuration.xml の保存パスをそれぞれ指定します。 

```
> [FilePath]setup.exe /download [FilePath]Visio-configuration.xml

例 : > c:\Visio2016\setup.exe /download c:\Visio2016\Visio-configuration.xml
```

6\. コマンド実行後、プロンプトが戻ってくるまで少し待ちます。(完了時にメッセージは表示されません。)  
 ダウンロード完了後、「Add SourcePath=」で指定したフォルダ内に "Office" フォルダが作成され、インストール モジュールが保存されます。

7\. ダウンロード作業とインストール作業を行う PC が別である場合、ダウンロード完了後の C:\\Visio2016 フォルダの内容をすべてコピーし、インストールを行う PC にも C:\\Visio2016 フォルダを作成して貼り付けてください。

8\. 以下のコマンドを実行して、ダウンロードしたモジュールからインストールを実行します。

```
> [FilePath]setup.exe /configure [FilePath]Visio-configuration.xml

例 : > c:\Visio2016\setup.exe /configure c:\Visio2016\Visio-configuration.xml
```
9\. インストールが完了したら、Visio 2016を起動し、ライセンス認証を行います。  
 後述の “ライセンス認証について” の項目にお進みください。

<span style="color:#ff0000">**!!注意点!!**</span>  
Office 365 ProPlus が先にインストールされている場合、Office 2016 のビルド バージョンは、後からインストールした Visio 2016 のビルド バージョンに統一されます。  
Office 2016 と Visio 2016 を、同一端末にて、異なるビルド バージョンで利用することはできません。  
また、この手順にてインストールした Visio 2016 については、ボリュームライセンスでのご利用となりますが、クイック実行形式でインストールされます。  
インストール後の更新の管理についても、クイック実行形式と同様となりますので、ご注意ください。

### <span style="color:#0000ff">**ライセンス認証について**</span>

Office 展開ツールを使用して Visio 2016 ボリュームライセンス版をインストールする場合のライセンス認証について、以下にご説明します。  
現在、ボリューム ライセンス サービス センター (VLSC) で入手いただける MAK キーは、今回ご案内している方法でインストールした場合には、ご利用いただけません。  
インストールした Visio 2016 は、既定では KMS キーでのライセンス認証が必要となります。

既に Visio 2016 が認証可能な KMS ホストを構築が存在し、KMS ライセンス認証を実施できる環境にある場合には、追加作業の必要はありませんが、もし KMS ライセンス認証の環境が無い場合には、以下のいずれかの方法にて、ボリュームライセンスでの認証を実施する必要があります。

a) KMS ホストを構築し、ライセンス認証を行う  
b) 専用 C2R-P プロダクトキー (MAK キー) を入手して端末毎に入力し、ライセンス認証を行う

以下にそれぞれの手順をご案内いたします。

<span style="color:#0000ff">**a) KMS ホストを構築し、ライセンス認証を行う手順**</span>

通常のボリュームライセンス用の KMS ホストを構築することで、ライセンス認証が実施可能です。  
KMS でのライセンス認証を行うためには、Office 2016 (Visio 2016) を利用する 5 台以上のクライアントが存在する必要があります。  
Office 2016 KMS ホストの構築や、ライセンス認証の詳細手順については、以下をご参照ください。

Office 2016 KMS ホスト コンピューターの準備とセットアップ  
[https://technet.microsoft.com/ja-jp/library/dn385356(v=office.16).aspx](https://technet.microsoft.com/ja-jp/library/dn385356(v=office.16).aspx)

Office 2016 のボリューム ライセンス認証管理ツール  
[https://technet.microsoft.com/ja-jp/library/ee624350(v=office.16).aspx](https://technet.microsoft.com/ja-jp/library/ee624350(v=office.16).aspx)

<span style="color:#0000ff">**b)** **専用プロダクトキーでライセンス認証を行う手順**</span>

<del>前述のとおり、ボリューム ライセンス サービス センター (VLSC) で使用できる MAK キーでのライセンス認証は実施できませんが、対象端末についてのライセンス契約状況についてご確認のうえ、弊社までお問い合わせをいただくことで、専用のプロダクト キー  (MAK キー) の発行を行い、MAK ライセンス認証と同様に、ライセンス認証を実施することが可能です。  
KMS ホストの構築が難しい場合や、端末が 5 台未満の場合には、この方法を選択ください。 </del>

<span style="color:#ff0000">**※ 2017/8/8 Update**</span>  
ボリュームライセンス サービスセンター (VLSC) にて、Office 展開ツールを使用して Visio 2016 ボリュームライセンス版をインストールした場合専用のプロダクトキー (MAKキー) をご確認いただけるようになりました。  
VLSC にて製品ごとに、C2R-P キーという項目が自動的に追加され、専用のプロダクトキーが表示されておりますので、このプロダクトキーを MAK キーと同様に端末毎に入力し、ライセンス認証を実施してください。

例)
- Visio Professional 2016 C2R-P for use with the Office Development Tool  
- Visio Professional 2016 C2R-P 2016 MAK  

### <span style="color:#0000ff">**参考情報**</span>  

今回ご案内した詳細につきましては、以下の公開情報に記載しております。

同一コンピューター上に異なるバージョンの Office、Visio、Project をインストールするためのサポート対象シナリオ  
[https://technet.microsoft.com/ja-jp/library/mt712177(v=office.16).aspx](https://technet.microsoft.com/ja-jp/library/mt712177(v=office.16).aspx)

Office 展開ツールを使用して Visio 2016 および Project 2016 のボリューム ライセンス エディションをインストールする  
[https://technet.microsoft.com/ja-jp/library/mt703272.aspx](https://technet.microsoft.com/ja-jp/library/mt703272.aspx)

Office 展開ツール を利用する際の xml ファイル、および、コマンドの書式については、以下の技術情報をご覧ください。

概要: Office 展開ツール  
[https://technet.microsoft.com/ja-jp/library/jj219422.aspx  
](https://technet.microsoft.com/ja-jp/library/jj219422.aspx)

  
一般的な Office 展開ツールでのインストール手順については、以下でもご案内しております。

\[2017 年 3 月更新\] Office 365 ProPlus (Office 2016 バージョン) オフラインインストール手順  
[http://answers.microsoft.com/ja-jp/msoffice/forum/msoffice\_install-mso\_winother/office-365-proplus-2016/7208e618-5fc7-4b21-93a6-ab7dfb839409](http://answers.microsoft.com/ja-jp/msoffice/forum/msoffice_install-mso_winother/office-365-proplus-2016/7208e618-5fc7-4b21-93a6-ab7dfb839409)

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**