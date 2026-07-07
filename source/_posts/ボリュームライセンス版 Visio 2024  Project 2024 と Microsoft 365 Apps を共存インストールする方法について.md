---
title: >-
  ボリュームライセンス版 Visio 2024 / Project 2024 と Microsoft 365 Apps を共存インストールする方法について
date: '2026-07-07'
id: cmraa1xdj0000zw6y82mu5nfw
tags:
  - インストール
---

こんにちは、Office サポート チームです。  
本記事では、ボリュームライセンス版の Visio 2024 / Project 2024 と Microsoft 365 Apps を共存インストールする方法についてご案内します。


ボリュームライセンス版の Visio 2024 / Project 2024 と Microsoft 365 Apps は、共存してインストールすることが可能です。

### <span style="color:#0000ff">**前提条件**</span>

共存インストールのためにはバージョンによって以下の前提条件があるため、ご注意ください。

- Visio LTSC 2024 / Project 2024 を共存インストールする場合
Microsoft 365 Apps バージョン 2408 以上を使用している必要があります。

また、同一端末に共存してインストール可能なクイック実行形式の Office 製品は、同一のチャネル、バージョン (ビルド)、bit である必要があります。
Microsoft 365 Apps と Visio / Project を異なるバージョンとする等、複数のチャネルやビルド バージョンを同一端末にインストールすることはできません。


以下にインストールの詳細手順をご案内いたします。

### <span style="color:#0000ff">**インストール手順**</span>
※インストール手順は Visio LTSC 2024 をインストールする手順を示します。 Project 2024 も同様の手順でインストールすることができます。

1. 現在インストールされている M365Apps の更新チャネル、ビルドを確認します。
  いずれかの Microsoft 365 Apps のソフトウェアを起動し、[ファイル] - [アカウント] の順にクリックし、[Wordのバージョン情報] (Word の場合) 欄に表示されている更新チャネル、ビルドを記録してください。

    例
    ```
    バージョン 2605 (ビルド 20026.20166 クイック実行)
    月次エンタープライズ チャネル
    ```

    更新チャネル：月次エンタープライズ チャネル (MonthlyEnterprise)  
    ビルド：20026.20166

2. インストール関連ファイルを保存するためのフォルダを用意します。  
 ※ この手順では "C:\Visio2024" を使用する例でご説明します。

3. 以下のサイトの \[Download\] から、Office 展開ツールを入手して、手順 2 で作成したフォルダに保存します。 

Office 展開ツール (ODT)  
https://www.microsoft.com/en-us/download/details.aspx?id=49117  
ファイル名 [officedeploymenttool_xxxxx-xxxxx.exe]  

※ Office LTSC 2024 製品をインストールするには、Office 展開ツールの setup.exe が 2024/08/12 にリリースされた「setup.exe バージョン 16.0.17830.20162」以降である必要がございます。  
これより古いバージョンの ODT ではインストールできないため、すでに ODT をお持ちの場合でも、最新版を入手し直してご利用ください。  
setup.exe のバージョンは、後述の手順 4 でファイルを展開した後、対象の setup.exe を右クリックして [プロパティ] を開き、[詳細] タブの [ファイル バージョン] からご確認いただけます。

4. ダウンロードしたファイルをダブルクリックします。  
  マイクロソフト ソフトウェア ライセンス条項の確認画面が表示されるので、チェックボックスにチェックを付け、\[Continue\] をクリックします。その後、手順 2 で作成したフォルダを指定してファイルを展開します。  
  ※ この手順では、"C:\Visio2024" フォルダに保存する例でご説明します。
    - 展開後に生成されるファイル
      - configuration-Office365-x64.xml
      - setup.exe

5. ダウンロード用構成ファイルを作成します。

    ※この手順はツールを使った構成ファイル作成方法もございます。後述の「Officeカスタマイズツールを使った XML ファイルの作成」をご参照ください。

    展開した上記ファイルのうち、いずれかの XML ファイルをメモ帳で開き、全ての内容を削除した後、以下のサンプル コードの <Configuration\> から </Configuration\>までをコピーして、貼り付けます。  
    編集したファイルを、任意のファイル名で .xml ファイルとして保存します。  
    Channelの要素は手順 1 で取得したチャネル、Versionの要素は "16.0." + (手順 1 で取得したビルド)となります。  
    ※ この手順では、"Visio-configuration.xml" という名前で保存する例でご説明します。

    ```
    <Configuration>
      <Add SourcePath="C:\Visio2024" OfficeClientEdition="64" Channel="MonthlyEnterprise" Version="16.0.20026.20166">
        <Product ID="VisioPro2024Volume" PIDKEY="B7TN8-FJ8V3-7QYCP-HQPMV-YY89G">
          <Language ID="ja-jp" />
        </Product>
      </Add>
    </Configuration>
    ```

    ※ 冒頭の「Add SourcePath=」以下のパスには、手順 3 で指定した setup.exe の保存パスを指定します。  
    ※ 上記のサンプルは、Visio LTSC Professional 2024 の 64 ビット版、月次エンタープライズ チャネル、日本語版をダウンロードするための構成ファイルとなります。  
    ご利用の Visio/Project のエディションによって、以下のように Product ID と PIDKEY を変更してください。<br>
      |  Product  |  ID  |  PIDKEY  |
      | ---- | ---- |  --  |
      |  Visio LTSC Standard 2024  |  VisioStd2024Volume  | JMMVY-XFNQC-KK4HK-9H7R3-WQQTV |
      |  Visio LTSC Professional 2024  |  VisioPro2024Volume  | B7TN8-FJ8V3-7QYCP-HQPMV-YY89G |
      |  Project Standard 2024  |  ProjectStd2024Volume  | PD3TT-NTHQQ-VC7CY-MFXK3-G87F8 |
      |  Project Professional 2024  |  ProjectPro2024Volume  | FQQ23-N4YCY-73HQ3-FM9WC-76HF4 |
    

    ※ 上記 PIDKEY は KMS 認証を行うための GVLK となります。MAK 認証を行う場合、個別の専用プロダクトキーを入力ください。詳細は後述の「ライセンス認証について」の b) に記載されております。


    KMS および Office、Project、Visio の Active Directory によるライセンス認証用の GVLKs  
    https://learn.microsoft.com/ja-jp/DeployOffice/vlactivation/gvlks?redirectedfrom=MSDN


    下記公開情報をご参照のうえ、適宜 XML ファイルの要素を変更してください<br>

    Office 展開ツールのオプションの構成  
    https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/office-deployment-tool-configuration-options


6. 以下のコマンドを実行して、インストール実行モジュールのダウンロードを行います。コマンドは、「管理者として実行」したコマンド プロンプトから行ってください。

  1) スタート ボタンの横の検索ボックスに「cmd」と入力します。
  2) 検索結果に表示された [コマンド プロンプト] を右クリックし、[管理者として実行] をクリックします。
  3) ウィンドウのタイトル バーに「管理者:」（英語版では「Administrator:」）と表示されていれば、管理者として起動できています。  
  \[FilePath\] 部分には、setup.exe の保存パス、および、Visio-configuration.xml の保存パスをそれぞれ指定します。  
  ```
    > [FilePath]setup.exe /download [FilePath]Visio-configuration.xml
    例 : > C:\Visio2016\setup.exe /download C:\Visio2019\Visio-configuration.xml   
   ```

7. コマンド実行後、プロンプトが戻ってくるまで少し待ちます。(完了時にメッセージは表示されません。)  
 ダウンロード完了後、「Add SourcePath=」で指定したフォルダ内に "Office" フォルダが作成され、インストール モジュールが保存されます。

8. ダウンロード作業とインストール作業を行う PC が別である場合、ダウンロード完了後の C:\Visio2024 フォルダの内容をすべてコピーし、インストールを行う PC にも C:\Visio2024 フォルダを作成して貼り付けてください。  
※ クライアントがインターネット（Office CDN）に接続できないクローズド環境では、必ずこのローカル ソース方式を使用してください。CDN に到達できない場合、エラー 30182-1016 などが発生することがあります。

9. 以下のコマンドを実行して、ダウンロードしたモジュールからインストールを実行します。
    ```
    > [FilePath]setup.exe /configure [FilePath]Visio-configuration.xml
    例 : > C:\Visio2024\setup.exe /configure C:\Visio2024\Visio-configuration.xml
    ```

10. インストールが完了したら、Visio 2024のライセンス認証を行います。  
    ライセンス認証するために、Office ソフトウェア保護プラットフォーム スクリプト (ospp.vbs) を使います。
    64 ビット版では Program Files\Microsoft Office\Office16 フォルダに、32 ビット版では Program Files (x86)\Microsoft Office\Office16 フォルダにあります。以下が認証するためのコマンドになります。

    ospp.vbsのあるフォルダに移動します。
    ```
    cd C:\Program Files\Microsoft Office\Office16
    ```

    ospp.vbsを実行して、ライセンス認証を行います。
    ```
    cscript ospp.vbs /act
    ```

    実行結果として、VisioPro2024 について Product activation successful と表示されると認証成功となります。
    ※ ライセンス認証の詳細に関しましては、後述の「ライセンス認証について」をご参照ください。

### <span style="color:#0000ff">**Officeカスタマイズツールを使った XML ファイルの作成**</span>

1. Office カスタマイズ ツールにアクセス

    Office カスタマイズ ツール ( https://config.office.com )にアクセスしてください。

2. 展開の設定

    構成する製品、言語、およびアプリケーションの基本設定を選択します。 
    
    今回のインストール手順では以下の設定項目が最低限必要となります。その他の項目は適宜設定を追加してください。

    - 製品とリソース
        
        アーキテクチャで対応するビットを選択、製品は今回のインストールする Visio / Project を選択してください。
        
        ※各構成ファイルでデプロイできるアーキテクチャは 1 つだけです。

        更新チャネルでは手順 1 で取得したチャネルとバージョンを選択してください。  
        ※ Office LTSC 2024 永続エンタープライズ チャネル / Office LTSC 2021 永続エンタープライズ チャネル が選択されていると、Microsoft 365 Apps と共存できません。 
    
    - 言語

        主言語を選択してください。オペレーティング システムに合わせる を選択すると、クライアント デバイスで使用されているのと同じ言語が自動的にインストールされます。

    - インストール
        
        Office ファイルをクラウドから直接インストールするか、ネットワーク上のローカル ソースから直接インストールするかを選択します。今回のインストール手順では [インストールのオプション] - [ローカル ソース] よりソース パスに C:\Visio2024 と入力してください。

    - ライセンスとアクティブ化

        プロダクト キーは試用する認証方式を選択してください。MAK 認証を選択した場合は適切なライセンスキーを入力してください。

3. 構成ファイルのエクスポート

    右側のウィンドウで構成済みの設定を確認し、[エクスポート] を選択してください。
    使用許諾契約書の条項に同意し、構成ファイルに名前を付けてから、[ エクスポート ] を選択してダウンロードしてください。今回の手順では Visio-configuration と名前を設定します。

4. 構成ファイルの配置

    ダウンロードしたファイルを C:\Visio2024 に移動し、インストール手順 6 よりインストールを進めてください。


Office カスタマイズ ツールの概要  
https://learn.microsoft.com/ja-jp/microsoft-365-apps/admin-center/overview-office-customization-tool



### <span style="color:#0000ff">**ライセンス認証について**</span>
ボリュームライセンス版の Office では、KMS 認証や MAK 認証などによりライセンス認証が実施可能です。
KMS 認証を行う場合、事前に KMS ホストを構築し、認証を行える環境を構築いただく必要がございます。
それぞれの詳細につきましては、以下をご参照ください。


<span style="color:#0000ff">**a) KMS ホストを構築し、ライセンス認証を行う手順**</span>

通常のボリュームライセンス用の KMS ホストを構築することで、ライセンス認証が実施可能です。  
KMS でのライセンス認証を行うためには、対象 Office を利用する 5 台以上のクライアントが存在する必要があります。  
ボリューム ライセンス版の Office のライセンス認証の KMS ホストの構築や、ライセンス認証の詳細手順については、以下をご参照ください。

Office のボリューム ライセンス認証の概要  
https://learn.microsoft.com/ja-jp/office/volume-license-activation/plan-volume-activation-of-office

ボリューム ライセンス版の Office のライセンス認証を行うように KMS ホスト コンピューターを構成する  
https://learn.microsoft.com/ja-jp/office/volume-license-activation/configure-a-kms-host-computer-for-office



<span style="color:#0000ff">**b) MAK を使用したコンピューターのライセンス認証手順**</span>

Office 展開ツールで使用される configuration.xml ファイルの PIDKEY を指定し、認証を行います。
今回のインストールの手順ではサンプル XML ファイルの 「PIDKY=　」 に KMS キーではなく MAK キーを入力してください。
詳細手順については、以下をご参照ください。

MAK を使用してボリューム ライセンスバージョンの Office をアクティブ化する  
https://learn.microsoft.com/ja-jp/office/volume-license-activation/activate-office-by-using-mak

Office LTSC 2024 を展開する  
https://learn.microsoft.com/ja-jp/office/ltsc/2024/deploy



### <span style="color:#0000ff">**参考情報**</span>  
今回ご案内した詳細につきましては、以下の公開情報に記載しております。
<br>
同一コンピューター上に異なるバージョンの Office、Visio、Project をインストールするためのサポート対象シナリオ  
https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/install-different-office-visio-and-project-versions-on-the-same-computer

Office 展開ツール を利用する際の xml ファイル、および、コマンドの書式については、以下の技術情報をご覧ください。
<br>
概要: Office 展開ツール  
https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/overview-office-deployment-tool

Visio の展開ガイド  
https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/deployment-guide-for-visio

Project の展開ガイド  
https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/deployment-guide-for-project


**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**

