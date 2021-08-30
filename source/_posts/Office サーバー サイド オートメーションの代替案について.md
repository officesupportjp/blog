---
title: Office サーバー サイド オートメーションの代替案について
date: 2019-02-28
---

(※ 2014年 5 月 7 日に Japan SharePoint Support Team Blog に公開した情報のアーカイブです。)

こんにちは、SharePoint サポートの 森　健吾 (kenmori) です。  
  
今回の投稿では、Office 開発サポートを担当していた時期に記載した下記投稿 Office サーバーサイド オートメーションの危険性の続編として、その代替案を記載いたします。

タイトル : Office サーバー サイド オートメーションの危険性について  
アドレス: [https://social.msdn.microsoft.com/Forums/ja-JP/e7a5c553-e01f-41c9-a277-40fdbb8e198e/office-12469125401249612540-124691245212489](https://social.msdn.microsoft.com/Forums/ja-JP/e7a5c553-e01f-41c9-a277-40fdbb8e198e/office-12469125401249612540-124691245212489)

上記投稿のおさらいとなりますが、例えば以下のような機能を Web サーバー上で Office API を使用したオートメーションで実装することはサポートされていません。

・夜間にデータを集計しOffice ファイルを作成  
・Web ページ上でファイルを選択して PDF などに変換  
   
本投稿では、この対処方法として以下の 2 つの方法をご紹介します。

1\. OpenXML SDK を使用した開発を実施する。  
2\. SharePoint Server 2013 を使用し、Word / PowerPoint Automation Service を使用したドキュメント変換を実施する。  
 

### 1\. OpenXML SDK を使用した開発を実施する。

Office サーバーサイド オートメーションの代替案として、まず筆頭に挙がるのが Open XML SDK を使用した開発となります。本ソリューションを実装するにあたり、SharePoint は必要ありません。  
以前の 97-2003 形式 (OLE 複合ドキュメント) 形式 (例. doc, xls, ppt) とは異なり、Open XML 形式の Office ファイル (例. docx, xlsx, pptx 等) は、公開された規格にそった XML によって成り立っています。Open XML 形式の Office ファイルは、拡張子を \*.zip に変更して開くと定義を確認できます。

そのため、Office アプリケーションがインストールされていなくても、ファイルを編集できるようになっています。

タイトル : Open XML SDK ドキュメント  
アドレス : [http://msdn.microsoft.com/ja-jp/library/bb226703.aspx](http://msdn.microsoft.com/ja-jp/library/bb226703.aspx)

タイトル : Open XML SDK 2.0 for Microsoft Office  
アドレス : [http://www.microsoft.com/ja-jp/download/details.aspx?id=5124](http://www.microsoft.com/ja-jp/download/details.aspx?id=5124)

ただし、多くの開発者が VBA などを通じて感覚的に養ってきたこれまでの開発手法と比較すると全く異なる内容である点や、膨大な SDK をはじめから覚えないといけないということで、移行のコスト等を敬遠していることをよく聞きます。

<span style="color:#ff0000">しかし、Open XML SDK 2.0 には Open XML 2.0 Productivity Tools for Microsoft Office という非常に強力な支援ツールがあります。</span>  
OpenXML Productivity Tools for Microsoft Office は上記 Open XML SDK 2.0 のサイトからダウンロードが可能です。  
インストール後、スタートメニューから起動し、さっそく動作を見てみましょう。

**使用手順**  
1\. 画面左上 \[ファイルを開く\] をクリックし、Office ファイルを選択して開きます。  
2\. \[コードの反映\] をクリックします。

 ![](image1.png) 

![](image2.png)  

参考) 読み込まれたファイル

これだけで、今読み込んだファイルを作成するのに必要なサンプル コード (ベストなコードではない) が右ペインに自動生成されます。やはり、新規のファイルから起こすと生成されるコード量は多いですが、多くの場合は編集を実施するだけですので、必要な個所をコピーペーストしながら新規コードに移植して実装すると想像よりも簡単に実装できると思います。

移行を考える場合は、このツールを駆使してコードの実装方法を確認し、必要な処理を組み込んでテストを実施するという流れになると思います。

ということで、さっそく私も自動生成されたコードをもとに、さっそくパラグラフを 1 つ追加してみました。WindowsBase.dll (System.IO.Packagin 名前空間) とDocumentFormat.OpenXml.dll を参照に加え、下記のコードを実装することで Hello World という文字を含むパラグラフが以下のコードで追加されました。

```
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
{
    // Productivity Tool が生成したコードの貼り付け
    Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00496DA6", RsidRunAdditionDefault = "00104CC6" };
    Run run1 = new Run();
    RunProperties runProperties1 = new RunProperties();
    RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
    runProperties1.Append(runFonts1);
    Text text1 = new Text();
    text1.Text = "Hello World";
    run1.Append(runProperties1);
    run1.Append(text1);
    paragraph1.Append(run1);
    // ここまで
    package.MainDocumentPart.Document.Body.Append(paragraph1);
}
```

上記画面キャプチャを見ても一目瞭然の通り、コードはほとんど Productivity Tool が自動生成したものです。

  
補足  
最新バージョンとして公開されている、Open XML SDK 2.5 for Microsoft Office は、英語のみのご提供となりますが、以下のサイトよりダウンロードすることが可能となります。  
Productivity Tool には Office 2013 フォーマットでの検証機能が追加されており、SDK のアセンブリは .NET Framework 4.0 ベースのアセンブリとして作成されております。必要に応じてご活用ください。  

タイトル : Open XML SDK 2.5 for Microsoft Office  
アドレス :[http://www.microsoft.com/en-us/download/details.aspx?id=30425](http://www.microsoft.com/en-us/download/details.aspx?id=30425)

### 2\. SharePoint Server 2013 を使用し、Word / PowerPoint Automation Service を使用したドキュメント変換を実施する。

Open XML SDK を使用することで、ファイルの自動編集は可能になりますが、私が Office 開発サポートを実施していた際の経験からも、サーバー サイド オートメーションにおける最もよくあるお問い合わせは、ドキュメントの変換 (例. PDF) や印刷です。

<span style="color:#ff0000">SharePoint Server 2013 (Standard Edition 以上) では、Word と PowerPoint においてのみ、ドキュメントの変換を実現できます。</span>残念ながら、Excel ワークブックについては、ファイル変換機能を持つサービスはありません。  
なお、Office はバックグラウンド サービスからの実行および印刷をサポートしておりません。また、このような構成で他形式のファイル (例. PDF) を印刷する場合は該当アプリケーションの開発元にご確認ください。  
  
その他、オートメーション サービスを使用したカスタマイズは SSOM (サーバーサイド オブジェクト モデル) のみの提供となりますため、SharePoint Online では本機能はご利用いただけません。

**Word**

Word Automation Service が提供する ConversionJob.Start メソッドを使用した開発にて実現することが可能です。この機能は SharePoint Server 2010 以降にてご利用可能です。

  
事前準備  
•サーバーの全体管理コンソールの \[システム設定\] で、\[サーバーのサービスの管理\] を選択し、\[Word Automation Service\] が \[開始\] に設定されていることを確認します。  
•サーバーの全体管理コンソールの \[アプリケーション構成の管理\] で \[サービス アプリケーションの管理\] を選択し、\[Word Automation Services\] および \[Word Automation Services Proxy\] が \[開始済み\] に設定されていることを確認します。

  
タイトル : SharePoint 2010 Word Automation Services での開発  
アドレス : [http://msdn.microsoft.com/ja-jp/library/ff742315(v=office.14).aspx](http://msdn.microsoft.com/ja-jp/library/ff742315(v=office.14).aspx)

タイトル : ConversionJob.Start method  
アドレス : [http://msdn.microsoft.com/en-us/library/office/microsoft.office.word.server.conversions.conversionjob.start.aspx](http://msdn.microsoft.com/en-us/library/office/microsoft.office.word.server.conversions.conversionjob.start.aspx)

下記のようなコードを実装して、ドキュメントの変換を実施します。上記サイトからの抜粋となります。

```
\-- 抜粋 開始 --

   string siteUrl = "http://localhost";
    // If you manually installed Word automation services, then replace the name
    // in the following line with the name that you assigned to the service when
    // you installed it.
    string wordAutomationServiceName = "Word Automation Services";
    using (SPSite spSite = new SPSite(siteUrl))
    {
        ConversionJob job = new ConversionJob(wordAutomationServiceName);
        job.UserToken = spSite.UserToken;
        job.Settings.UpdateFields = true;
        job.Settings.OutputFormat = SaveFormat.PDF;
        job.AddFile(siteUrl + "/Shared%20Documents/Test.docx",
        siteUrl + "/Shared%20Documents/Test.pdf");
        job.Start();
    }

\-- 抜粋 終了 --
```

**PowerPoint**

PowerPoint Automation Service が提供する PdfRequest.BeginConvert メソッドを実行することでドキュメントの変換が可能です。

  
事前準備  
•サーバーの全体管理コンソールの \[システム設定\] で、\[サーバーのサービスの管理\] を選択し、\[PowerPoint Conversion Service\] が \[開始\] に設定されていることを確認します。  
•サーバーの全体管理コンソールの \[アプリケーション構成の管理\] で \[サービス アプリケーションの管理\] を選択し、\[PowerPoint Conversion Service Application\] および \[PowerPoint Conversion Service Application Proxy\] が \[開始済み\] に設定されていることを確認します。

タイトル : SharePoint 2013 の PowerPoint Automation Services  
アドレス : [http://msdn.microsoft.com/ja-jp/library/office/fp179894(v=office.15).aspx](http://msdn.microsoft.com/ja-jp/library/office/fp179894(v=office.15).aspx)

タイトル : PdfRequest class  
アドレス : [http://msdn.microsoft.com/en-us/library/office/microsoft.office.server.powerpoint.conversion.pdfrequest(v=office.15).aspx](http://msdn.microsoft.com/en-us/library/office/microsoft.office.server.powerpoint.conversion.pdfrequest(v=office.15).aspx)

下記のようなコードを実装して、ドキュメントの変換を実施します。上記サイトからの抜粋となります。

```
\-- 抜粋 開始 --

  string siteURL = "http://localhost";
  using (SPSite site = new SPSite(siteURL))
  {
      using (SPWeb web = site.OpenWeb())
      {
          Console.WriteLine("Begin conversion");

          // Get a reference to the "Shared Documents" library
          // and the presentation file to be converted.
          SPFolder docs = web.Folders[siteURL + "/Shared Documents"];
          SPFile file = docs.Files[siteURL + "/Shared Documents/Pres1.ppt"];

          // Convert the file to a stream and create an
          // SPFileStream object for the conversion output.
          Stream fStream = file.OpenBinaryStream();
          SPFileStream stream = new SPFileStream(web, 0x1000);

          // Create the presentation conversion request.
          PresentationRequest request = new PresentationRequest(fStream, ".ppt", stream);

          // Send the request synchronously, passing
          // in a ‘null’ value for the callback parameter,
          // and capturing the response in the result object.
          IAsyncResult result = request.BeginConvert(
              SPServiceContext.GetContext(site),
              null,
              null);
          // Use the EndConvert method to get the result.
          request.EndConvert(result);

          // Add the converted file to the document library.
          SPFile newFile = docs.Files.Add(
              "newPres1.pptx",
              stream,
              true);
          Console.WriteLine("Output: {0}", newFile.Url);
      }
  }

\-- 抜粋 終了 --
```
  

いかがでしたでしょうか。  
Open XML および SharePoint Server 2013 における 2 つのオートメーション サービスについてご説明しました。  
上記情報をお役立ていただき、少しでも移行先のテクノロジーに対する抵抗をなくしていただけますと幸いです。

  
今回の投稿は以上となります。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**