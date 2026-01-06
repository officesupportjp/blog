---
title: VBA マクロの Office バージョンアップ時の互換性
date: '2024-01-11'
lastupdate: '2025-01-06'
id: clr8yxe6w00043gsebywv7n74
tags:
  - VBA マクロ
---

<span style="color:#ff0000">**2026/01/06 Update**</span>  
<span style="color:#339966">- "付録 3: OS 更新による VBA 動作への影響" を追記しました。</span>  
<span style="color:#339966">- Office 2016 / 2019 のサポート終了に伴い、Office バージョンに言及する記載を更新しました。</span>  
<br>

こんにちは、Office サポート チームの中村です。

Office LTSC 2024 等の Office 永続ライセンスの新しいバージョンへの変更や、Microsoft 365 Apps のバージョン更新にあたって、これまで動作していた VBA マクロが引き続き同じように動作するのかを懸念するご質問を頂くことがあります。  
今回の記事では、このような Office バージョンアップ時の VBA マクロの互換性の考え方について記載します。

初めに結論を記載すると、元々利用していたバージョンが Office 2016 以降であれば、基本的な互換性は保持されており、ほとんどの場合は問題ありません。ただし、VBA マクロの処理の流れに依存して、Office の細かな動作変化が影響することは稀にあります。  
このような影響で動作しなくなるシナリオを網羅して弊社から案内することはできません。
最終的に確実な互換性を担保したい場合は、検証を通して確認していただく必要があります。

以下、VBA マクロの動作に関わる要素ごとに詳しく解説します。

**<h3>1. VBA ランタイム</h3>**
VBA はプログラム言語の 1 つであり、この VBA 言語を解釈して実行する VBA ランタイム自体にバージョンを持ちます。  
VBA ランタイムは Office アプリケーションに同梱されていますが、Office 2016 以降の永続ライセンス版 Office、および Microsoft 365 Apps のいずれもバージョン 7.1 で統一されており、VBA ランタイムのバージョン更新はありません。  
また仮に VBA 7.1 より古いバージョンで作成した VBA コードを含む Office ファイルであっても、VBA コードを実行環境で再コンパイルして実行できます。  

以上のことから、VBA ランタイムの観点で互換性を気にする必要はありません。
<br>

**<h3>2. Office オブジェクト ライブラリ</h3>**
VBA マクロの処理内容の中心は、Office アプリケーションに関する処理を行うことです。これには通常、Office オブジェクト ライブラリ内のオブジェクト / メソッド / プロパティなどを使用します。  
例えば、「Microsoft Excel 16.0 Object Library」ライブラリでは、Workbooks.Open メソッドなどの Excel アプリケーションとのやり取りを行うためのオブジェクトを提供しています。VBA マクロからは、[参照設定] でこれらのライブラリを参照したり、レイト バインディングと呼ばれる手法で実行時にライブラリを参照して利用します。

**<h4>2-1. ライブラリ バージョン</h4>**
これらの Office アプリケーションとやり取りするためのオブジェクトを提供するライブラリは、Office に同梱されています。Office 2016 以降の永続ライセンス版 Office、および Microsoft 365 Apps のいずれにおいても、バージョン「16.0」から変更はありません。このため、ライブラリ自体のバージョンが変わり、参照できなくなるといった問題は発生しません。

また、Office 2013 以前は「15.0」など、Office バージョンごとに異なるバージョンでした。この場合も、古いバージョンから新しいバージョンへの互換性は考慮されています。例えば Office 2013 で作成した VBA コードを、Microsoft 365 Apps 環境で実行する場合、自動的に新しいバージョンの同じライブラリを参照して動作します。  
ただし、反対に新しいバージョンで作成した VBA コードを古いバージョンで実行する場合は、利用アプリケーションとライブラリの組み合わせによっては参照できません。

**<h4>2-2. ライブラリ内のオブジェクト モデルの廃止 / 変更</h4>**
ライブラリ内で提供するオブジェクト モデルについては、Office 2016 より前のバージョンからの更新の場合、いくつかの互換性に影響が生じる廃止・変更があります。これらは、以下の公開情報でご案内しています。  

Office の互換性に関する問題  
[https://learn.microsoft.com/ja-jp/office/client-developer/shared/compatibility-issues-in-office](https://learn.microsoft.com/ja-jp/office/client-developer/shared/compatibility-issues-in-office)  

一方、Office 2016 以降は、既存のオブジェクト モデルに対し、メソッドなどが削除されたり、メソッド呼び出し時の必須パラメーターが追加されるといった、これまでのオブジェクト モデルの用法に影響を与えるような破壊的変更はありません。  
<br>

**<h3>3. Office アプリケーションの動作変化</h3>**
ここまでの 1. / 2. の観点からは、Office 2016 以降のバージョンから、より新しいバージョンへの更新において、通常 VBA マクロの動作に影響はありません。  
ただ重要な点として、VBA マクロは、Office アプリケーション等に対する処理要求を行っているということです。このため、VBA マクロからの処理においても、Office アプリケーションの動作変化の影響を受けます。  

例えば VBA マクロで Workbooks.Open を実行すると、Excel アプリケーションに対し「ファイルを開く」という要求を行い、実際にファイルを開く処理を行うのは VBA コード自身ではなく、Excel アプリケーション側です。そして、「ファイルを開く」という動作の中には、実際のファイル内容の読み取りの他、様々な内部処理が存在します。例えば、ファイルのフォーマットチェック、セキュリティ設定によるファイル内要素のブロック要否のチェック、最近使ったファイルへの情報追加・・・など、多岐に渡る内部処理をまとめて Workbooks.Open により実行される「ファイルを開く」動作となります。

ここまでに述べた通り、新しいバージョンの Office でも、Workbooks.Open メソッド自体はこれまでと同じ用法で実行することはできます。ただ、例えば Workbooks.Open から実行する「ファイルを開く」処理内のフォーマットチェック項目が強化される、といった Excel 側の動作変化はあり得ます。  
先述の例の場合、新しいチェック項目に反するフォーマットのファイルを VBA マクロから開くと、以前のバージョンでは開けるが、新しいバージョンでは開けない、といった事象が発生する可能性があります。

このような Office アプリケーションの動作変化は、バージョン更新に伴って無数に存在し、かつ内部の細かな実装であるため、すべての変更点を網羅して一覧化することはできません。リリース ノートで主要な変更の概要は確認できますが、記載のない細かな変更もあり、また、リリース ノートの端的な記載から影響有無を判断することは難しいことがほとんどです。  
また一方、ユーザーが作成する VBA マクロの処理内容も千差万別であり、どの Office 動作変化が VBA マクロの処理に影響するかを弊社で事前に推測し、「VBA に影響が出る変更リスト」のような形でまとめることも困難です。  

したがって、VBA マクロが新しい Office バージョンで確実に動作するかは、実際に新しい環境での動作検証を通してご確認ください。  

なお、このように記載しましたが、もちろん、広く一般的に想定される状況で、既存 VBA マクロの動作に影響するような変更は基本的には行いません。以下の公開情報でも触れている通り、ほとんどの VBA マクロはこれまでと同じように動作するため、過度に心配する必要はありません。

Microsoft 365 Appsをデプロイするための環境と要件を評価する  
手順 4 - アプリケーション互換性の評価  
[https://learn.microsoft.com/ja-jp/deployoffice/assess-microsoft-365-apps](https://learn.microsoft.com/ja-jp/deployoffice/assess-microsoft-365-apps)

社内環境の Office バージョンアップなどに際しては、VBA マクロを含むファイルは非常に多く存在することが一般的かと思います。そして、このようなファイルすべてを事前検証することは現実的ではないことが多いと考えられます。
例えば、業務上重要なファイルのみ事前検証を行い、その他のファイルは新バージョンの利用開始後に問題が確認された時点で、弊社サポートへのお問い合わせなどを通して調査するなど、皆様のご事情に合わせて、互換性に関する対応方針をご検討いただければ幸いです。  
<br>

**<h3>付録 1: Office 動作変化の具体例</h3>**
<3. Office アプリケーションの動作変化> について、Office の動作変化の具体例を紹介します。

**OneDrive 同期フォルダーの Workbook.Path**  
エクスプローラーで OneDrive や SharePoint ライブラリの同期フォルダからファイルを開く場合、以前の Office バージョンでは、ThisWorkbook.Path などで開いたファイルのパスを取得すると、同期フォルダのローカル パスが返されました。  
Microsoft 365 Apps の最新チャネル バージョン 1903 および半期チャネル バージョン 2008 以降、および Office LTSC 2021 では、このとき、クラウド ストレージ側の URL パスが返るようになりました。

この変化により、ローカル パスが返ることを期待して後続処理 (例: ”￥” 文字でのパス分割など) を記載していると影響が生じます。

**Excel 表示形式の和暦元年対応**  
 Office 2019 までは、Excel のセルの表示形式において、和暦の元年表記に未対応だったため、「令和1年5月1日」のように表示されていました。  
 Microsoft 365 のバージョン 1909 以降や Office LTSC 2021 以降の永続ライセンス版では、元年表記に対応するとともに、既定で元年表記が有効となっています。

 ![](Gannen.png)

この変化により、新しいバージョンでは、セルに「令和1年5月1日」のような和暦文字列を設定すると「令和元年5月1日」と表示内容が変わります。
<br>

**<h3>付録 2: bit 変更の影響</h3>**
Office のバージョンアップと合わせて、32 bit から 64 bit に変更することも多いかと思います。bit 変更に伴う VBA マクロへの影響として主に注意する点は、以下の 2 点です。
- Office 以外のライブラリを利用している場合、64 bit ライブラリを実行環境にインストール済みか
- Declare ステートメントを利用している場合、PtrSafe の付与と、パラメーターの Long 型を適切に LongPtr 型に変更しているか (パラメーターの値を扱う変数のデータ型も合わせて変更が必要)

詳しくは、以下の公開情報をご参照ください。  

64 ビット版または 32 ビット版の Office を選択する  
[https://support.office.com/ja-jp/article/2dee7807-8f95-4d0c-b5fe-6c6f49b8d261](https://support.office.com/ja-jp/article/2dee7807-8f95-4d0c-b5fe-6c6f49b8d261)

64 ビット Visual Basic for Applications の概要  
[https://learn.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview](https://learn.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview)

<br>

**<h3>付録 3: OS 更新による VBA 動作への影響</h3>**
Office のバージョン アップだけでなく、OS のバージョン アップで VBA の動作に影響が出るか、という質問を頂くこともありますが、考え方は「3. Office アプリケーションの動作変化」と同じです。  

Office アプリケーションは、内部的に OS で用意された API を呼び出し、OS の機能を利用して動作しています。またファイルを開く処理を例にすると、実際のファイルへのアクセスには、OS のファイル操作 API を呼び出しています。また、ファイル サーバーやクラウド ストレージにあるファイルの場合は SMB や HTTP でネットワーク通信を行う OS 機能も利用します。そして、開いたファイルを画面に表示するには、OS の描画 API を使います。 (これらは利用する OS 機能の一部の例です。)  

このように「VBA → Office アプリケーション → OS」と連携する仕組み上、OS API の動作変化が Office アプリケーションの動作、ひいては VBA の動作に影響する理論上の可能性はあります。  

ただし、OS API の動作は Office に限らず、その OS 上で動作するあらゆるアプリケーションに影響します。このため、Windows は互換性を非常に重視して設計されており、Office 更新に比べると OS 更新による VBA 動作への影響はさらに少なく、実際にはほぼありません。

参考)  
以下の Windows IT Pro Blog では、Windows OS の互換性への取り組みについて詳しくご案内しています。  

Microsoft extends application compatibility promise to Windows 11  
[https://techcommunity.microsoft.com/blog/windows-itpro-blog/microsoft-extends-application-compatibility-promise-to-windows-11/2810546](https://techcommunity.microsoft.com/blog/windows-itpro-blog/microsoft-extends-application-compatibility-promise-to-windows-11/2810546)

App confidence: From our compatibility story to yours  
[https://techcommunity.microsoft.com/blog/windows-itpro-blog/app-confidence-from-our-compatibility-story-to-yours/3750085](https://techcommunity.microsoft.com/blog/windows-itpro-blog/app-confidence-from-our-compatibility-story-to-yours/3750085)

<br>

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**
