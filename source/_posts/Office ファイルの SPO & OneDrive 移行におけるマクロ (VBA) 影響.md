---
title: Office ファイルの SPO & OneDrive 移行におけるマクロ (VBA) 影響
date: '2022-02-17'
id: cl0m4rqa2001y50vshhi75we2
tags:
  - オートメーション
  - クラウド ストレージ

---

<br>

*****  
**2022/3/18 Update**  
2. / 3. にそれぞれ、他のファイル内の VBA を参照している場合の影響と回避策を追加しました。
*****  
<br>

こんにちは、Office サポート チームの中村です。
  
Office ファイルの保存場所として、ファイル サーバーの共有フォルダから、SharePoint Online のドキュメント ライブラリや 個人用 OneDrive / OneDrive for Business への移行を検討する企業も増えてきたかと思います。Office ファイルでは業務効率化のためマクロ (VBA) を利用されていることが多く、弊社サポートにも、マクロ ファイル自身やマクロ内で操作するファイルの保存場所がクラウドになることの影響を確認したい、というお問い合わせを時折頂きます。
  
今後、さらにクラウド活用が進むと想定されるため、今回の記事で、このようなシナリオで一般的に生じる影響と対応方法、考慮すべき点についてまとめたいと思います。
<br>

## **1. クラウド移行で変化すること**
ファイルの格納場所を変更したとき、マクロに影響を与える要素は主に以下の 2 点です。

・フォルダやファイルのパスが UNC 共有パス形式 (\\\Server\\folder) から URL 形式 (http\://tenantname.sharepoint.com/xxxx など) に変わる  
・SharePoint や OneDrive によってアクセス権が管理される
<br>

## **2. ファイル パス変更の影響**
マクロからファイルやフォルダを操作するときに使用する関数によっては、URL 形式に対応していません。このような関数を利用している場合は、マクロ処理の見直しが必要です。以下に、マクロで主に使用されるファイル操作関数の URL 形式対応状況を記載します。
<br>

### **Excel / Word / PowerPoint のライブラリ**
これらのアプリ用のライブラリ Microsoft xxx 16.0 Object Library (xxx は各種アプリ名) で提供されるファイル操作メソッドやプロパティは、Web 上のファイルを直接読み書きすることを想定しているため、基本的に URL 形式を扱うことができます。これらのメソッド・プロパティを利用している場合は、特に対応の必要はありません。  
例) Excel の Workbooks.Open や Workbook.Path  
※ Access ファイルは Web から直接開くことができず、ダウンロードしてから開くため URL 形式を扱うことは想定されません。
<br>

**補足情報**  
エクスプローラーで、OneDrive や SharePoint ライブラリの同期フォルダからファイルを開く場合、以前の Office バージョンでは、ThisWorkbook.Path などで開いたファイルのパスを取得すると、同期フォルダのローカル パスが返されました。  
Microsoft 365 の最新チャネル バージョン 1903 および半期チャネル バージョン 2008 以降は、この操作シナリオでも、サーバー側の URL パスが返るようになりました。  
古い Office バージョンを利用する場合は、取得されるパスが異なることに注意してください。
<br>

### **VBA 関数**
Dir や Open といった VBA 関数は、URL 形式に対応していません。以下に、ファイル操作に関わる主な VBA 関数がまとめられていますが、ここに記載されている関数を利用している場合、他の仕組みへの変更を検討してください。

Directories and files keyword summary  
<https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/directories-and-files-keyword-summary>  
Input and output keyword summary  
<https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/input-and-output-keyword-summary>
<br>

### **FileSystemObject ライブラリ**
Scripting.FileSystemObject ライブラリで提供される各種メソッドは、URL 形式に対応していません。以下の公開情報で Scripting.FileSystemObject ライブラリで提供されるファイル・フォルダ操作メソッドがまとめられています。これらを利用している場合は、他の仕組みへの変更を検討してください。

FileSystemObject object  
<https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object>  
File object  
<https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object>  
Folder object  
<https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folder-object>
<br>

### **補足 1) その他、独自にパスを解析・組み立てる処理**  
マクロの独自処理として、例えばマクロ ファイル自身の格納フォルダ パスの一部を加工して処理結果ファイルの出力パスを組み立てるなど、パスのテキストを意識した処理がある場合は注意が必要です。URL 形式への変更に伴い、既存の文字列操作処理では、期待する結果が得られない場合があります。
<br>

### **補足 2) 他のファイルへの参照**
先述したような一般的なライブラリは、クライアント マシンに登録されたものを呼び出すため、ファイルの保存場所が変わっても影響を受けません。
しかし、他の Excel ブックを [参照設定] で参照して、そのブック内の独自に作成した VBA プロシージャを呼び出している場合、参照先のファイルをクラウドに移行すると、参照不可となります。(VBA の参照設定は URL 形式に対応していないため、参照設定を再設定してもクラウド上のファイルは参照できません。)

参照設定の他にも、マクロ処理内でパスを指定して他のファイルを開いている場合は、パスの指定方法がファイルの移動後も有効な実装となっているか十分に確認してください。
<br>

## **3. 影響を受ける場合の対処方法**
一般的には、以下のような対処方法が検討できます。実際の処理内容や環境構成、対応にかけられる工数などを考慮して対処方法をご検討ください。
<br>

### **a. Office ライブラリを用いて操作する**
先述の通り、Microsoft xxx 16.0 Object Library で提供されるメソッド・プロパティは基本的に URL 形式を扱うことができます。そして、Word や Excel は、テキスト ファイルを扱うこともできます。現在 VBA 関数や FileSystemObject でテキスト ファイルを操作している場合には、これを Word や Excel アプリで Microsoft xxx 16.0 Object Library を用いて扱うよう変更することが検討できます。
<br>

### **b. 不要なファイル・フォルダ操作を止める**
例えば、次回利用時に引き継ぎたい内容を動的に保持する設定ファイルを別ファイルとして用意し、これを読み書きするために対応していない関数を使っている、というシナリオを考えます。このような場合、これをマクロ ファイル自身のどこかに保持する (Excel の非表示シート、Word の隠し文字、ファイルのユーザー設定プロパティ等) ことで、他のファイルへの読み書きを行わないようにする、といった対応もマクロの内容次第では検討できます。
<br>

### **c. Graph API に変更する (ファイル操作と Excel 操作が対応)**
SharePoint Online や OneDrive に対し、HTTP リクエスト ベースで操作する API として Graph API が提供されています。現在、SharePoint ライブラリや OneDrive のフォルダ・ファイル自体の操作 (作成やコピー、削除など) と、Excel ファイルの操作用 API が利用できます。ただし、プログラムの実装イメージがマクロと大きく異なりますので、Graph API 自体の利用に習熟している開発者向けの代替案です。  

Microsoft Graph API を使用する  
<https://docs.microsoft.com/ja-jp/graph/use-the-api>  
Microsoft Graph REST API v1.0 リファレンス  
<https://docs.microsoft.com/ja-jp/graph/api/overview>  

現時点では、マクロで可能な全ての操作が Graph API からも行えるものではありませんが、簡単な操作は可能です。ただし、一般的に、マクロから呼び出すことには適していません。一般的には、独立したアプリケーションを作成して HTTP リクエストを発行し、Web 上のファイルを操作します。また、Microsoft 365 サービスへの認証も考慮する必要があります。  

**参考資料**  
(WebAPI)OAuth Bearer Token (Access Token) の取得方法について  
[https://officesupportjp.github.io/blog/cl0m8t2ac00023gvs6bwo9tp2/](https://officesupportjp.github.io/blog/cl0m8t2ac00023gvs6bwo9tp2/)  
(WebAPI)Microsoft Graph を使用した開発に便利なツール群  
[https://officesupportjp.github.io/blog/cl0m8t2ae00033gvsbtlo3btc/](https://officesupportjp.github.io/blog/cl0m8t2ae00033gvsbtlo3btc/)  
(WebAPI)Microsoft Graph - OneDrive API (C Sharp) を使ったサンプル コード  
[https://officesupportjp.github.io/blog/cl0m8t2aa00013gvsa2mr8xj5/](https://officesupportjp.github.io/blog/cl0m8t2aa00013gvsa2mr8xj5/)  
(WebAPI)Microsoft Graph – Excel REST API (C Sharp) を使い Range を操作するサンプル コード  
[https://officesupportjp.github.io/blog/cl0m8t2a400003gvsgcr96w6h/](https://officesupportjp.github.io/blog/cl0m8t2a400003gvsgcr96w6h/)  
<br>

### **d. OneDrive や SharePoint ライブラリの同期フォルダを操作する**
OneDrive 同期クライアントで、ファイル格納フォルダをクライアント コンピューターと同期すると、ローカルの同期フォルダにファイルがコピーされます。この同期フォルダに対してマクロから操作を行うと、通常のローカル パスとして扱うことができます。(なお、先述の通り、ThisWorkbook.Path の結果は最新の Microsoft 365 では URL 形式になり、同期フォルダのローカル パスではない点にご注意ください。)  

Windows の OneDrive とファイルを同期する  
<https://support.microsoft.com/ja-jp/office/615391c4-2bd3-4aae-a42a-858262e42a49>  

ただし、マクロから編集した内容がサーバー上へ同期されることは、マクロの動作としては保証されません。マクロからは、ローカルの同期フォルダへの操作が成功することをもって成功と判断されます。ネットワークの問題で同期に失敗したり、クラウド上のファイルに他のユーザーから競合する変更が行われて同期エラーが生じるなど、クラウド上にマクロの実行結果が期待通り反映されない可能性についての考慮が必要です。
<br>

#### **補足: WebDAV パスでのアクセス**  
SharePoint ライブラリを Internet Explorer で開くと、\[エクスプローラ―で開く\] というメニューが使用でき、"\\\SharePoint サイト\DavWWWRoot\Folder" のような UNC パスで、ライブラリが開かれていました。これは、WebDAV と呼ばれる Web フォルダを読み書き可能とする技術を用いて​実現されています。  
以前は、このパスを用いてマクロからアクセスすることも検討できる方法の 1 つでした。
しかしながら、この仕組みはレガシな技術を用いており、Edge ブラウザでは対応していません。  
また、WebDAV でのアクセスには別のサービス (WebClient サービス) が利用されるため、マクロが実行されている Office アプリの認証情報は引き継がれません。
WebClient サービス用の認証情報の保存にはまもなくサポートを終了する Internet Explorer を利用する必要があることなどからも、新たに WebDAV パスでの利用を検討することは推奨しません。
<br>

### **e. 参照設定を Application.Run に変更する**
参照設定で他のファイルの VBA プロシージャを呼び出している場合、参照設定を止め、以下のようにファイルを開いた上で Application.Run で実行することが検討できます。

例: 
Ref.xlsm に SampleProc() が含まれているとします。   

<変更前>  
[参照設定] から Ref.xlsm を参照し、VBA コードから SampleProc をそのまま呼び出し  

<変更後>  
参照設定は解除し、以下のように呼び出し  
```
Workbooks.Open "https://xxxx/Ref.xlsm"
Application.Run "Ref.xlsm!SampleProc"
```
<br>

## **4. Microsoft 365 サービスへの認証とアクセス権**
対応方法のうち、引き続きマクロから Microsoft 365 サービス上のファイルへのアクセスを行う "a. Office ライブラリを用いて操作する" の場合、マクロが実行されている Office アプリケーションにサインインしているユーザーとして認証が行われます。このため、マクロ実行時の Office サインイン ユーザーに、マクロで操作するファイルへのアクセス権を付与する必要があります。
<br>

## **5. サポートされる Office クライアント アプリケーション**
マクロからの操作においても、Microsoft 365 サービスとの接続がサポートされる Office クライアント アプリケーションがサポート対象です。サポートされる製品は、以下の公開情報でご確認ください。

Office versions and connectivity to Microsoft 365 services  
<https://docs.microsoft.com/en-us/deployoffice/endofsupport/microsoft-365-services-connectivity>


今回の投稿は以上です。
<br>

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**
