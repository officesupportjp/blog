---
title: Office のパフォーマンス トラブルシュート (その 4：はじめに試す設定 – 起動・オープン①、印刷編 -)
date: 2019-12-19
---

こんにちは、Office サポート チームの中村です。

[前回](https://officesupportjp.github.io/blog/Office%20%E3%81%AE%E3%83%91%E3%83%95%E3%82%A9%E3%83%BC%E3%83%9E%E3%83%B3%E3%82%B9%20%E3%83%88%E3%83%A9%E3%83%96%E3%83%AB%E3%82%B7%E3%83%A5%E3%83%BC%E3%83%88%20(%E3%81%9D%E3%81%AE%203%EF%BC%9A%E3%81%AF%E3%81%98%E3%82%81%E3%81%AB%E8%A9%A6%E3%81%99%E8%A8%AD%E5%AE%9A%20-%20%E6%8F%8F%E7%94%BB%E7%B7%A8%20-)/)に続いて、パフォーマンス改善効果が期待できる設定をご紹介します。

今回は、Office アプリケーションの起動が遅い場合、およびファイルを開くのが遅い状況に効果が期待できる設定を紹介します。

一般的なユーザー操作では Office ファイルをダブルクリックで開くことが多く、このような場合は起動とファイルを開く操作がまとめて行われ、どちらの処理が遅いのか判断しづらいことがあります。このため、この 2 つの動作 (Office の起動、ファイルを開く) に関する設定を合わせて紹介しています。また、印刷に関するパフォーマンス低下要因は詳細を後述する通りファイルを開くタイミングにも影響することがあるため、合わせてこの資料で紹介します。

なお、起動・オープン編は次の記事でもさらに設定項目を紹介予定です。
<br>

## **目次**  
1\. アドインの無効化  
2\. 既定のプリンターの変更  
3\. ユーザー設定のビュー削除 (Excel 共有ブック)  
4\. 改ページ情報の非表示 (Excel)  
5\. ファイル・フォルダ履歴削除  
6\. ファイル検証の無効化
<br>

## **1\.** **アドインの無効化**  
Office アドインが設定されている場合、起動時に読み込みが行われます。アドインが有効になっていると、読み込み処理自体や、アドインのバージョン更新動作、またアドインに実装された処理の実行などによりパフォーマンスが低下することがあります。必要のないアドインは以下の手順で無効化します。

<有効なアドインの確認\>  
\[ファイル\] – \[オプション\] – \[アドイン\] で "アクティブなアプリケーション アドイン" に表示されているものが有効なアドインです。

<u>**無効化手順**</u>  

<オプション\>  
\[ファイル\] – \[オプション\] – \[アドイン\] – \[管理\] (画面下部) で無効化したいアドインの種類 (<アプリケーション名\>アドイン または COM アドイン) を選択し、\[設定\] をクリックして起動したダイアログで無効化したいアドインのチェックをオフ

補足:
*   (Excel を例に説明) Excelアドインは、VBA で実装された xlam / xla 拡張子の Excel ブック形式のアドインです。一方、COM アドインは、C# / VB.NET で作成された VSTO アドインや、C++ 等で実装されたDLL 形式のアドインを指します。
*   Web アドインと呼ばれる JavaScript で実装されるアドインはこのメニューからは確認できません。\[挿入\] – \[アドイン\] から確認、無効化しますが本記事では詳細は割愛します。
*   ほとんどのアドインはユーザーごとに登録されているため、ユーザーそれぞれで上記作業が必要です。一方、アドインによってはマシン レベルでインストールされており、一般ユーザーの権限で無効化できないものもあります。

<グループポリシー\>  
\[ユーザーの設定\] - \[Microsoft <アプリケーション名\> 2016\] – \[その他\] – \[非管理対象のアドインをすべてブロックする\]  
(<アプリケーション名\> には Excel / Word / PowerPoint のいずれかが入ります。以降の記述でも同様です。)  
設定値 : 有効  
※ 個別に無効化するのではなく、有効にするアドインを全て管理するための設定です。合わせて \[管理対象アドインの一覧\] ポリシーで有効化するアドインを指定します。

<レジストリ\>  
**Excel** **アドイン**  
以下のアドインを管理するレジストリを削除します。  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Excel\\Options  
名前 : OPENn (OPEN, OPEN1, OPEN2・・・と増加していきます)  
種類 : REG\_SZ  
値 : アドイン ファイルのフルパス (アドインフォルダに格納されている場合はファイル名のみ)  

**Word** **アドイン**  
Word アドインは起動中のプロセスでのみ読み込まれ、次回起動時には読み込まれません。このため、レジストリであらかじめ無効化するシナリオが想定されないため割愛します。

**PowerPoint** **アドイン**  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\PowerPoint\\AddIns\\<アドインファイル名 (拡張子なし)>  
名前 : AutoLoad  
種類 : REG\_DWORD  
値 : 0  

**COM** **アドイン**  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\<アプリケーション名\>\\Addins\\<アドインの名前\>  
(ユーザー単位に登録されたアドインの場合)  
キー : HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\<アプリケーション名\>\\Addins\\<アドインの名前\>  
(マシン単位に登録されたアドインの場合で OS と Office の bit が同じ場合)  
キー : HKEY\_LOCAL\_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\Office\\<アプリケーション名\>\\Addins\\<アドインの名前\>  
(マシン単位に登録されたアドインの場合で 64 bit OS と 32 bit Office の場合)  
名前 : LoadBehavior  
種類 : REG\_DWORD  
値 : 2  
<br>

## **2\.** **既定のプリンターの変更**  
印刷実行時だけでなく、ページ設定などが関連する処理では選択されたプリンターに対するアクセスが行われます。特に、Excel の改ページ ビューで保存されたブックを開くシナリオでは、ブックを開くときにこの処理が行われます。この結果、ネットワークプリンターなどでプリンターへのアクセスに時間を要する場合などにはファイルをパフォーマンスに影響を与えることがあります。また、Office は、OS の \[通常使うプリンター\] に設定されたプリンターが既定で選択されますので、\[通常使うプリンター\] を変更するとパフォーマンスを改善できます。

### **2-1. OS の\[Windows で通常使うプリンターを管理する\] をオフにする**  
この設定は、Windows 10 バージョン 1511 で追加された機能です。接続されているネットワークに応じて通常使うプリンターが選択される機能です。  
この機能が有効な場合、これまでの処理に比べて内部処理に多少時間を要しますので、常に固定のネットワークに接続する環境などでは無効化することをご検討ください。

<オプション\>  
OS の \[設定\] – \[デバイス\] – \[プリンターとスキャナー\] – \[Windows で通常使うプリンターを管理する\] のチェックをオフ

### **2-2.** **通常使うプリンターをローカルプリンターに変更する** 

<オプション\>  
\[コントロール パネル\] – \[ハードウェアとサウンド\] – \[デバイスとプリンター\] で、通常使うプリンターに設定されている (緑のチェック マークがついている) プリンターを ”Mcicrosoft XPS Document Writer" などのローカル プリンターに変更

<プロパティ\>  
プログラムの開始時に ActivePrinter プロパティをローカルプリンターに変更し、ファイル オープンやプリンタ―関連処理を行った後、最後に元に戻す

**Excel**  
[https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.activeprinter](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fexcel.application.activeprinter&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832219521&sdata=HrFp%2BcdSWaW3%2FpVDwvhoc3FVUtojBfJMUcSjxcYhSTc%3D&reserved=0)  
**Word**  
[https://docs.microsoft.com/ja-jp/office/vba/api/word.application.activeprinter](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fword.application.activeprinter&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832229518&sdata=fD3vukAn%2FLLS25q%2FZH0LcNgc8adue9Kg1RYhDIcXxD8%3D&reserved=0)  
**PowerPoint**  
[https://docs.microsoft.com/ja-jp/office/vba/api/powerpoint.application.activeprinter](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fpowerpoint.application.activeprinter&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832229518&sdata=lZtZ7RnC9yaxk%2FjT3KVCCkmi1Cjsiso90OMvCt37ijA%3D&reserved=0)
<br>

### **3\. ユーザー設定のビュー削除 (Excel 共有ブック)**  
Excel でブックの共有設定 (レガシ機能) が行われているとき、そのブックをこれまでに開いたユーザーごとの表示設定がブックに保存されています。ビュー設定の情報が増えると、ブックを開くときに読み込みのため時間がかかりますので、不要なビュー設定を削除することでパフォーマンスを改善できます。

**既存ビューの削除**  
対象の Excel ブックを開き、\[表示\] – \[ユーザー設定のビュー\] に表示されたビューから不要なものを選択して \[削除\] を実行します。

また、ビューの情報には既定で印刷設定情報が含まれます。これを含む場合、2\. で説明したようなプリンター関連の処理によってパフォーマンスが低下します。以下の設定を行うと、そのブックのビューには以後、印刷設定を含めないようにできます。ビューは保存したいが少しでもパフォーマンスを改善したい、という場合に利用できます。

**今後作成される対象ブックのビューに印刷設定を含めない**　　

<手順\>  
1\) 対象の Excel ブックを開き、\[ブックの共有 (レガシ)\] – \[詳細設定\] – \[個人用ビューに含む\] – \[印刷の設定\] のチェックをオフ  
2\) \[編集\] タブで \[新しい共同編集機能ではなく、以前の共有ブック機能を使用します。\] のチェックを一旦オフにして \[OK\] をクリック  
3\) \[新しい共同編集機能ではなく、以前の共有ブック機能を使用します。\] を再度有効にしてブックを保存  
※ \[ブックの共有 (レガシ)\] メニューは既定ではリボンに表示されません。表示する方法は以下補足に記載のリンク先をご覧ください。

<レジストリ\>  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Excel\\Options  
名前 : QFE\_Sitka  
種類 : REG\_DWORD  
値 : 1  

補足:  
現在ブックの共有機能は互換性のために残されているレガシ機能の位置づけとなり、Excel を複数のユーザーで同時編集する場合、共同編集機能を推奨しています。

共有ブック機能について  
[https://support.office.com/ja-jp/article/49b833c0-873b-48d8-8bf2-c1c59a628534](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fsupport.office.com%2Fja-jp%2Farticle%2F49b833c0-873b-48d8-8bf2-c1c59a628534&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832239510&sdata=DB5s6KZTEyhoXVIIDiWCYU78Y3glw3291SK%2FZCMrGaw%3D&reserved=0)
<br>

### **4\. 改ページ情報の非表示 (Excel)**  
Excel 改ページ設定を行うと、2\. で述べたようにプリンター関連処理でパフォーマンスが低下します。これを避けるための方法を紹介します。

### **4-1.** **保存時の標準ビュー設定**  
\[改ページ プレビュー\] や \[ページ レイアウト ビュー\] の状態でブックを保存すると、次に開いたときもこのビューの状態で開かれ、プリンター情報の取得が発生します。これを避けるため、保存時に標準ビューに変更して保存します。  

<設定\>  
\[表示\] タブ – \[ブックの表示\] – \[標準\]

<プロパティ\>  
Window.View プロパティに xlNormalView を設定  
[https://docs.microsoft.com/ja-jp/office/vba/api/excel.window.view](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fexcel.window.view&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832239510&sdata=pwSYp5RZNiu5gzvH6dpZF5pchHxBlppSgYuEB5io1pw%3D&reserved=0)

### **4-2. Worksheet.DisplayPageBreaks** **の設定**
一旦改ページを設定すると、標準ビューでも、改ページ位置が表示されるようになります。この状態でもパフォーマンス低下が生じる場合があります。主にプログラムからの処理が遅い場合に、初めに設定すると有効なプロパティです。

<プロパティ\>  
Worksheet.DisplayPageBreaks プロパティ を False に設定  
[https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.displaypagebreaks](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fexcel.worksheet.displaypagebreaks&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832249506&sdata=04OYmavm2zYeMu5FldVB4rLZt2yV5RUpklZDO5bCRVs%3D&reserved=0)
<br>
  
## **5\.** **ファイル・フォルダ履歴削除**  
Office の \[開く\] や \[名前を付けて保存\] メニューに表示される最近開いたファイルやフォルダーの情報が多く残されていると、その読み込みに時間を要します。情報を減らすことでパフォーマンスを改善できます。

<オプション\>  
**現時点での履歴情報をクリアする**  
個別削除: \[ファイル\] タブ – \[開く\] でファイルやフォルダーの履歴から削除しても良い項目を右クリックし \[一覧から削除\]  
一括削除: \[ファイル\] タブ – \[開く\] でファイルやフォルダーの履歴で任意のファイル/フォルダを右クリックし \[固定されていない項目をクリア\]  
※ あらかじめ \[一覧にピン留め\] しておくと、一括削除時に残すことができます。

**今後の履歴情報保持数を変更する**  
<u>ファイル</u>  
\[詳細設定\] – \[表示\] – \[最近使ったブックの一覧に表示するブックの数\] を少ない数に変更 (0 も設定可能) ※ Excel の例  
("ブック" の部分は、Word の場合は "文書"、PowerPoint は "プレゼンテーション”)  
<u>フォルダ</u>  
\[詳細設定\] – \[表示\] – \[最近使ったフォルダーの一覧から固定表示を解除するフォルダーの数\] を少ない数に変更

<グループポリシー\>  
\[ユーザーの設定\] - \[Microsoft <アプリケーション名\> 2016\] – \[<アプリケーション名\> のオプション\] - \[詳細設定\]  
– \[最近使ったブック (ドキュメント、プレゼンテーション) の一覧に表示するブック(ドキュメント、プレゼンテーション) の数\]  
– \[最近使用したフォルダーの一覧に表示するフォルダーの数\]  
設定値 : 表示したい数 (最小値は 0)

<レジストリ\>  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\<アプリケーション名\>\\Options\\File MRU (ファイル)  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\<アプリケーション名\>\\Options\\Place MRU (フォルダ)  

**現時点での履歴情報をクリアする場合は以下を削除**  
名前 : Item n (n には 1 ~ の連番が入ります)  
種類 : REG\_SZ  
値 : ファイルまたはフォルダのフルパス  

**今後の履歴情報保持数を変更する場合は以下の値を変更**  
名前 : Max Display  
種類 : REG\_DWORD  
値 : 履歴を保持したい数  
 <br> 

##  **6\.** **ファイル検証の無効化**  
バイナリ形式ファイル (xls / doc / ppt など) を開くときには、ファイル検証処理が行われます。ファイル検証処理は、バイナリ形式であるこれらのファイルのフォーマットが正しいかをチェックし、不正なファイルを検出する仕組みです。(Office 2007 形式 (xlsx / docx / pptx) 等に対してはファイル検証は行われません。) この検証に時間を要する場合、セキュリティが十分担保されている環境であればファイル検証を無効化してパフォーマンスを改善できます。

参考) Office 2016 の Office ファイル検証を使った、ファイル形式攻撃の防止  
[https://docs.microsoft.com/ja-jp/DeployOffice/security/prevent-file-format-attacks-by-using-file-validation-in-office](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2FDeployOffice%2Fsecurity%2Fprevent-file-format-attacks-by-using-file-validation-in-office&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832249506&sdata=OJKjGU%2FbD8vIdClW0KqjaVmRRS%2F3U%2F6IBjo0V8MWydM%3D&reserved=0)

セキュリティ上、基本的にはファイル検証機能の無効化は推奨しません。無効にする場合はバイナリ形式ファイルがインターネットなどの安全でない場所から入手されないことを十分に確認してください。

<グループポリシー\>  
\[ユーザーの設定\] - \[Microsoft <アプリケーション名\> 2016\] – \[<アプリケーション名\> のオプション\] – \[セキュリティ\] – \[ファイル検証機能をオフにする\]
設定値 : 有効

<レジストリ\>  
キー : HKEY\_CURRENT\_USER\\Software\\Policies\\Microsoft\\Office\\16.0\\<アプリケーション名\>\\Security\\FileValidation
  ※ Policies 配下のみ有効です。  
エントリ名 : EnableOnLoad  
型 : REG\_DWORD  
値 : 0  

<プロパティ\>  
Application.Filevalidation プロパティを false に設定

**Excel**  
[https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.filevalidation](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fexcel.application.filevalidation&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832249506&sdata=6wC1wtjSQOox%2BsMhfmwHe41mejqY8fjNn3gWKjEoqCA%3D&reserved=0)  
**Word**  
[https://docs.microsoft.com/ja-jp/office/vba/api/word.application.filevalidation](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fword.application.filevalidation&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832259499&sdata=x1tV%2BxToA10xKanVXM1NIDDtFJq3eTDooZ3jlGnyHSA%3D&reserved=0)  
**PowerPoint**  
[https://docs.microsoft.com/ja-jp/office/vba/api/powerpoint.application.filevalidation](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice%2Fvba%2Fapi%2Fpowerpoint.application.filevalidation&data=02%7C01%7Ckanakamu%40microsoft.com%7C88d7890adb404767999008d7796fe66c%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637111393832259499&sdata=LmPSMmVE0OfqX0gFTU6o1F5%2BKTed9uffv9CA3nVc6ZU%3D&reserved=0)
<br>

今回の投稿は以上です。
<br>
**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**