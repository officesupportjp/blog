---
title: Access でメモリ不足エラーが発生する現象の対処策について
date: '2019-03-09'
id: cl0m7ghoi000fk8vs8yjpdr8b
tags:
  - Access
  - パフォーマンス

---

(※ 2017 年 2 月 16 日に Japan Office Support Blog に公開した情報のアーカイブです。)

こんにちは、Office サポート チームです。

Access や Access ODBC/OLEDBを使用したプログラムでは、mdb/accdb ファイル上のテーブルへレコードを追加/変更/削除を行うことが可能です。

本稿では、Access および Access ODBC/OLEDBを使用したプログラムから、mdb/accdb ファイル内のテーブルにレコードの追加/更新/削除を繰り返し実施した際にメモリ不足が生じる現象の対処方法をご案内します。

### **現象**

Access および Access ODBC/OLEDBを使用したプログラムで mdb/accdb ファイル内のテーブルに連続したレコードを追加/更新/削除を行うと、Access および Access ODBC/OLEDBを使用したプログラムでメモリ不足エラーが発生することがあります。

### **原因**

本現象は、トランザクション処理を行うためのページロックやメモリバッファが不足した場合や、Access Database Engine の処理に必要なメモリが解放されないため、アプリケーションで利用する必要なメモリが確保できない場合に発生します。

### **対処方法**

この問題を解決するには、必要に応じて、指定された順序で以下の方法を実行します。

**方法 1. レジストリ WorkingSetSleep を設定する**  
**方法 2. レジストリ MaxLocksPerFile を設定する**  
**方法 3. レジストリ MaxBufferSize を設定する**  
**方法 4 Access.exe プロセスの CPU マッピングを 1 つに限定する**  
**方法 5. プログラムを再起動する**  
**方法 6. 一度に Insert/Update/Delete するレコード数を減らす**

<u>レジストリ エディターについて</u>  
レジストリ エディターの誤った使用は、システム全般に渡る重大な問題を引き起こす可能性があります。  
こうした問題を解決するためには、Windows をインストールしなおさなければいけません。  
Microsoft では、レジストリ エディターを使用することによって引き起こされた障害の解決については一切保証しておりません。  
レジストリ エディターを使用する場合には、お客様の責任において使用してください。

**備考 : この記事内に記載しているレジストリを変更するためには管理者権限が必要です**。

### **方法 1.レジストリ WorkingSetSleep を設定する**

バージョン 2010 以降の Access では、複雑なクエリを実行した際に「システム リソースの超過」というメモリ不足エラーが表示される場合があります。  
この現象を回避するには以下のレジストリ WorkingSetSleep を作成します。

キー: 作成する場所は後述の「各バージョンのキー」をご参照ください。  
型 : DWORD  
名前 : WorkingSetSleep  
値 : 1

補足:  
WorkingSetSleep に 1 をセットすると、Access Database Engineで利用するメモリを確保する際に 50 ミリ秒の Sleep 処理を呼び出すように動作が変更されます。  
この 50 ミリ秒の待機中に未使用になったメモリの解放処理が実行されることで、メモリ不足の状態を避けることができ、現象が回避できる可能性があります。  
なお、Sleep 処理が呼び出されるため、クエリの実行時間は長くなります。

<u>キーの各バージョンの場所</u>

|バージョン|インストール形式|アーキテクチャ|レジストリキー|備考|
|---|---|---|---|---|
|Access 2010|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\Software\\Microsoft\\Office\\14.0\\Access Connectivity Engine\\Engines|[KB2760394](https://support.microsoft.com/ja-jp/help/2760394/) が適用されている必要があります。|
|Access 2010|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\14.0\\Access Connectivity Engine\\Engines|[KB2760394](https://support.microsoft.com/ja-jp/help/2760394/) が適用されている必要があります。|
|Access 2013|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines|[クイック実行形式 (C2R) と Windows インストーラー形式 (MSI) を見分ける方法](https://blogs.technet.microsoft.com/officesupportjp/2016/09/08/howto_c2r_or_msi)|
|Access 2013|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines||
|Access 2013|C2R|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\<br />REGISTRY\\MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines||
|Access 2013|C2R|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\<br />REGISTRY\\MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines||
|Access 2016|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\MICROSOFT\\Office\\16.0\\Access Connectivity Engine\\Engines||
|Access 2016|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\MICROSOFT\\Office\\16.0\\Access Connectivity Engine\\Engines||
|Access 2016|C2R|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\<br />REGISTRY\\MACHINE\\SOFTWARE\\Microsoft\\Office\\16.0\\Access Connectivity Engine\\Engines||
|Access 2016|C2R|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\<br />MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\16.0\\Access Connectivity Engine\\Engines||


以上の手順で問題が解決しない場合は、作成した WorkingSetSleep のレジストリを削除するか、WorkingSetSleep に 0 を設定し、元の設定に戻します。  
設定を戻した後、方法 2 に進みます。

### **方法 2. レジストリ MaxLocksPerFile を設定する**

トランザクションに必要なページロックの数が既定値を超えることでメモリ不足のエラーが発生する場合があります。

多数のトランザクション処理中にエラー メッセージ "ファイルの共有ロック数が制限を超えています" が表示される  
[https://support.microsoft.com/ja-jp/help/815281](https://support.microsoft.com/ja-jp/help/815281)

You may receive a "There isn't enough disk space or memory" error message when you perform an operation on an Access table  
[https://support.microsoft.com/en-us/help/286153](https://support.microsoft.com/en-us/help/286153)

キー: 作成する場所は後述の「各バージョンのキー」をご参照ください。  
型 : DWORD  
名前 : MaxLocksPerFile  
値 : 200000 (10 進数)

※ 既定値 : 9500 (10 進数)  
※ MaxLocksPerFile に設定可能な値 (理論値) : 0x00000000 ～ 0xffffffff ※ 16進数  
(0xffffffff の10 進数 : 4294967295)

なお、大量のトランザクションを処理している場合は、200000 よりも大きな値を設定して、何パターンかお試しください。  
MaxLocksPerFile を変更して改善する場合は、可能であれば、コミットするタイミングを変更しトランザクションで処理するレコード数を減らすこともご検討ください。

<u>キーの各バージョンの場所</u>

|バージョン|インストール形式|アーキテクチャ|レジストリキー|備考|
|---|---|---|---|---|
|Access 2010|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\Software\\Microsoft\\Office\\<br />14.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2010|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\<br />Microsoft\\Office\\14.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2013|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\<br />Access Connectivity Engine\\Engines\\ACE|[クイック実行形式 (C2R) と Windows インストーラー形式 (MSI) を見分ける方法](https://blogs.technet.microsoft.com/officesupportjp/2016/09/08/howto_c2r_or_msi)|
|Access 2013|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\<br />Office\\15.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2013|C2R|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun<br />\\REGISTRY\\MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2013|C2R|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun<br />\\REGISTRY\\MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2016|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\MICROSOFT\\Office<br />\\16.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2016|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\<br />MICROSOFT\\Office\\16.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2016|C2R|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\<br />REGISTRY\\MACHINE\\SOFTWARE\\Microsoft\\Office\\16.0\\Access Connectivity Engine\\Engines\\ACE||
|Access 2016|C2R|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\<br />MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\16.0\\Access Connectivity Engine\\Engines\\ACE||

以上の手順で問題が解決しない場合は、方法 3 に進みます。

### **方法 3. レジストリ MaxBufferSize の設定**

トランザクションを実行中、大量のレコードの変更をバッファに保存する必要があるため、大量のデータを処理する場合、バッファが不足してメモリ不足が発生している可能性があります。  
MaxBufferSize を変更後も改善しない場合は、方法 2 の MaxLocksPerFile の設定も一緒に変更した組み合わせをお試しください。  
（例 : MaxLocksPerFile を 200000 にして、かつ、MaxBufferSize も102400 にする、等）

キー: 作成する場所は後述の「各バージョンのキー」をご参照ください。  
型 : DWORD  
名前 : MaxBufferSize  
値 : 102400 (10 進数)

※ 既定値 : 0(10 進数)  
※ MaxBufferSizeに設定可能な値 (理論値) : 0x00000000 ～ 0xffffffff ※ 16進数  
(0xffffffff の10 進数 : 4294967295)

なお、大量のトランザクションを処理している場合は、102400よりも大きな値を設定して、何パターンかお試しください。  
<u>キーの各バージョンの場所</u>

|バージョン|インストール形式|アーキテクチャ|レジストリキー|
|---|---|---|---|
|Access 2010|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\Software\\Microsoft\\Office\\<br />14.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2010|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\<br />Microsoft\\Office\\14.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2013|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\<br />15.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2013|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\<br />Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2013|C2R|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\<br />ClickToRun\\REGISTRY\\MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2013|C2R|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\REGISTRY<br />\\MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2016|MSI|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\MICROSOFT\\Office\\<br />16.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2016|MSI|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Wow6432Node\\<br />MICROSOFT\\Office\\16.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2016|C2R|OS と Office が同じ bit|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\<br />REGISTRY\\MACHINE\\SOFTWARE\\Microsoft\\Office\\16.0\\Access Connectivity Engine\\Engines\\ACE|
|Access 2016|C2R|OS と Office の bit が異なる|HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\<br />REGISTRY\\MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\16.0\\Access Connectivity Engine\\Engines\\ACE|

以上の手順で問題が解決しない場合は、MaxLocksPerFile および MaxBufferSize を既定値に戻します。  
設定を戻した後、方法 4 に進みます。

### **方法 4 Access.exe プロセスの CPU マッピングを 1 つに限定する**

Access はマルチスレッドで動作することを想定していないエンジンを利用しています。  
この動作が影響してメモリ不足が生じる場合があります。

回避策 :  
1\. 現象が発生している Access ファイルを開きます。  
2\. \[タスク マネージャー\] を起動します。タスクマネージャの \[プロセス\] タブで “MSACCESS.EXE” のプロセスを右クリックします。  
3\. \[関係の設定\] メニューを選択します。  
4\. \[プロセッサの関係\] ダイアログ ボックスで、任意の一つの CPU のみチェックをオンにし、その他の CPU のチェックをオフにします。  
5\. OK でダイアログを閉じます。

複数コアの CPU 環境で実行中の Access が応答なしになる  
[https://support.microsoft.com/ja-jp/help/2646068](https://support.microsoft.com/ja-jp/help/2646068)

以上の手順で問題が解決しない場合は、方法 5 に進みます。

### **方法 5. プログラムを再起動する**

メモリ不足が発生している Access もしくは Access ODBC/OLEDBを使用したプログラムを終了し、再起動します。

以上の手順で問題が解決しない場合は、方法 6 に進みます。

### **方法 6. 一度に Insert/Update/Delete するレコード数を減らす**

トランザクションで Insert/Update/Delete するレコードの数を減らします。

**本情報の内容 (添付文書、リンク先などを含む) は作成日時点でのものであり、予告なく変更される場合があります。**