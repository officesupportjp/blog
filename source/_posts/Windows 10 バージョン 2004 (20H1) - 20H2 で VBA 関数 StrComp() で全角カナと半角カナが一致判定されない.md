---
title: Windows 10 バージョン 2004 以降で VBA 関数 StrComp() で全角カナと半角カナが一致判定されない
date: 2020-11-12
lastupdate: 2021-10-28
---

<span style="color:#ff0000">**2020/11/24 Update**</span>  
<span style="color:#339966">「4. その他の発生シナリオ」セクションを追加し、Excel での検索・置換 (Find / Replace) での発生を追記しました。</span>  

<span style="color:#ff0000">**2021/10/22 Update**</span>  
<span style="color:#339966">現象発生 OS と回避策に記載の対応を行った場合の既知の問題を追記しました。</span>  

<span style="color:#ff0000">**2021/10/18 Update**</span>  
<span style="color:#339966">Windows 11 では修正済みであることを追記しました。</span>  

  
  

こんにちは、Office サポート チームの中村です。

先日、以下の記事で同じ原因により Access フォームで発生する問題について記載していますが、VBA 関数 StrComp() も影響を受けることが分かったため、本記事にて詳細をご案内します。  
  
**関連記事**

Windows 10 バージョン 2004 以降で 半角カナのフォームを含んだ Access ファイルでエラーが発生する  
[https://officesupportjp.github.io/blog/Windows 10 バージョン 2004 以降で 半角カナのフォームを含んだ Access ファイルでエラーが発生する/](https://officesupportjp.github.io/blog/Windows%2010%20%E3%83%90%E3%83%BC%E3%82%B8%E3%83%A7%E3%83%B3%202004%20(20H1)%20-%2020H2%20%E4%B8%8A%E3%81%A7%20%E5%8D%8A%E8%A7%92%E3%82%AB%E3%83%8A%E3%81%AE%E3%83%95%E3%82%A9%E3%83%BC%E3%83%A0%E3%82%92%E5%90%AB%E3%82%93%E3%81%A0%20Access%20%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%A7%E3%82%A8%E3%83%A9%E3%83%BC%E3%81%8C%E7%99%BA%E7%94%9F%E3%81%99%E3%82%8B/)  

#### **1\. 現象**

VBA には、StrComp() という 2 つの文字列が一致するかを比較する関数が用意されています。

StrComp 関数  
[https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/strcomp-function](https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/strcomp-function)  

例えば以下のサンプル コードのように、濁点または半濁点を含む文字列を、一方は全角、一方半角で比較すると、以前の OS では一致とみなされ 0 が返っていました。

しかし、Windows 10 バージョン 2004 以降の OS 上で実行すると、不一致を示す -1 が返ります。  

<サンプル コード\>

    Sub Sample()
        MsgBox StrComp("サンプル", "ｻﾝﾌﾟﾙ", vbTextCompare)
    End Sub

  
本現象は、Office バージョンに関係なく発生します。  
  

#### **2\. 原因**

Windows 10 バージョン 2004 で、OS の自然言語処理を担う NLS (National Language Support) バージョンが更新されました。

<OS バージョンと NLS バージョン\>  

|OS バージョン|NLS バージョン|
|:---|:---|
|Windows 10 バージョン 1909|6.2|
|Windows 10 バージョン 2004 以降|6.3|
|Windows 11|6.4|


  
この新しいバージョン 6.3 の NLS ソート テーブルにて、濁点および半濁点を含むカナ文字について、半角と全角が異なる文字と認識される問題の影響で現象が発生します。上記の例の場合、"プ" (全角) と "ﾌﾟ" (半角) が違う文字と判定されます。

この問題は、OS の更新プログラムについての公開情報でも既知の問題として記載されています。

Current status of Windows 10, version 2004 and Windows Server, version 2004  
[https://support.microsoft.com/en-us/topic/windows-10-update-history-24ea91f4-36e7-d8fd-0ddb-d79d9d0cdbda](https://support.microsoft.com/en-us/topic/windows-10-update-history-24ea91f4-36e7-d8fd-0ddb-d79d9d0cdbda)


Windows 10 update history - Updates for Windows 10, version 20H2 and Windows Server, version 20H2  
[https://support.microsoft.com/en-us/topic/windows-10-update-history-7dd3071a-3906-fa2c-c342-f7f86728a6e3](https://support.microsoft.com/en-us/topic/windows-10-update-history-7dd3071a-3906-fa2c-c342-f7f86728a6e3)


Windows 10 update history - Updates for Windows 10, version 21H1
[https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11](https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11)


<公開情報より該当の記載を抜粋\>

|**Symptom**|**Workaround**|
|:---|:---|
|Certain Japanese half-width Katakana and full-width Katakana characters that have a consonant mark aren’t interpreted as the same character. When you use the **CompareStringEx()** function with the _NORM\_IGNOREWIDTH_ flag to compare them, these characters are evaluated as different because of an issue in the sorting rule**.** This issue affects all the updates starting on June 9, 2020 for Windows 10, version 2004.|Currently, there is no workaround for this issue.|

※ 次期リリース予定の Windows 10 バージョン 21H2 においても引き続きこの問題は発生します。
  

#### **3\. 対応状況 / 回避策**

Windows 11 では、NLS バージョン 6.4 に更新され、この問題は修正されています。  
Windows 10 の現象が発生するバージョンでは、以下のいずれかの回避策で現象を回避できます。
  

**a.** **比較対象の文字列を全角、または半角に統一する**

実際の VBA コードでは、値は VBA 定数、Excel シート上のセルの値や Access データベースのテーブルの値、ユーザーフォームなど、様々な場所から取得されることが想定されます。VBA コード取得元となる文字列を探して全角 / 半角が統一されるよう変更します。  
  

**b. StrConv** **関数で全角、または半角に統一してから比較する**

複雑なコードで比較対象文字列の取得元を特定することが困難な場合や、ユーザーが入力するため制御が難しい場合などには、StrConv 関数を用いて全角または半角に変換してから、StrComp 関数の比較に使用することが検討できます。

StrConv 関数  
[https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/strconv-function](https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/strconv-function)

  
<回避案 b サンプル コード\>

以下の例では、比較する文字列を両方とも全角カナに変換しています。

    Sub Sample()
        str1 = "サンプル"
        str2 = "ｻﾝﾌﾟﾙ"
        MsgBox StrComp(StrConv(str1, vbWide), StrConv(str2, vbWide), vbTextCompare)
    End Sub
    

  
  
  
基本的には、上記の a. / b. いずれかの方法での対応を強く推奨します。

ただ、ファイルや VBA の変更が難しい場合などには、クライアントの環境設定による以下の回避策もあります。いずれも、Office の他の動作や、OS 全体に影響を及ぼす設定となるため、以下 c. / d. 案を利用する場合は、影響に十分にご留意ください。  
  

**c. Office** **モジュールを互換モードに設定する**

Office アプリケーションの実行モジュールを互換モードで動作するよう設定すると、古い NLS バージョンを使用するため現象が発生しません。

主な Office アプリケーションの実行モジュール名は以下になります。

Excel : EXCEL.EXE  
Word : WINWORD.EXE  
PowerPoint : POWERPNT.EXE  
Access : MSACCESS.EXE

互換モードに変更したアプリケーションでの全ての動作が互換モードになるため、Windows 10 OS で新たにサポートされた機能を前提とした機能が動作しなくなるなどの影響が生じる場合があります。(例 : Excel の \[オプション\] – \[全般\] にある "複数ディスプレイを使用する場合" で \[表示を優先した最適化\] を選択することができなくなります)  
  

**d. レジストリを使用して NLS バージョンを以前のバージョン (6.2) に変更する**

システム全体で影響を受けるため、Office 以外のアプリケーションも以前の NLS バージョンを使用するようになります。アプリケーションの実装やビジネスロジックによっては、ソートの整合が崩れたり、データが不正となる可能性があります。

c. / d. の具体的な設定方法や注意点については、既に公開された Access フォームの現象をご案内する以下のフォーラム記事で説明していますので、こちらを参照してください。

Windows 10 バージョン 2004 以降で 半角カナのフォームを含んだ Access ファイルでエラーが発生する  
[https://officesupportjp.github.io/blog/Windows 10 バージョン 2004 以降で 半角カナのフォームを含んだ Access ファイルでエラーが発生する/](https://officesupportjp.github.io/blog/Windows%2010%20%E3%83%90%E3%83%BC%E3%82%B8%E3%83%A7%E3%83%B3%202004%20(20H1)%20-%2020H2%20%E4%B8%8A%E3%81%A7%20%E5%8D%8A%E8%A7%92%E3%82%AB%E3%83%8A%E3%81%AE%E3%83%95%E3%82%A9%E3%83%BC%E3%83%A0%E3%82%92%E5%90%AB%E3%82%93%E3%81%A0%20Access%20%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%A7%E3%82%A8%E3%83%A9%E3%83%BC%E3%81%8C%E7%99%BA%E7%94%9F%E3%81%99%E3%82%8B/)

※ 参照箇所  
c 案 : 3-1. Access アプリケーション観点 - Access の EXE (MSACCESS.EXE) について、"Windows 8" の互換設定を行う  
d 案 : 3-3. OS 観点 – Windows OS の NLS バージョンを 6.3 から 6.2 に変更する  
  

**注意**
この d 案の設定を行うと、Windows のライセンス認証が行われない場合があることを確認しています。ライセンス認証を実行する際はこの変更を元に戻してから実行してください。


#### **4\. その他の発生シナリオ**

  
ここまで、一般に利用頻度の比較的高い StrComp() での発生シナリオを中心に説明してきましたが、内部的に文字の比較を行う以下の機能でも同様の問題による現象が発生します。  
  
・Range.Find() / Range.FindNext() / Range.FindPrevious() メソッドによる検索や Range.Replace() メソッドによる置換で、濁点・半濁点を含む半角カナと全角カナが一致判定されません。  
※ 手動での検索や置換も同様です。

今回の投稿は以上です。  

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**