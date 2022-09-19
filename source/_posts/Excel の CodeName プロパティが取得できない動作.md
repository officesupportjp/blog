---
title: Excel の CodeName プロパティが取得できない動作
date: '2019-03-01'
id: cl0mc9o14000tb0vs85lvc2xs
tags:
  - オートメーション

---

(※ 2017 年 2 月 9 日に Japan Office Developer Support Blog に公開した情報のアーカイブです。)

こんにちは、Office 開発 サポート チームの中村です。

今回は、Excel の CodeName プロパティの動作について、開発時に気付きづらい動作上の注意点を記載します。この状況に直面した開発者の皆様へのヒントになれば幸いです。

Excel をオートメーションするとき、ブックやシートを特定するために、CodeName プロパティを使用できます。CodeName プロパティは、Visual Basic Editor (VBE) のプロパティ ウィンドウの \[(オブジェクト名)\] で確認できる値です。

Excel の CodeName プロパティは、以下の 3 つのオブジェクトに用意されています。

タイトル : Workbook.CodeName プロパティ (Excel)  
アドレス : [https://msdn.microsoft.com/ja-jp/library/office/ff195162.aspx](https://msdn.microsoft.com/ja-jp/library/office/ff195162.aspx)

タイトル : Worksheet.CodeName プロパティ (Excel)  
アドレス : [https://msdn.microsoft.com/ja-jp/library/office/ff837552.aspx](https://msdn.microsoft.com/ja-jp/library/office/ff837552.aspx)

タイトル : Chart.CodeName プロパティ (Excel)  
アドレス : [https://msdn.microsoft.com/ja-jp/library/office/ff835278.aspx](https://msdn.microsoft.com/ja-jp/library/office/ff835278.aspx)

CodeName プロパティを使用すると、ユーザーが任意に変更できるシート名などではなく、通常開発者しか目にしない VBProject の名前でシート オブジェクトなどを取り扱うことができます。例えば、あるシート オブジェクト (例 : シート名 「Sheet1」/ CodeName 「Code1」) に記載した VBA コード (例 : SampleCode) にアクセスするとき、「Worksheets("Sheet1").SampleCode」 と書くと、ユーザーがシート名を変更するとアクセスできなくなるため、\[Code1.SampleCode\] のように実装するといった用途で利用できます。

ただし、CodeName プロパティは、そのオブジェクトが作成されてから、VBProject へのアクセスが発生しないと取得することができません。

例えば、以下のようなサンプル コードを Excel の VBA として記述し、**必ず VBE をいったん閉じてから**、\[開発\] タブ – \[マクロ\] から Sample マクロを実行してみるとこの動作を再現することができます。

以下のサンプルでは、新しいブックと、その中に新しいワークシート、グラフシートを作成し、それぞれの CodeName を順にメッセージボックスに表示します。VBE ウィンドウが開かれた状態だと CodeName が表示されますが、VBE を閉じてから実行すると、何も表示されません。

```
Sub Sample()
    ' VBE を閉じていると、各 MsgBox では、空のメッセージボックスが表示されます
    Set newBook = Workbooks.Add
    MsgBox newBook.CodeName
    
    newBook.Worksheets.Add
    MsgBox newBook.ActiveSheet.CodeName
    
    newBook.Charts.Add
    MsgBox newBook.ActiveChart.CodeName
End Sub
```


#### **回避策**

**1\. CodeName** **プロパティを取得できるようにする**

先述の通り、CodeName プロパティは VBProject へのアクセスが生じると取得できるようになります。先程の例では、VBE を起動しておくことによって VBProject へのアクセスが生じ、CodeName が取得できていました。

これをコードから実現する場合、対象のオブジェクトが作成されてから CodeName を取得するまでの間に、以下のプロパティを用いて当該 VBProject にアクセスすることで回避できます。

タイトル : Workbook.VBProject プロパティ (Excel)  
アドレス : [https://msdn.microsoft.com/ja-jp/library/office/ff194737.aspx](https://msdn.microsoft.com/ja-jp/library/office/ff194737.aspx)

ただし、プログラムからこのVBProject へアクセスするには、Excel の \[オプション\] – \[セキュリティ センター\] – \[セキュリティ センターの設定\] – \[マクロの設定\] – \[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する\] を有効にする必要があります。このため、セキュリティ リスクや、運用上オプションを強制できるかどうかを十分に検討してください。

**2\. 他の方法でオブジェクトにアクセスする**

例えばワークシートの場合、Worksheets(1) のようにインデックスでアクセスしたり、Set newSheet = Worksheets.Add のように初めにシートを扱ったときにオブジェクトに格納しておき、これを用いてアクセスするなど、他の方法でオブジェクトにアクセスします。

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**