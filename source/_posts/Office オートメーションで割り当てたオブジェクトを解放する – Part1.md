---
title: Office オートメーションで割り当てたオブジェクトを解放する – Part1
date: 2019-02-28
lastupdate: 2021-04-28
---

(※ 2012 年 2 月 9 日に Japan Office Developer Support Blog に公開した情報のアーカイブです。)

こんにちは、Office 開発系サポート 森　健吾 (kenmori) です。

今回の投稿では、Office オートメーションの実装コードで割り当てたオブジェクトを解放するというテーマにて記載いたします。  
.NET Framework 上で動作するカスタム アプリケーションにおいて、Office オートメーションで処理を実装する場合には割り当てたオブジェクトを確実に解放することをお勧めします。

<span style="color:#ff0000">**2016/12/2 Update**</span>  
<span style="color:#339966">サンプル コードを中間オブジェクトも解放するよう、より適切な形に変更しました。</span>

<span style="color:#ff0000">**2017/3/28 Update**</span>  
<span style="color:#339966">アプリケーション終了前に一部のオブジェクトを解放するよう、サンプル コードをさらに適切な形に変更しました。</span>

これは、Office が内部的に OLE や DDE などを通じて実施するオブジェクト インスタンス制御 (<span style="color:#ff9900">※</span>) によって CLR 上で解放漏れ (または解放待ち) のオブジェクトが誤って参照されてしまい、カスタム アプリケーション側の動作に様々な予期せぬ影響を与えることにあります。

<span style="color:#ff9900">※ OLE や DDE などを通じて実施するオブジェクト インスタンス制御や、解放漏れによる影響等については、別途記載を予定しております。</span>

今回は、詳細な上記の詳細な理由や解放しない際の影響等は省略し、オブジェクト解放の正しい実装方法について記載します。

### 1\. 割り当てたオブジェクトを解放する

Office オートメーションで割り当てたオブジェクトについては、自分で解放処理を記載して必ず破棄される必要があります。

例えば、以下の Windows アプリケーションにおけるサンプル コードでは VB.NET で Dim 宣言を使用して Application, Workbooks, Workbook, Worksheet オブジェクトを割り当てております。割り当てたオブジェクトは Marshal.ReleaseComObject メソッドを使用して必ず解放するようにしております。

また、以下のサンプル コードの Workbooks オブジェクトのような中間オブジェクトは、「Dim workbook As Excel.Workbook = app.Workbooks.Add()」のように、中間オブジェクトを変数に格納しない形で記述されることもあるかと思いますが、COM オブジェクトの解放の観点では、これは適切な実装ではありません。

このように実装した場合も、暗黙的に Workbooks オブジェクトは作成されており、このオブジェクトも適切に解放する必要があります。したがって、後で Marshal.ReleaseComObject で解放できるよう、以下のサンプル コードのように変数に受けておくのが適切な実装です。

アプリケーションの終了時には、Excel からの COM オブジェクトの参照解放処理が行われます。このタイミングまでに適切に COM オブジェクトが解放されていないと、ガベージコレクトのタイミングによっては予期せぬエラーが生じる場合があります。このため、アプリケーションの終了前にまずガベージコレクトを実行します。

さらに、Application インスタンスを解放する際なのですが、Marshal.ReleaseComObject メソッドを使用して参照カウンタをデクリメントしただけでは、プロセスが終了することを保証できませんので、GC.Collect メソッドでガベージコレクトを強制してオブジェクトを解放しています。

注意 : なお、サンプル コードの簡素化のため、敢えて例外処理等は記載しておりません。

**VB.NET**

```
Imports System.Runtime.InteropServicesPublic

Class Form1

  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    Dim app As Excel.Application = New Excel.Application()  
    Dim workbooks As Excel.Workbooks = app.Workbooks  
    Dim workbook As Excel.Workbook = workbooks.Add()  
    Dim worksheet As Excel.Worksheet = workbook.Sheets(1)  
    Dim range As Excel.Range = worksheet.Range("A1")

    range.Value = TextBox1.Text  
    workbook.SaveAs("C:\Out\a.xlsx")  
    workbook.Close()

    ' アプリケーションの終了前に破棄可能なオブジェクトを破棄します。  
    Marshal.ReleaseComObject(range)  
    range = Nothing

    Marshal.ReleaseComObject(worksheet)  
    worksheet = Nothing

    Marshal.ReleaseComObject(workbook)  
    workbook = Nothing

    Marshal.ReleaseComObject(workbooks)  
    workbooks = Nothing

    ' アプリケーションの終了前にガベージ コレクトを強制します。  
    GC.Collect()  
    GC.WaitForPendingFinalizers()  
    GC.Collect()

    ' アプリケーションを終了します。  
    app.Quit()

    ' Application オブジェクトを破棄します。  
    Marshal.ReleaseComObject(app)  
    app = Nothing

    ' Application オブジェクトのガベージ コレクトを強制します。  
    GC.Collect()  
    GC.WaitForPendingFinalizers()  
    GC.Collect()

  End Sub

End Class
```

なお、ガベージ コレクトの処理は若干冗長に見えますが、上記がベスト プラクティスになります。詳細は以下のサイトをご確認ください。

タイトル :  第 5 章 「マネージ コード パフォーマンスの向上」  
アドレス : [http://msdn.microsoft.com/ja-jp/library/ms998547.aspx](http://msdn.microsoft.com/ja-jp/library/ms998547.aspx)  
参考箇所 : GC.Collect の呼び出しを避ける

以下に C# の場合のコーディングも併記します。

**C#**
```
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace winExcelSample{
  public partial class Form1 : Form  {
    public Form1() {
      InitializeComponent();
    }
    private void button1_Click(object sender, EventArgs e){
      Excel.Application app = new Excel.Application();
      Excel.Workbooks workbooks = app.Workbooks;
      Excel.Workbook workbook = workbooks.Add(Type.Missing);
      Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
      Excel.Range range = worksheet.get_Range("A1", Type.Missing);

      range.set_Value(Type.Missing, textBox1.Text);
      workbook.SaveAs(@"C:\Out\a.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
      workbook.Close(Type.Missing, Type.Missing, Type.Missing);
      
      // アプリケーションの終了前に破棄可能なオブジェクトを破棄します。
      Marshal.ReleaseComObject(range);
      range = null;
      
      Marshal.ReleaseComObject(worksheet);
      worksheet = null;
      
      Marshal.ReleaseComObject(workbook);
      workbook = null;
      
      Marshal.ReleaseComObject(workbooks);
      workbooks = null;
      
      // アプリケーションの終了前にガベージ コレクトを強制します。
      GC.Collect();
      GC.WaitForPendingFinalizers();
      GC.Collect();
      
      // アプリケーションを終了します。
      app.Quit();
      
      // Application オブジェクトを破棄します。
      Marshal.ReleaseComObject(app);
      app = null;
      
      // Application オブジェクトのガベージ コレクトを強制します。
      GC.Collect();
      GC.WaitForPendingFinalizers();
      GC.Collect();
    }
  }
}
```

なお、コンソール アプリケーション等では、処理が一通り実行された後、プロセスが終了する際に CLR ランタイムが終了するのに合わせて、Application オブジェクトの解放および EXCEL.EXE プロセスの終了が実施されます。そのため、解放漏れがあったとしても、プログラム終了時に解放されるため、ほとんどの場合大きな影響がない傾向があります。

### 2\. イベントで取得したオブジェクトを解放する

Office オートメーションでイベント ハンドラを追加した場合においても、引数として取得されるオブジェクトは解放する必要があります。上述の例と同様に Marshal.ReleaseComObject メソッドで解放を実施しております。

**VB.NET**

```
Public Sub Workbook_Open(ByVal Wb As Excel.Workbook)
  Try
    ' 処理
  Finally
    ' オブジェクトを破棄します。
  Marshal.ReleaseComObject(Wb)
  Wb = Nothing
  End Try
End Sub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
  Dim app As Excel.Application = New Excel.Application()
  AddHandler app.WorkbookOpen, AddressOf Workbook_Open
  ' 処理
End Sub
```
**C#**

```
public void app_WorkbookOpen(Excel.Workbook Wb){
  try
  {
    // 処理
  }
  finally
  {
    //オブジェクトを破棄します。
    Marshal.ReleaseComObject(Wb);
    Wb = null;
  }
}

private void button2_Click(object sender, EventArgs e){
  Excel.Application app = new Excel.Application();
  app.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(app_WorkbookOpen);
  // 処理
}
```
以上となります。

一般的には、上記の内容をお約束事と覚えていただいて実装を検討することで問題ないと思われます。

なお、ガベージ コレクトの強制など、一般的な開発では推奨されないとされる実装に踏み切るのは極めて大きな抵抗があるという開発者様のために、次回の投稿 ([Office オートメーションで割り当てたオブジェクトを解放する - Part2](https://officesupportjp.github.io/blog/Office%20%E3%82%AA%E3%83%BC%E3%83%88%E3%83%A1%E3%83%BC%E3%82%B7%E3%83%A7%E3%83%B3%E3%81%A7%E5%89%B2%E3%82%8A%E5%BD%93%E3%81%A6%E3%81%9F%E3%82%AA%E3%83%96%E3%82%B8%E3%82%A7%E3%82%AF%E3%83%88%E3%82%92%E8%A7%A3%E6%94%BE%E3%81%99%E3%82%8B%20-%20Part2/)) ではオブジェクト解放漏れが即時解放されない場合に具体的にどのように影響するかについて例を挙げて記載し、いわゆる「特別な事情がある場合」に該当することをご説明します。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります**。