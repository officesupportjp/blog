---
title: Office 開発におけるパフォーマンス トラブルシュート (その 2：ボトルネックの特定)
date: 2019-03-01
---

(※ 2018 年 7 月 12 日に Japan Office Developer Support Blog に公開した情報のアーカイブです。)

こんにちは、Office 開発サポート チームの中村です。

[前回の投稿](https://officesupportjp.github.io/blog/Office%20%E9%96%8B%E7%99%BA%E3%81%AB%E3%81%8A%E3%81%91%E3%82%8B%E3%83%91%E3%83%95%E3%82%A9%E3%83%BC%E3%83%9E%E3%83%B3%E3%82%B9%20%E3%83%88%E3%83%A9%E3%83%96%E3%83%AB%E3%82%B7%E3%83%A5%E3%83%BC%E3%83%88%20(%E3%81%9D%E3%81%AE%201%EF%BC%9A%E6%A6%82%E8%A6%81%E3%81%A8%E5%AF%BE%E5%87%A6%E6%96%B9%E6%B3%95)/)で、Office 開発のパフォーマンスに関する調査の進め方をご紹介しました。その中で予告した通り、今回の記事ではボトルネックとなるコードの特定手法について、特定作業に使えるサンプル コードとともに詳しく解説していきます。

今回も、「Office バージョンアップに伴い、これまで利用していたプログラムのパフォーマンスが低下した」という場合を例に説明します。コードの掲載などのため記事が長くなりますがご容赦ください。

**目次**  
[1\. 説明用サンプル プログラムの紹介](#1-説明用サンプル-プログラムの紹介)  
[2\. デバッグ ログの追加 (関数単位)](#2-デバッグ-ログの追加-関数単位)  
[3\. デバッグ ログの解析](#3-デバッグ-ログの解析)  
[4\. デバッグ ログの追加 (行単位)](#4-デバッグ-ログの追加-行単位)  
[5\. 関数単位の時刻ログ出力処理追加プログラム](#5-関数単位の時刻ログ出力処理追加プログラム)

#### **1\. 説明用サンプル プログラムの紹介**

かなりシンプルな例ですが、以下の VBA マクロ (TestProgram.xlsm) を例に説明します。(この VBA マクロの内容で大きくパフォーマンスが低下するわけではありません。あくまでも調査の流れを説明するためのサンプルです。そして説明用にあえて効率の悪い処理などを書いています。)

前回の記事でも述べた通り、業務で使用するプログラムは膨大なステップ数となり、多くの場合は処理の内容ごとに関数を分け、これを呼び出すよう実装されています。今回のサンプルでは、OpenWorkBook / SetValue / ChangeFontColor / CloseWorkBook の 4 つの関数を呼び出しています。

![](image1.png)

図 1. テスト プログラム実行結果

  

**\<テスト プログラム 処理概要\>**

ボタン コントロールをクリックすると、データ ブック (sample.xlsx) を開き、A1\~A3 セルの値を、このマクロ ブックの A4\~A6 に転記します。その後、A4\~A6 セルの書式をいくつか設定した後、ボタン コントロールの表示を「転記済み」に変更し、データブックを閉じて処理終了のメッセージボックスを表示します。

**テスト プログラム コード**

<Sheet1 オブジェクト>

```
Private Sub CommandButton1_Click()
    Dim dataFilePath As String
    Dim dataBook As Workbook
    Dim i As Integer
    
    '① マクロ ブックのシート上のデータを取得
    dataFilePath = ThisWorkbook.Worksheets(1).Range("B1").Value
    
    '② データブックをオープン
    Set dataBook = OpenWorkBook(dataFilePath)
    If dataBook Is Nothing Then
        MsgBox "データ ブックが存在しないため処理中断"
        Exit Sub
    End If
    
    '③ データブックの A1 ～ A3 の値をマクロブックの A4 ～ A6 に転記
    For i = 1 To 3
        Call SetValue(dataBook, i)
    Next
    
    Call ChangeFormat '④ 書式設定を変更
    Call CloseBook(dataBook) '⑤ データ ブックを閉じる
    
    '⑥ 処理完了時にコントロールの Caption を変更
    CommandButton1.Caption = "転記済み"
    
    MsgBox "処理終了"
End Sub

Function OpenWorkBook(filePath As String) As Workbook
    Dim wb As Workbook
    If Dir(filePath) = "" Then
        Exit Function
    End If
    
    Set wb = Workbooks.Open(filePath)
    Set OpenWorkBook = wb
End Function

Sub SetValue(dataBook As Workbook, i As Integer)
    ThisWorkbook.Worksheets(1).Range("A" & 3 + i).Value = dataBook.Worksheets(1).Range("A" & i).Value
End Sub

Sub ChangeFormat()
    With ThisWorkbook.Worksheets(1).Range("A4:A6")
        .Font.ColorIndex = 3
        .Font.Bold = True
        .Borders.LineStyle = xlDouble
        .Borders.ColorIndex = 5
    End With
End Sub

Sub CloseBook(dataBook As Workbook)
    dataBook.Close
End Sub
```
   

#### **2\. デバッグ ログの追加 (関数単位)**

通常、パフォーマンスが低下したと気づいたときの出発点では、「CommandButton1\_Click() 全体の処理が遅くなった」という状況から調査を開始するかと思います。この段階では、①～⑥のどこが遅くなったのか、まだ分かりません。またさらに、例えば ChangeFormat() の中であれば、いくつかの書式を変更しているので、どの書式変更が遅いのかが分かりません。

前回の記事で述べたように、ボトルネックとなる処理の内容によって、検討される対応方法は異なります。このため、まずはこのコードの中のどこが遅いのか、1 行レベルでの特定を行うことを目指します。

いきなり 1 行レベルの特定を行うのは効率が悪いので、まずは、大まかに関数単位で切り分けを行うことをお勧めします。

ここで役立つのが、「[5\. 関数単位の時刻ログ出力処理追加プログラム](#5-関数単位の時刻ログ出力処理追加プログラム)」のコードです。5. で後述する使い方に従ってこのプログラムを実行すると、先述のテスト プログラムは以下のようなコードに変更されます。

<Sheet1 オブジェクト>

```
Private Sub CommandButton1_Click()
LogWriteToBuffer "IN,Sheet1,CommandButton1_Click"
    Dim dataFilePath As String
    Dim dataBook As Workbook
    Dim i As Integer
    
    '① マクロ ブックのシート上のデータを取得
    dataFilePath = ThisWorkbook.Worksheets(1).Range("B1").Value
    
    '② データブックをオープン
    Set dataBook = OpenWorkBook(dataFilePath)
    If dataBook Is Nothing Then
        MsgBox "データ ブックが存在しないため処理中断"
LogWriteToBuffer "OUT,Sheet1,CommandButton1_Click"
        Exit Sub
    End If
    
    '③ データブックの A1 ～ A3 の値をマクロブックの A4 ～ A6 に転記
    For i = 1 To 3
        Call SetValue(dataBook, i)
    Next
    
    Call ChangeFormat '④ 書式設定を変更
    Call CloseBook(dataBook) '⑤ データ ブックを閉じる
    
    '⑥ 処理完了時にコントロールの Caption を変更
    CommandButton1.Caption = "転記済み"
    
    MsgBox "処理終了"
LogWriteToBuffer "OUT,Sheet1,CommandButton1_Click"

'★以下は手動で追加
LogWrite logOutputCollection
End Sub

Function OpenWorkBook(filePath As String) As Workbook
LogWriteToBuffer "IN,Sheet1,OpenWorkBook"
    Dim wb As Workbook
    If Dir(filePath) = "" Then
LogWriteToBuffer "OUT,Sheet1,OpenWorkBook"
        Exit Function
    End If
    
    Set wb = Workbooks.Open(filePath)
    Set OpenWorkBook = wb
LogWriteToBuffer "OUT,Sheet1,OpenWorkBook"
End Function

Sub SetValue(dataBook As Workbook, i As Integer)
LogWriteToBuffer "IN,Sheet1,SetValue"
    ThisWorkbook.Worksheets(1).Range("A" & 3 + i).Value = dataBook.Worksheets(1).Range("A" & i).Value
LogWriteToBuffer "OUT,Sheet1,SetValue"
End Sub

Sub ChangeFormat()
LogWriteToBuffer "IN,Sheet1,ChangeFormat"
    With ThisWorkbook.Worksheets(1).Range("A4:A6")
        .Font.ColorIndex = 3
        .Font.Bold = True
        .Borders.LineStyle = xlDouble
        .Borders.ColorIndex = 5
    End With
LogWriteToBuffer "OUT,Sheet1,ChangeFormat"
End Sub

Sub CloseBook(dataBook As Workbook)
LogWriteToBuffer "IN,Sheet1,CloseBook"
    dataBook.Close
LogWriteToBuffer "OUT,Sheet1,CloseBook"
End Sub
<Module1 オブジェクト> ※ 共通関数記述用に追加されます

Public logOutputCollection As Collection
Public Sub LogWriteToBuffer(strMsg As String)
   If logOutputCollection Is Nothing Then
      Set logOutputCollection = New Collection
   End If
   logOutputCollection.Add getNowWithMS & "," & strMsg
End Sub

Public Sub LogWrite(logOutputCollection As Collection)
    Dim j As Integer
    Dim iFileNo As Integer
    iFileNo = FreeFile
    Open "C:\temp\VBAPerf.log" For Append As #iFileNo
    If Not logOutputCollection Is Nothing Then
        For j = 1 To logOutputCollection.Count
           Print #iFileNo, logOutputCollection(j)
        Next
    End If
    Close #iFileNo
End Sub

Function getNowWithMS() As String
   Dim dtmNowTime      ' 現在時刻
   Dim lngHour         ' 時
   Dim lngMinute       ' 分
   Dim lngSecond       ' 秒
   Dim lngMilliSecond  ' ミリ秒
   dtmNowTime = Timer
   lngMilliSecond = dtmNowTime - Fix(dtmNowTime)
   lngMilliSecond = Right("000" & Fix(lngMilliSecond * 1000), 3)
   dtmNowTime = Fix(dtmNowTime)
   lngSecond = Right("0" & dtmNowTime Mod 60, 2)
   dtmNowTime = dtmNowTime \ 60
   lngMinute = Right("0" & dtmNowTime Mod 60, 2)
   dtmNowTime = dtmNowTime \ 60
   lngHour = Right("0" & dtmNowTime, 2)
   getNowWithMS = lngHour & ":" & lngMinute & ":" & lngSecond & "." & lngMilliSecond
End Function
```

  

Sub や Function に入った直後と出る直前に、「LogWriteToBuffer "IN またはOUT,<モジュール名>,<関数名>"」という処理が自動的に追加されているのがお分かりでしょうか。また、標準モジュールにモジュールが追加され、ログ出力関数 (LogWriteToBuffer() / LogWrite()) と、時刻をミリ秒まで取得する関数 (getNowWithMS()) が追加されます。

   

#### **3\. デバッグ ログの解析**

この状態でプログラムを実行すると、「[5\. 関数単位の時刻ログ出力処理追加プログラム](#5-関数単位の時刻ログ出力処理追加プログラム)」の中で指定しているログ ファイル (上記では C:\\temp\\VBAPerf.log) に以下のようにログが出力されます。

<ログ出力結果 サンプル> (数値は説明用に作成しているので、実際はこんなにかかりません。)

```
19:08:08.562,IN,Sheet1,CommandButton1_Click
19:08:08.562,IN,Sheet1,OpenWorkBook
19:08:09.210,OUT,Sheet1,OpenWorkBook
19:08:09.210,IN,Sheet1,SetValue
19:08:10.230,OUT,Sheet1,SetValue
19:08:10.230,IN,Sheet1,SetValue
19:08:11.732,OUT,Sheet1,SetValue
19:08:12.733,IN,Sheet1,SetValue
19:08:14.221,OUT,Sheet1,SetValue
19:08:14.221,IN,Sheet1,ChangeFormat
19:08:23.123,OUT,Sheet1,ChangeFormat
19:08:23.125,IN,Sheet1,CloseBook
19:08:23.532,OUT,Sheet1,CloseBook
19:08:23.541,OUT,Sheet1,CommandButton1_Click
```
  

このログを分析してどこで時間がかかっているかを特定していきます。Excel にカンマを区切り文字として貼り付けると、Excel の機能やシート関数を使って様々な観点で分析できます。

(例)

*   直前の処理との時間差 (「=A3-A2」ような数式で求められます) を算出し、特に時間がかかっている箇所を見つける
*   関数名のカウントを数え、繰り返し行われている処理を特定する

今回の場合、上記のログを解析すると、以下のあたりがボトルネックとなっていることが分かります。

1.  ChangeFormat 関数の IN ~ OUT の間が 1 回で 8.902 秒かかっている
2.  SetValue 関数の IN ～ OUT の間が 3 回の合計で約 4 秒かかっている

   

#### **4\. デバッグ ログの追加 (行単位)**

SetValue 関数の中は 1 行しか処理がないのでこれ以上の切り分けは必要ありませんが、ChangeFormat 関数の中ではいくつかの処理を行っているので、さらにボトルネックとなる処理を特定します。ここは手動で 1 行ごとにログを追加します。関数内のコード量に応じて、いきなり 1 行単位にログを追加するのではなく、何段階かに分けて絞り込んでも良いかと思います。

**ログ追加の例**

```
Sub ChangeFormat()
LogWriteToBuffer "IN,Sheet1,ChangeFormat"
    With ThisWorkbook.Worksheets(1).Range("A4:A6")
        .Font.ColorIndex = 3
LogWriteToBuffer "①,Sheet1,FontColorIndex"
        .Font.Bold = True
LogWriteToBuffer "②,Sheet1,FontBold"
        .Borders.LineStyle = xlDouble
LogWriteToBuffer "③,Sheet1,BorderLineStyle"
        .Borders.ColorIndex = 5
LogWriteToBuffer "④,Sheet1,BorderColorIndex"
    End With
LogWriteToBuffer "OUT,Sheet1,ChangeFormat"
End Sub
```

  

このようにログを追加したプログラムを再度実行し、3. と同じようにログを解析していきます。

注 : 本記事のログ出力処理追加サンプルプログラムでは、1 回の Excel 起動で複数回現象再現処理が実行されることは想定していません。(出力ログのコレクションが前回実行分と重複します) 実行の都度、Excel プログラムを開き直してください。

**対応方法の検討について**

これについては前回の記事で詳しく説明しているため今回はあまり触れませんが、例えば Font.ColorIndex の設定箇所が遅かったとします。この場合、前回の記事でいう[3-1. ボトルネックとなる処理の速度改善](https://officesupportjp.github.io/blog/Office%20%E9%96%8B%E7%99%BA%E3%81%AB%E3%81%8A%E3%81%91%E3%82%8B%E3%83%91%E3%83%95%E3%82%A9%E3%83%BC%E3%83%9E%E3%83%B3%E3%82%B9%20%E3%83%88%E3%83%A9%E3%83%96%E3%83%AB%E3%82%B7%E3%83%A5%E3%83%BC%E3%83%88%20(%E3%81%9D%E3%81%AE%201%EF%BC%9A%E6%A6%82%E8%A6%81%E3%81%A8%E5%AF%BE%E5%87%A6%E6%96%B9%E6%B3%95)/#3-1-%E3%83%9C%E3%83%88%E3%83%AB%E3%83%8D%E3%83%83%E3%82%AF%E3%81%A8%E3%81%AA%E3%82%8B%E5%87%A6%E7%90%86%E3%81%AE%E9%80%9F%E5%BA%A6%E6%94%B9%E5%96%84)として、Font.ColorIndex の代わりに Font.Color だったらどうか？などを試すことが検討できます。また、システム全体の流れによっては、マクロ ブックのテンプレートの段階でフォント色を設定しておき、マクロ内では変更しないといった対応も検討できるかもしれません。

SetValue については、ループ処理ではなく複数セルをまとめて貼り付けることが効果的です (前回記事 [3-2. プログラム構成の見直し](https://officesupportjp.github.io/blog/Office%20%E9%96%8B%E7%99%BA%E3%81%AB%E3%81%8A%E3%81%91%E3%82%8B%E3%83%91%E3%83%95%E3%82%A9%E3%83%BC%E3%83%9E%E3%83%B3%E3%82%B9%20%E3%83%88%E3%83%A9%E3%83%96%E3%83%AB%E3%82%B7%E3%83%A5%E3%83%BC%E3%83%88%20(%E3%81%9D%E3%81%AE%201%EF%BC%9A%E6%A6%82%E8%A6%81%E3%81%A8%E5%AF%BE%E5%87%A6%E6%96%B9%E6%B3%95)/#3-2-%E3%83%97%E3%83%AD%E3%82%B0%E3%83%A9%E3%83%A0%E6%A7%8B%E6%88%90%E3%81%AE%E8%A6%8B%E7%9B%B4%E3%81%97) に当たります)。

また、もし例えば⑤のコントロールの Caption 変更が遅いようであれば、フォーム コントロールに変更したらどうか (前回記事 3-1. ボトルネックとなる処理の速度改善)、Caption 変更以外の方法で処理完了が分かるようにできないか (前回記事 3-2. プログラム構成の見直し)、などの対応方法が検討できます。

   

#### **5\. 関数単位の時刻ログ出力処理追加プログラム**

2\. プログラム上のボトルネックの特定 で触れた、関数単位にログ出力処理を自動追加できるサンプル コードを以下に記載します。VBA の標準モジュールなどにコピーして利用できます。

なお、このサンプル コードはあらゆる VBA 上の記述方法に対応することを保証するものではありません。必要に応じて、お客様ご自身で修正してご利用ください。

**利用上の諸注意**

*   <span style="color:#ff0000">このサンプル コードの実行時</span>、ログ出力処理を追加する対象のファイルを開く Office アプリケーションのオプション設定で、\[セキュリティ センター\] – \[セキュリティ センターの設定\] – \[マクロの設定\] – <span style="color:#ff0000">\[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する\] を有効にする必要があります</span>。(パフォーマンス計測時には有効にする必要はありません)
*   VBAProject を参照・変更するため、VBAProject のパスワードは解除してください。
*   ご利用の際は、AddToLogFunc() 内初めの 2 行の Const 定義をご自身の環境に合わせて変更してください。
*   VBA 関数がどの順序で呼び出されるかは判断できないため、処理全体の最後に行うログのファイル出力処理呼び出しが含まれていません。<span style="color:#ff0000">手動で、処理全体の最後に「LogWrite logOutputCollection」 を追加してください</span>。(2. のサンプルでは CommandButton1\_Click() の最後に「'★以下は手動で追加」とコメントを入れている箇所です。)

**利用手順**

事前準備 : 上記の通り、Excel のオプション変更と VBAProject のパスワード解除を行っておきます。

1.  新規 Excel ブックの標準モジュールに以下のサンプル コードを貼り付けます。
2.  AddToLogFunc() 内の Const 定義をご自身の環境に合わせて変更します。
3.  AddToLogFunc() を実行します。ログ追加対象ファイルが開かれ、ログの追加が行われます。
4.  "ログ出力処理追加完了 : 該当ファイルを別名で保存し、VBA の内容に問題がないか確認してください" と表示されたらVBE でログ追加対象ファイルのコードを開き、現象再現手順の処理全体の最後に「LogWrite logOutputCollection」を追記します。
5.  ログ追加対象ファイルに別名を付けて保存します。
6.  Excel をいったん閉じ、ログを追加したファイルを再度開いて再現手順を実行します。
7.  出力されたログファイル (手順 2. で設定したファイル パスに出力) を開き、ログを解析します。

**サンプル コード**

```
'2018/7/12 公開版
'注 : このログ追加コードは、":" を用いて複数行を 1 行にまとめたコードには対応していません。

'Exit Sub / Exit Function へのログ出力処理追加
Function CheckExit(CurrentVBComponent As Object, strProcName As String, InProcCurrentLine As Long)

    Dim strCurrentCode As String
    Dim FoundPos As Long

    strCurrentCode = Trim(CurrentVBComponent.CodeModule.Lines(InProcCurrentLine, 1))

    'コメント行は除外
    If Left(strCurrentCode, 1) = "'" Or UCase(Left(strCurrentCode, 4)) = "REM " Then
        Exit Function
    End If

    'If xxx Then Exit Sub のような行の先頭以外に Exit がある書き方を考慮し InStr で合致確認
    FoundPos = InStr(strCurrentCode, "Exit ")
    
    '"Exit " に合致しても以下のケースは除外
    'a) Exit Sub の位置が先頭以外で、前にスペースがなく "xxxExit " のように文字列がある場合
    'b) 行の途中からコメントで、コメント部分に "Exit " がある場合
    'c) "Exit Sub" と "Exit Function" 以外の "Exit aaa" のような場合
    If FoundPos > 0 Then
        If (FoundPos > 1 And InStr(strCurrentCode, " Exit ") = 0) _
            Or (InStr(strCurrentCode, "'") > 0 And InStr(strCurrentCode, "'") < FoundPos) _
            Or (InStr(strCurrentCode, "Exit Sub") = 0 And InStr(strCurrentCode, "Exit Function") = 0) Then
            FoundPos = 0 'Exit Sub / Exit Function は見つからなかったものとみなす
        End If
    End If

    'Exit Sub / Exit Function のいずれかがある場合、その前に関数を抜けるログ出力処理追加
    If FoundPos > 0 Then
    
        'If xxx Then Exit Sub の書き方を考慮しコード整形
        If Left(strCurrentCode, 3) = "If " Then
        
            'いったん End If を入れる (If xxx Then Exit Sub + 改行 + End If)
            strCurrentCode = strCurrentCode & vbCrLf & "End If"
            
            '以下のように If 文を分割してコードを置き換え
            'Exit Sub の手前まで (If xxx Then )
            'ログ出力 (LogWriteToBuffer "OUT : " & CurrentVBComponent.Name & " : " & strProcName & "")
            'Exit Sub 以降 (Exit Sub + 改行 + End If)
            CurrentVBComponent.CodeModule.ReplaceLine InProcCurrentLine, _
                        Left(strCurrentCode, FoundPos - 1) & vbCrLf & _
                        "LogWriteToBuffer ""OUT," & CurrentVBComponent.Name & "," & strProcName & """" & vbCrLf & _
                        Mid(strCurrentCode, FoundPos)
        Else
            'その他通常の Exit の場合は直前に関数を抜けるログ追加
            CurrentVBComponent.CodeModule.InsertLines InProcCurrentLine, "LogWriteToBuffer ""OUT," & CurrentVBComponent.Name & "," & strProcName & """"
        End If
    End If
    
End Function


'メイン モジュール
Sub AddLogToFunc()
    
    '****************************************************
    'お客様環境に合わせて書き換えてください
    Const szBookFile = "C:\work\testProgram.xlsm" 'ログ出力処理を追加するファイルのフルパス
    Const szLogFile = "C:\work\VBAPerf.log" 'ログ出力ファイルのフルパス
    '****************************************************
    
    Const vbext_ct_StdModule = 1 'VBComponent Type 定数 : 標準モジュール
    Const vbext_pk_Proc = 0 'prockind 定数 : プロパティ プロシージャ以外のすべてのプロシージャ
    
    Dim xlBook As Workbook
    Dim CurrentVBComponent As Object 'VBComponent
    Dim TotalLine As Long
    Dim CurrentLine As Long
    Dim strProcName As String
    Dim strCurrentCode As String
    Dim ProcStartLine As Long
    Dim ProcEndLine As Long
    Dim InProcCurrentLine As Long
    Dim FoundPos As Long
    Dim strFunc As String
    
    Application.EnableEvents = False
    
    Set xlBook = Workbooks.Open(szBookFile) 'ログ出力処理追加対象ブック オープン

    ' 対象ブックに含まれる各モジュール内関数の最初と最後にログ出力関数呼び出しを追加
    For Each CurrentVBComponent In xlBook.VBProject.VBComponents
    
        TotalLine = CurrentVBComponent.CodeModule.CountOfLines
        
        'コード末尾から 1 行ごとに確認
        For CurrentLine = TotalLine To 1 Step -1
        
            strProcName = CurrentVBComponent.CodeModule.ProcOfLine(CurrentLine, vbext_pk_Proc) 'その行が属するプロシージャ名を取得
            strCurrentCode = LTrim(CurrentVBComponent.CodeModule.Lines(CurrentLine, 1))
            
            'End で始まる場合 : そのプロシージャの初めと終わりに処理追加
            If strProcName <> Empty And Left(strCurrentCode, 4) = "End " Then
            
                ProcStartLine = CurrentVBComponent.CodeModule.ProcBodyLine(strProcName, vbext_pk_Proc) + 1 'プロシージャの先頭行を取得
                ProcEndLine = CurrentLine
                
                'End 行の前に関数を抜けるログ出力処理追加
                CurrentVBComponent.CodeModule.InsertLines ProcEndLine, "LogWriteToBuffer ""OUT," & CurrentVBComponent.Name & "," & strProcName & """"
                'プロシージャ開始行の直後に関数に入るログ出力処理追加
                CurrentVBComponent.CodeModule.InsertLines ProcStartLine, "LogWriteToBuffer ""IN," & CurrentVBComponent.Name & "," & strProcName & """"
                
                'さらにこのプロシージャ内の途中で処理追加すべき箇所 (Exit Sub / Exit Function) をチェック
                '(OUT ログ追加により CurrentLine は End Sub の 1 行前を指す)
                For InProcCurrentLine = CurrentLine To ProcStartLine Step -1
                    CheckExit CurrentVBComponent, strProcName, InProcCurrentLine
                Next
                
                'このプロシージャの処理は終わっているためプロシージャ先頭までスキップ
                CurrentLine = ProcStartLine - 1
                
            End If
        Next
    Next

    '****************************************************
    'サブ関数追加処理
    '****************************************************
    
    ' ファイル出力時のディスク アクセスによるパフォーマンス影響を抑えるため、
    ' 処理実行中のログはいったん配列に書き込み、最後にファイル出力する

    'ログを配列に格納するための関数
    strFunc = "" & vbCrLf & _
              "Public logOutputCollection As Collection" & vbCrLf & _
              "Public Sub LogWriteToBuffer(strMsg As String)" & vbCrLf & _
              "   If logOutputCollection Is Nothing Then" & vbCrLf & _
              "      Set logOutputCollection = New Collection" & vbCrLf & _
              "   End If" & vbCrLf & _
              "   logOutputCollection.Add getNowWithMS & "","" & strMsg" & vbCrLf & _
              "End Sub"

    '配列からログファイルへの出力のための関数
    strFunc = strFunc & vbCrLf & _
              "" & vbCrLf & _
              "Public Sub LogWrite(logOutputCollection As Collection)" & vbCrLf & _
              "    Dim j As Integer" & vbCrLf & _
              "    Dim iFileNo As Integer" & vbCrLf & _
              "    iFileNo = FreeFile" & vbCrLf & _
              "    Open ""[LOG_FILE]"" For Append As #iFileNo" & vbCrLf & _
              "    If Not logOutputCollection Is Nothing Then" & vbCrLf & _
              "        For j = 1 To logOutputCollection.Count" & vbCrLf & _
              "           Print #iFileNo,  logOutputCollection(j)" & vbCrLf & _
              "        Next" & vbCrLf & _
              "    End If" & vbCrLf & _
              "    Close #iFileNo" & vbCrLf & _
              "End Sub"


    '現在時刻取得関数
    strFunc = strFunc & vbCrLf & _
              "" & vbCrLf & _
              "Function getNowWithMS() As String" & vbCrLf & _
              "   Dim dtmNowTime      ' 現在時刻" & vbCrLf & _
              "   Dim lngHour         ' 時" & vbCrLf & _
              "   Dim lngMinute       ' 分" & vbCrLf & _
              "   Dim lngSecond       ' 秒" & vbCrLf & _
              "   Dim lngMilliSecond  ' ミリ秒" & vbCrLf & _
              "   dtmNowTime = Timer" & vbCrLf & _
              "   lngMilliSecond = dtmNowTime - Fix(dtmNowTime)" & vbCrLf & _
              "   lngMilliSecond = Right(""000"" & Fix(lngMilliSecond * 1000), 3)" & vbCrLf & _
              "   dtmNowTime = Fix(dtmNowTime)" & vbCrLf & _
              "   lngSecond = Right(""0"" & dtmNowTime Mod 60, 2)" & vbCrLf & _
              "   dtmNowTime = dtmNowTime \ 60" & vbCrLf & _
              "   lngMinute = Right(""0"" & dtmNowTime Mod 60, 2)" & vbCrLf & _
              "   dtmNowTime = dtmNowTime \ 60" & vbCrLf & _
              "   lngHour = Right(""0"" & dtmNowTime, 2)" & vbCrLf & _
              "   getNowWithMS = lngHour & "":"" & lngMinute & "":"" & lngSecond & ""."" & lngMilliSecond" & vbCrLf & _
              "End Function"


    strFunc = Replace(strFunc, "[LOG_FILE]", szLogFile) '冒頭で設定したログファイル パスに書き換え
    xlBook.VBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.AddFromString strFunc '新しい標準モジュールを作成し関数を追記

    Application.EnableEvents = True
    
    MsgBox "ログ出力処理追加完了 : 該当ファイルを別名で保存し、VBA の内容に問題がないか確認してください"
    
End Sub
```

  

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**