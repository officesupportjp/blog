---
title: ラベル ポリシー「ユーザーがメールとドキュメントにラベルを適用することを必須にする」が有効の場合 Excel アドイン (.xlam) 保存時にラベル適用が求められ保存できない
date: 2024-12-25
id: cm53l15yg0000j46y7a4s6u9f
tags:
- Excel
- Office 開発全般
- アドイン
- MIP ラベル
---

こんにちは、Office サポート チームです。
今回の投稿では、Microsoft Information Protection ラベル ポリシー [ユーザーがメールとドキュメントにラベルを適用することを必須にする] が有効の場合、Excel アドイン (*.xlam) 編集後保存時にラベル適用を求められ保存ができない現象について説明します。


## 現象
Microsoft Information Protection が構成されており、ラベル ポリシーの [ユーザーがメールとドキュメントにラベルを適用することを必須にする] が有効の場合、ラベル付けされていないドキュメントを開くと、"必須ラベル付け" によってユーザー操作によるラベルの付与を強制するプロンプトが表示されます。
そのため、通常のユーザー利用シナリオにおいて、ラベルの付与が行われていないドキュメントを保存しようとした、または、ラベルが適用されていないドキュメントを操作した場合などにおいて、プロンプトが繰り返し表示されることが想定されます。

また、ラベル付けの処理が組み込まれていないマクロや RPA の処理が実行された場合、ラベルが適用されていないドキュメントを操作した場合も "必須ラベル付け" のプロンプトが表示される可能性があります。
これにより Microsoft Information Protection ラベル ポリシーに依存し、ファイル形式によっては Microsoft 365 Apps にて保存ができないケースが確認されています。

<br>

## 再現条件
以下の構成を全て満たす場合、"必須ラベル付け" のプロンプトが表示され Microsoft 365 Apps にてファイルを保存することができません。

- Microsoft Information Protection ラベル ポリシー [ユーザーがメールとドキュメントにラベルを適用することを必須にする] が有効に設定されている
- Microsoft Information Protectionラベル ポリシー [既存のドキュメントに既定のラベルを適用する] が設定されていない
- Microsoft 365 Apps で Excel アドイン形式 (*.xlam) のファイルを開く、かつ、Excel アドイン形式 (*.xlam) にラベルが未設定である

<br>

## 再現手順の例
1. Excel アプリを起動し、[ファイル]-[開く] から、ラベルが付与されていない Excel アドイン形式 (*.xlam) のファイルを開きます。
※ Excel アドイン形式 (*.xlam) のファイルは、ファイルを開いても、ワークシートが表示されないことにご留意ください。
2. Alt + F12 キーを押下し、Visual Basic Editor を起動します。
3. Visual Basic Editor で、任意の編集を加えます。例えば、ThisWorkbookに下記コードを追加します。

```
Private Sub Workbook_Open()
    Debug.Print "Test"
End Sub
```

4. Visual Basic Editor 画面で保存ボタンをクリックします。

<br>

#### 結果
ラベルの適用を求めるプロンプトが表示されます。
[ラベルの選択] ボタンをクリックしても、ラベル選択のダイアログが表示されないため、このユーザー インターフェースからラベルを設定することはできません。
このため、Visual Basic Editor 画面で行った変更内容は保存することができません。

![](image0.png)

<br>

## 回避策
以下いずれかの方法による対応をご検討いただきますようお願いいたします。

(a). Microsoft Information Protectionラベル ポリシー [既存のドキュメントに既定のラベルを適用する] で暗号化を含まないラベルを設定する

あるいは

(b). VBA マクロにて、暗号化を含まないラベルを設定する

<br>

#### (b). の設定手順例

本手順は一度実行していただければ結構です。設定したラベルを削除しない限り、毎回実行する必要はございません。

1. Excel アプリを起動し、ラベルが付与されていない Excel アドイン形式 (*.xlam) のファイルを開きます。
※ Excel アドイン形式 (*.xlam) のファイルは、ファイルを開いても、ワークシートは表示されないことにご留意ください。
2. Alt + F12 キーを押下し、Visual Basic Editor を起動します。
3. Visual Basic Editor で、以下のコードを追加したら、F5 キーを押下してマクロ実行します。
※ "<暗号化設定のないラベルの GUID>" は、お客様の環境で作成済みの暗号化設定のないラベルの GUID を指定してください。
※ 暗号化設定を含むラベルは設定しないでください。暗号化設定を含むラベルを設定すると、Excel アドイン形式 (*.xlam) として使用できません。

```
Sub Sample()
Dim mysensitivityLabel As Office.SensitivityLabel
Dim myLabelInfo As Office.LabelInfo

    Set mysensitivityLabel = ThisWorkbook.SensitivityLabel  
    Set myLabelInfo = ThisWorkbook.SensitivityLabel.CreateLabelInfo()

        With myLabelInfo
            .AssignmentMethod = MsoAssignmentMethod.PRIVILEGED
            .LabelId = "<暗号化設定のないラベルの GUID>"
        End With

    mysensitivityLabel.SetLabel myLabelInfo, myLabelInfo
End Sub

```
4. Excel アドイン形式 (*.xlam) に暗号化設定のないラベルが設定できたら、Visual Basic Editor 画面で保存ボタンをクリックして保存を行います。
5. 保存が完了したら、3. で追加したマクロ自体は削除して構いません。削除後は再度、Visual Basic Editor 画面で保存を行います。

<br>

## 今後の対応予定
本動作について、現在開発部門にて今後の改善可否等を含め調査中となります。
新たな情報が確認された場合は、この記事に追記する形で公開していく予定です。

<br>

## ご参考情報
ラベル ポリシーにおいて "必須ラベル付け" を有効にする場合、併せて "既定のラベル" によってプロンプトの表示が抑制されることを公開情報によりご案内しております。
"既定のラベル" は、"必須ラベル付け" よりも常に優先されるため、最低限、すべてのドキュメントに既定のラベルが付与されます。
その結果、"必須ラベル付け" によって表示されるプロンプトを抑制することが可能です。

Title: 秘密度ラベルの詳細
URL: [https://learn.microsoft.com/ja-jp/purview/sensitivity-labels?view=o365-worldwide#what-label-policies-can-do](https://learn.microsoft.com/ja-jp/purview/sensitivity-labels?view=o365-worldwide#what-label-policies-can-do)
> 項目 [ラベル ポリシーでできること] > [ドキュメントとメール、文書のみ、コンテナー、および Power BI コンテンツに ラベルを適用することをユーザーに要求します。] をご参照ください \
>> このオプションを使用して、ラベル付けの適用範囲を広げることを検討してください。 ただし、ユーザーのトレーニングを実施しない場合、この設定は不正確なラベル付けにつながる可能性があります。 さらに、対応する既定のラベルも設定しない限り、必須のラベル付けにより頻繁にダイアログが表示され、ユーザーを苛立たせる可能性があります。

Title: Office アプリで秘密度ラベルを管理する
URL: [https://learn.microsoft.com/ja-jp/purview/sensitivity-labels-office-apps?view=o365-worldwide#require-users-to-apply-a-label-to-their-email-and-documents](https://learn.microsoft.com/ja-jp/purview/sensitivity-labels-office-apps?view=o365-worldwide#require-users-to-apply-a-label-to-their-email-and-documents)
> 項目 [ユーザーがメールとドキュメントにラベルを適用することを必須にする] をご参照ください \
>> 既定のラベルは、必須のラベル付けよりも常に優先されます。



今回の投稿は以上です。
本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。



