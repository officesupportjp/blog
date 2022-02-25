---
title: Office のパフォーマンス トラブルシュート (その 3：はじめに試す設定 - 描画編 -)
date: 2019-10-29
---

こんにちは、Office サポート チームの中村です。 

以前の投稿で、Office 開発に特化して Office のパフォーマンスが低下する一般的な理由や、調査手法などをご説明しました。 

Office 開発におけるパフォーマンス トラブルシュート (その 1 : 概要と対処方法)  
[https://officesupportjp.github.io/blog/Office 開発におけるパフォーマンス トラブルシュート (その 1：概要と対処方法)/](https://officesupportjp.github.io/blog/Office%20%E9%96%8B%E7%99%BA%E3%81%AB%E3%81%8A%E3%81%91%E3%82%8B%E3%83%91%E3%83%95%E3%82%A9%E3%83%BC%E3%83%9E%E3%83%B3%E3%82%B9%20%E3%83%88%E3%83%A9%E3%83%96%E3%83%AB%E3%82%B7%E3%83%A5%E3%83%BC%E3%83%88%20(%E3%81%9D%E3%81%AE%201%EF%BC%9A%E6%A6%82%E8%A6%81%E3%81%A8%E5%AF%BE%E5%87%A6%E6%96%B9%E6%B3%95)/)

Office 開発におけるパフォーマンス トラブルシュート (その 2 : ボトルネックの特定)  
[https://officesupportjp.github.io/blog/Office 開発におけるパフォーマンス トラブルシュート (その 2：ボトルネックの特定)/](https://officesupportjp.github.io/blog/Office%20%E9%96%8B%E7%99%BA%E3%81%AB%E3%81%8A%E3%81%91%E3%82%8B%E3%83%91%E3%83%95%E3%82%A9%E3%83%BC%E3%83%9E%E3%83%B3%E3%82%B9%20%E3%83%88%E3%83%A9%E3%83%96%E3%83%AB%E3%82%B7%E3%83%A5%E3%83%BC%E3%83%88%20(%E3%81%9D%E3%81%AE%202%EF%BC%9A%E3%83%9C%E3%83%88%E3%83%AB%E3%83%8D%E3%83%83%E3%82%AF%E3%81%AE%E7%89%B9%E5%AE%9A)/)

過去の投稿で説明した通り、パフォーマンスの問題に直面したとき、そのボトルネックとなっている処理の内容によって対応は異なるため、本来のアプローチとしては、ボトルネックを特定してから対応を検討する必要があります。

ただ今回からの記事では、これまでの弊社サポートへのお問い合わせ実績や Office の仕組みなどから、効果があることが多い設定をいくつかご案内したいと思います。Office アドインの処理が遅いけれど自社開発ではないので詳細な調査はできない、プレミア サポートで時間をかけて調査をする前に何か試したい、といった状況での初期対応として、以下の設定を順に試してみていただければと思います。なお、これらの設定は Office 開発に限らず、通常の利用でのパフォーマンス低下の改善にも活用できます。 

多数の項目をご案内するため、記事を何回かに分けて投稿します。今回は描画に関わる設定をご案内します。

OS や Office のバージョンが新しくなるにつれ、よりリッチな表現、滑らかで精密な描画を行うため、描画処理のパフォーマンスコストは増加傾向にあります。OS や Office アプリケーション内の実装で工夫している部分や、マシンの性能向上でカバーされる部分もありますが、どうしても以前のバージョンと比較してはパフォーマンスがトレードオフとなることがあり、これによって、ファイルオープン、ファイル内容の変更、シートの切り替え、スクロール、セルのフォーカス移動などでパフォーマンス影響が生じます。リッチな表現が不要な場合は、描画に関わる動作を無効化・軽減する設定を行うとパフォーマンス改善につながります。

## **目次**  
0\. Office の設定変更手法  
1\. アニメーションの無効化  
2\. リアルタイム プレビューの無効化  
3\. ハードウェア グラフィック アクセラレータの設定変更  
4\. シートの表示倍率  
5\. 図のリンク (カメラ)  
6\. 数式バー表示  
7\. ステータスバー表示  
8\. (プログラム向け) 描画処理全般の停止  
9\. (プログラム向け) Excel の非表示化  
10\. (OS) アニメーションの無効化  
11\. (OS) 視覚効果  
12\. (OS) 省電力設定 ※ これは描画以外も含め、すべての動作に関わります
<br>

## **0\. Office の設定変更手法**  
設定項目を案内する前に、Office の各種設定を変更する手法について説明します。

Office では設定変更 GUI として、画面メニュー \[オプション\] や、グループポリシーで Office 管理用テンプレートが用意されています。これらの項目を設定すると、その設定内容はレジストリに保持されます。  
このため、ある項目の動作を制御する場合、まず以下の 3 種類が検討できます (項目により、3 種類すべてが用意されていないものもあります)。お客様の利用しやすい方法をご利用ください。

*   オプションから変更する (オプションに項目がある場合)
*   グループポリシーから変更する (グループポリシーの Office 管理用テンプレートに項目がある場合)
*   レジストリを変更する (オプションやグループポリシーに対応するレジストリを直接変更する場合、メニューはなくレジストリのみで制御可能な場合)

また、VBA や .NET などのプログラムからの操作の場合、プロパティを利用して設定を変更できる項目もあります。プロパティには、Office の一般的な動作変更の他に、マクロの処理パフォーマンスに影響を与えるものもあります。

本記事では、Office 2016 / Office 2019 / Office 365 の Excel / Word / PowerPoint を対象に、これらの設定項目を記載します。ご利用の Officeバージョンによっては、記載したメニューが用意されていない場合や、バージョン名に関わる部分を読み替える必要があります。読み替える場合、レジストリについては "16.0" の部分がバージョンによって以下のパスとなります。

Office 2010 → 14.0  
Office 2013 → 15.0

一部の設定は、Access や Visio、Outlook といった他の Office アプリケーションでも有効なもの、同等の設定があるものもありますが、本記事では言及しません。
<br>

## **1\. アニメーションの無効化 (Excel / Word / PowerPoint)**  
画面上での移動を滑らかにする動作を無効化します。設定 1 ～ 3 では影響する操作などが異なるため、それぞれお試しください。

<u>設定 1</u>  
<グループポリシー\>  
\[ユーザーの設定\] - \[Microsoft Office 2016\] - \[その他\] - \[Office アニメーションを無効にする\]  
設定値 : 有効  

<レジストリ\>  
キー: HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Graphics  
名前: DisableAnimations  
種類: REG\_DWORD  
値: 1  


<u>設定 2</u>
**Excel**  
<オプション\>  
\[簡単操作\] – \[操作をアニメーションで表示する\] チェックをオフ  

<レジストリ\>  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Excel\\Options  
名前 : EnableAnimations  
種類 : REG\_DWORD  
値 : 0  

**Word**  
<オプション\>  
\[簡単操作\] – \[操作の結果をアニメーションで表示する\] のチェックをオフ  

<レジストリ\>  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Word\\Options  
名前 : AnimateScreenMove  
種類 : REG\_DWORD  
値 : 0  

※ PowerPoint には本項目はありません。

<u>設定 3</u>  
**Excel**  
Application.EnableMacroAnimations プロパティを False に設定  
[https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.enablemacroanimations](https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.enablemacroanimations)  
\* 既定は False ですが、マクロ内や他のアドインなどで True に設定している場合は効果が期待できます。  

※ Word / PowerPoint には本項目はありません。  
<br>

## **2\. リアルタイム プレビューの無効化 (Excel / Word / PowerPoint)**  
例えば文字の色をカラーパレットから選択するときなど、確定前に文字に色が反映される動作を無効化します。

<オプション\>  
\[基本設定\] - \[ユーザー インターフェイスのオプション\] - \[リアルタイムのプレビュー表示機能を有効にする\] のチェックをオフ  

<グループポリシー\>  
\[ユーザーの設定\] - \[Microsoft <アプリケーション名\> 2016\] - \[<アプリケーション名\> のオプション\] - \[全般\] - \[リアルタイムのプレビュー表示機能を有効にする\]  
設定値 : 無効  
(<アプリケーション名\> には Excel / Word / PowerPoint のいずれかが入ります。以降の記述でも同様です。)

<レジストリ\>  
キー : HKEY\_CURRENT\_USER \\Software\\Microsoft\\Office\\16.0\\<アプリケーション名\>\\Options  
名前 : LivePreview  
種類 : REG\_DWORD  
値 : 0  

<プロパティ\>  
**Excel**  
Application.EnableLivePreview プロパティ をFalse に設定  
[https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.enablelivepreview](https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.enablelivepreview)  

**Word**  
Options.EnableLivePreview プロパティを False に設定  
[https://docs.microsoft.com/ja-jp/office/vba/api/word.options.enablelivepreview](https://docs.microsoft.com/ja-jp/office/vba/api/word.options.enablelivepreview)  

※ PowerPoint はプロパティはありません
<br>

## **3\. ハードウェア グラフィック アクセラレータの設定変更 (Excel / Word / PowerPoint)**  
描画処理にハードウェアの機能を利用すると、一般的には高速になります。ただ、ハードウェアとの相性によって、有効にするとかえって一部の処理が遅くなる場合もあります。ハードウェアグラフィック アクセラレータの利用有無を変更してみてください。

<オプション\>  
\[オプション\] - \[詳細設定\] - \[ハードウェアのグラフィックアクセラレータを無効にする\] のチェックを変更  
※ Office 共通設定のため、いずれかのアプリケーションで変更すると、Office 全体に反映されます。  

<グループポリシー\>  
\[Microsoft Office 2016\] – \[その他\] – \[ハードウェア グラフィック アクセラレータを使用しない\]  
設定値 : 有効または 無効  

<レジストリ\>  
キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Graphics  
名前 : DisableHardwareAcceleration  
種類 : REG\_DWORD  
値 : 1  
<br>

## **4\. シートの表示倍率 (Excel)**  
Excel でシートが 100% 以外のサイズで表示されていると、シート内容の表示位置やサイズを調整して描画します。パフォーマンスが気になるブックは、倍率 100% で保存してみてください。 

<手順\>  
Excel のステータス バーをスライド、またはリボンの \[表示\] – \[ズーム\] – \[100%\]  

<プロパティ\>  
Window.Zoom プロパティを 100 に設定  
[https://docs.microsoft.com/en-us/office/vba/api/excel.window.zoom](https://docs.microsoft.com/en-us/office/vba/api/excel.window.zoom)
<br>

## **5\. 図のリンク (カメラ) (Excel)**

図形や画像に、数式バーでセルを参照する数式を設定すると、参照先のセルをオブジェクト内に表示できます。この状態では、参照先セルの変更有無を判定する都度再描画され、パフォーマンスが低下します。リンクが不要な場合は、図のリンクを解除して静的な内容に変更します。  

<手順\>  
図形や画像を選択し、数式バーで数式を削除

<プロパティ\>  
対象オブジェクトの Formula プロパティの内容をクリア  
[https://docs.microsoft.com/en-us/office/vba/api/excel.range.formula](https://docs.microsoft.com/en-us/office/vba/api/excel.range.formula)
<br>

## **6\. 数式バー表示 (Excel)**  
Excel の数式バー部分は子ウィンドウになっており、数式バー自体の描画にある程度のコストがかかります。プログラムなどから繰り返し再描画を行う場合は、数式バーを非表示にすることをお試しください。  

<オプション\>  
\[詳細設定\] – \[表示\] – \[数式バーを表示する\] のチェックをオフ  
※ オプションからは標準表示の状態のみ制御できます。    

<グループポリシー\>  
\[Microsoft Excel 2016\] – \[Excel のオプション\] – \[詳細設定\]  
  \- \[標準表示で数式バーを表示する\] -> 通常のウィンドウ状態での数式バー表示を制御  
  \- \[全画面表示で数式バーを表示する\] -> 全画面表示 (Alt + V + U) での数式バー表示を制御  
設定値 : 無効    

<レジストリ\>  
**標準表示で数式バーを表示する**
キー : HKEY\_CURRENT\_USER\\Software\\Policies\\Microsoft\\Office\\16.0\\Excel\\Options\\Binaryoptions  
名前 : ffmlaf\_1\_1  
種類 : REG\_DWORD  
値 : 0  

**全画面表示で数式バーを表示する**
キー : HKEY\_CURRENT\_USER\\Software\\Policies\\Microsoft\\Office\\16.0\\Excel\\Options\\Binaryoptions  
名前 : ffmlafull\_68\_1  
種類 : REG\_DWORD  
値 : 0  
  
<プロパティ\>  
Application.DisplayFormulaBar プロパティ を False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayformulabar](https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayformulabar)
<br>

## **7\. ステータスバー表示 (Excel / Word / PowerPoint)**  
数式バーと同様、ステータスバー自体の描画、内容の更新もパフォーマンスに影響します。手動操作では、表示内容を減らすことをお試しください。Excel では、プログラムの場合、ステータスバー自体を非表示にできます。

<手順\>  
ステータス バーを右クリックし、\[ステータス バーのユーザー設定\] メニューで不要な項目のチェックを外します。  
参考) Excel のステータスバーのオプション  
[https://support.office.com/ja-jp/article/6055ecd9-e20f-4a7a-a611-4481bd488c55](https://support.office.com/ja-jp/article/6055ecd9-e20f-4a7a-a611-4481bd488c55)

<プログラム\>  
Application.DisplayStatusBar プロパティを False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/excel.application.displaystatusbar](https://docs.microsoft.com/en-us/office/vba/api/excel.application.displaystatusbar)
<br>

## **8\. (プログラム向け) 描画処理全般の停止**  
プログラムの処理中は、処理途中の描画状態をユーザーに見せる必要はないことが多いため、以下のプロパティを設定して画面描画を抑止することをお試しください。(多くの描画処理が抑止されますが、完全に描画が一切行われなくなるわけではありません。)

**Excel**  
Application.ScreenUpdating プロパティを False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/excel.application.screenupdating](https://docs.microsoft.com/en-us/office/vba/api/excel.application.screenupdating)

**Word**  
Application.ScreenUpdatingプロパティを False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/word.application.screenupdating](https://docs.microsoft.com/en-us/office/vba/api/word.application.screenupdating)

※ PowerPoint はプロパティはありません  
<br>

## **9\. (プログラム向け) Excel の非表示化**  
外部からの Office オートメーションなどでは、プログラムの処理中にアプリケーション ウィンドウ自体を非表示にすることも有効です。以下のプロパティで非表示にできます。

**Excel**  
Application.Visible プロパティを False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/excel.application.visible](https://docs.microsoft.com/en-us/office/vba/api/excel.application.visible)

**Word**  
Application.Visible プロパティを False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/word.application.visible](https://docs.microsoft.com/en-us/office/vba/api/word.application.visible)

**PowerPoint**  
Application.Visible プロパティを False に設定  
[https://docs.microsoft.com/en-us/office/vba/api/powerpoint.application.visible](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.application.visible)
<br>

## **10\. (OS) アニメーションの無効化**  
Office だけでなく他のアプリケーション全般に及ぶ設定ですが、OS レベルでアニメーションを無効化できます。

<オプション\>  
**Wndows 10**  
\[設定\] – \[簡単操作\] – \[ディスプレイ\] – \[Windows のシンプル化と個人用設定\] – \[Windows にアニメーションを表示する\]  
をオフに設定し、OS を再起動してパフォーマンスを確認してください。(バージョンによりメニュー構成が異なる場合があります。見つからない場合は \[設定\] 画面の検索から "アニメーション" で探してみてください。)
<br>

## **11\. (OS) 視覚効果**  
OS の設定で、視覚効果を細かく制御できます。パフォーマンスを優先する設定に変更してお試しください。

<オプション\>  
\[コントロール パネル\] – \[システムとセキュリティ\] – \[システム\] - \[システムの詳細設定\] – \[詳細設定\] タブ – \[パフォーマンス\] 内の \[設定\] – \[視覚効果\] タブで \[パフォーマンスを優先する\] を選択 (または \[カスタム\] で不要な効果のチェックをオフ)  
<br>

## **12\. (OS) 省電力設定 ※ これは描画以外も含め、すべての動作に関わります**  
大幅なパフォーマンス低下で多いのはこの設定です。ノートPC では特に、省電力設定となっていることがあり、省電力設定では CPU クロック数が制限されるため、CPU に負荷がかかる様々な処理が遅延します。Office の場合、この設定によって起動やスクロールが遅いというお問い合わせを頂きます。

<オプション\>  
\[コントロール パネル\] – \[ハードウェアとサウンド\] – \[電源オプション\] で現在の設定を確認できます。  
"省電力" に設定されている場合は、"バランス" に変更し、OS を再起動してパフォーマンスを確認してください。  
<br>

今回の投稿は以上です。
<br>

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**