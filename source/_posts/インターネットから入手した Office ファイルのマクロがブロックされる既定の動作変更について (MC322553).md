---
title: インターネットから入手した Office ファイルのマクロがブロックされる既定の動作変更について (MC322553)
date: '2022-02-18'
lastupdate: '2025-01-16'
tags:
  - VBA マクロ
  - セキュリティ
id: cl0m4f5o1003504vsa5tp25c2

---

<span style="color:#ff0000">**2022/03/18 Update**</span>  
<span style="color:#339966">- "【インターネット入手ファイルとみなされるシナリオ】" を追記しました。</span>  

<span style="color:#339966">- "【マクロを有効にする方法について】" 内に以下を追記しました。</span>  
<span style="color:#339966"> "1) ポリシーでマクロを有効にする方法" に [マクロの設定] による制御が実施されることを追記しました。</span>  
<span style="color:#339966"> "4) Web 上のファイルを開くシナリオで、信頼済みサイトに登録する方法" を追記しました。</span>  

<span style="color:#ff0000">**2022/06/08 Update**</span>  
<span style="color:#339966">ネットワーク共有がインターネット ゾーンとみなされるシナリオについて追記しました。</span>  

<span style="color:#ff0000">**2022/07/08 Update**</span>  
<span style="color:#339966">最新チャネルへのリリースがロールバックされたため、冒頭の記載を変更しました。</span>  

<span style="color:#ff0000">**2022/07/29 Update**</span>  
<span style="color:#339966">最新チャネルへのリリース再開されたため、記載を変更しました。</span>  

<span style="color:#ff0000">**2022/10/14 Update**</span>  
<span style="color:#339966">本動作変更に関する新しいブログ記事「インターネット入手マクロ ブロックのトラブルシュートの流れ」へのリンクを追加しました。</span>  

<span style="color:#ff0000">**2025/01/16 Update**</span>  
<span style="color:#339966">Project / Publisher も本動作に対応したため、対象アプリとして追記しました。</span>  


<br>

こんにちは、Office サポート チームです。

本記事では、2022 年 4 月 より Microsoft 365 Apps への展開が開始されるインターネットから入手した Office ファイルのマクロがブロックされる既定の動作変更について説明します。本記事に記載の既定の動作変更は 7 月 27 日 (米国時間) に最新チャネルにおいてリリースが開始されました。

<br>

# 【前提】

今回のご説明は以下公開情報でご案内している内容となります。

タイトル : Macros from the internet will be blocked by default in Office  
アドレス : [https://docs.microsoft.com/ja-jp/DeployOffice/security/internet-macros-blocked](https://docs.microsoft.com/ja-jp/DeployOffice/security/internet-macros-blocked) (日本語版 - 機械翻訳)  
アドレス : [https://docs.microsoft.com/en-us/DeployOffice/security/internet-macros-blocked](https://docs.microsoft.com/en-us/DeployOffice/security/internet-macros-blocked) (英語版)  

また、以下のとおり、ブログ記事も公開されています。

タイトル : Helping users stay safe: Blocking internet macros by default in Office  
アドレス : [https://techcommunity.microsoft.com/t5/microsoft-365-blog/helping-users-stay-safe-blocking-internet-macros-by-default-in/ba-p/3071805](https://techcommunity.microsoft.com/t5/microsoft-365-blog/helping-users-stay-safe-blocking-internet-macros-by-default-in/ba-p/3071805) (英語版)

上記公開情報 / ブログ記事には最新状況の更新がされている場合があるため、適宜ご確認ください。

**追加情報**  
この記事では変更の詳細をご説明していますが、ブロックされる理由は複数あります。実際にこの変更でマクロがブロックされる動作に直面した場合に、どの理由でブロックされているかを特定する流れについて以下のブログを作成しましたので、合わせてご参照ください。

タイトル : インターネット入手マクロ ブロックのトラブルシュートの流れ
アドレス : [https://officesupportjp.github.io/blog/cl97y6c2w0000ykse7nll9auc/](https://officesupportjp.github.io/blog/cl97y6c2w0000ykse7nll9auc/)

<br>

# 【概要】

上記公開情報で記載の概要は、Office のセキュリティ向上のため、 Office アプリケーションにてインターネットから入手した Office ファイルのマクロの実行をブロックするよう既定の動作を変更するという内容です。

これまでの Office では、インターネットから入手したファイルは保護ビューでは開かれますが、保護ビューを解除して編集モードにした後のマクロの実行可否の判断には、インターネットからの入手であることは考慮されず、他のセキュリティ設定でブロックされない限りはマクロを実行できました。今回の変更で、インターネットから入手したファイルは既定でマクロ実行がブロックされるようになります。

本変更はインターネットから入手したすべての Office ファイルが対象ではなく、特定の条件でマクロの有効 / 無効が決定されなかった場合にブロックされる動作となります。

つまり、この条件次第でインターネットから入手するファイルのマクロを今後も実行できるようになりますので、引き続きマクロを実行したい場合は、いずれかの先に判断される条件を利用してマクロを有効にできますので、本記事で詳しく紹介します。

なお、ここで言う「インターネットから入手したファイル」とは、Mark of the Web (MOTW) 属性と呼ばれる Windows によってインターネットからのファイル取得時に自動的に付与される Web マークが付与されたファイルを指します。


<br>

# 【インターネット入手ファイルとみなされるシナリオ】
以下のようなシナリオがインターネット入手ファイルとみなされ、既定でマクロが無効化されるようになります。
- 外部から受信したメールの添付ファイルを開く
- ブラウザーでインターネット ゾーンのサイトからダウンロードしたファイルを開く
- インターネット ゾーンの Web サイト上のファイルを直接開く (IP アドレスや FQDN で接続されたネットワーク共有のファイルを開く場合を含む)

<br>

# 【本変更の対象範囲について】
先の 公開情報 "この変更のOfficeのバージョン" に記載のとおり、本件の対象となる Office アプリケーションは Access / Excel / PowerPoint / Word / Publisher / Visio / Project です。

また、本変更の対象の Office は Windows デバイスにインストールされる以下となります。
- Microsoft 365 Apps 最新チャネル (プレビュー)
- Microsoft 365 Apps 最新チャネル
- Microsoft 365 Apps 月次エンタープライズ チャネル
- Microsoft 365 Apps 半期エンタープライズ チャネル (プレビュー)
- Microsoft 365 Apps 半期エンタープライズ チャネル
- 製品版 Office 2021
- 製品版 Office 2019
- 製品版 Office 2016

<br>

# 【変更される既定の動作の詳細】

今回の既定の動作の変更は、先の公開情報に記載の以下フローチャートにて <span style="color:#ff0000">**赤枠で囲った 7 の状況**</span> となった場合に該当します。

![](vba-macro-flowchart.png)

上記フローチャートにおいて今回の動作変更の影響を受ける条件は以下をすべて満たす場合となります。

- 開かれるマクロ付きファイルにインターネットから入手されたことを意味する Web マークが付与されている (フローチャート 1 = START)
- マクロ付きファイルが "信頼できる場所"  に保存されていない (フローチャート 2 = No)
- マクロ付きファイルが デジタル署名されていない、またはデジタル署名されているが "信頼できる発行元" に登録されている証明書と署名が一致しない (フローチャート 3 = No)
- クラウド ポリシー / グループ ポリシー等のポリシーにてマクロの有効 / 無効が設定されていない (フローチャート 4a,4b = No)
- ユーザーが以前このファイルを開いていない、または以前このファイルを開いたが表示される通知にて [コンテンツを有効にする] を選択していない (フローチャート 6 = No)

これらの条件すべてに合致した場合、以下のとおり動作が変更となります。

### 変更前
以下のオプション設定に従ってマクロの有効 / 無効を決定します。

#### <オプション設定 (Excel アプリケーションから確認する例)>
[ファイル] タブ - [オプション] - [トラスト センター]- [トラスト センターの設定] - [マクロの設定] - [マクロの設定] の選択

![](macro-settings.png)

なお、既定の設定である "警告して、VBA マクロを有効にする" の場合、以下が出力されます。

![](vba-security-warning-banner.png)


### 変更後
以下のメッセージが出力されマクロの実行がブロックされます。

![](vba-security-risk-banner.png)


<br>

# 【マクロを有効にする方法について】

上記の動作変更にて無効化されたマクロを有効にする方法を以下に説明します。

## <p id="EnableVBA-1">1) ポリシーでマクロを有効にする方法</p>
Office 管理用テンプレートのグループ ポリシーなどを利用し以下を設定するとインターネットから入手した Office ファイルについてもマクロを有効にできます。

#### <グループ ポリシー設定>
- Access  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft Access 2016] - [アプリケーション設定] - [セキュリティ] - [セキュリティ センター]  
\- [インターネットから取得した Office ファイル内のマクロの実行をブロックします] : 無効
- Excel / PowerPoint / Word / Visio / Project  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft xxx 2016] - [xxx のオプション] - [セキュリティ] - [セキュリティ センター]  
\- [インターネットから取得した Office ファイル内のマクロの実行をブロックします] : 無効  
注) xxx はそれぞれ Excel / PowerPoint / Word / Visio / Project です。
- Publisher  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft Publisher 2016] - [セキュリティ] - [セキュリティ センター]  
\- [インターネットから取得した Office ファイル内のマクロの実行をブロックします] : 無効
<br>

なお、この設定とは別にオプションの [トラスト センターの設定] - [マクロの設定] による、マクロ全体のセキュリティ設定による制御は行われます。例として、既定の “警告して、VBA マクロを有効にする” 設定の場合には、以下が表示されます。

![](vba-security-warning-banner.png)

<br>

## <p id="EnableVBA-2">2) 特定のファイルでマクロを有効にする方法</p>
特定のファイルでマクロを有効にする方法ついて、以下の "Web マークを削除する方法" 及び "信頼できる発行元のデジタル署名を追加する方法" があります。
### Web マークを削除する方法
インターネットから入手したファイルには Web マークが付与されると先ほど述べましたが、この Web マークの実体はWindows OS で広く利用されている NTFS ファイルシステムの代替データストリーム（ADS：Alternate Data Stream）に保存されているゾーン識別子 (Zone.Identifier) です。

そのため、このゾーン識別子を削除すると、本動作変更の対象外となり、ローカルで作成したマクロ付きファイルと同様の動作となります。以下画面から削除する手順と PowerShell で削除する手順をそれぞれ記載します。

#### <ゾーン識別子の削除手順 (画面) >
1.  エクスプローラーから該当のマクロ付きファイルを右クリックし、[プロパティ]を表示します。
1. [全般] タブの最下部に以下画面のように [セキュリティ] が表示されている場合は、右側の [許可する] をチェックし、[OK] をクリックします。 

![](file-property-security.png)


#### <ゾーン識別子の削除手順 (PowerShell コマンド) >  
以下のコマンドを PowerShell で実行します。  (対象ファイル : C:\temp\sample.xlsm)
```
Remove-Item 'C:\temp\sample.xlsm' -Stream Zone.Identifier;  
```

<br>

### <p id="EnableVBA-Sign">信頼できる発行元のデジタル署名を追加する方法</a>
Office アプリケーション実行環境にて、"信頼できる発行元" として登録されている証明書を使用したデジタル署名がマクロに追加されている場合にはマクロが実行可能となります。

この動作を利用してマクロを実行可能とするため、署名者は秘密鍵を持つ証明書で VBA マクロに署名を行い、この証明書をマクロ実行環境のユーザーの「信頼された発行元」ストアに登録します。証明書の信頼に関する動作詳細は Windows 全般の動作となるため本記事では割愛し、Office の動作に関わる部分のみ以下に記載します。

以下に各 Office アプリケーション上での "信頼できる発行元" として登録されている証明書の確認方法を記載します。

#### <オプション設定 - 信頼できる発行元>  
[ファイル] タブ - [オプション] - [トラスト センター] - [トラスト センターの設定] -  [信頼できる発行元]

![](trusted-publisher.png)

上記画面にて "信頼できる発行元" として登録されているデジタル証明書の一覧を確認することが可能です。  

マクロにデジタル署名を付与する手順は以下となります。

#### <デジタル署名追加手順>
1. マクロ付きファイルを各 Officeアプリケーションで開きます。
1. [Alt] + [F11] キーを を押下し、Visual Basic Editor (VBE) を開きます。
1. VBE の [ツール] タブ - [デジタル署名] をクリックします。
1. 表示されたウィンドウの [選択] をクリックし、使用する "信頼できる発行元" の証明書を選択して [OK] をクリックします。
1. "この VBA プロジェクトは以下のように署名されています。" の下の "証明書名" に選択した証明書名が表示されたことを確認し、[OK] をクリックします。
1. ファイルを保存して閉じます。

![](digital-signature.png)

<br>

## <p id="EnableVBA-3">3) 特定のフォルダーのファイルでマクロを有効にする方法</p>
各アプリケーションの以下オプション設定で確認 / 追加可能な "信頼できる場所" に登録されているパスにマクロ付きファイルを保存をすることにより、マクロが実行可能となります。

#### <オプション 設定>  
[ファイル] タブ - [オプション] - [トラスト センター] - [トラスト センターの設定] -  [信頼できる場所]
注) 既定で登録された場所は特殊なフォルダですので、マクロを実行したいファイルを保存するためのフォルダは [新しい場所の追加] から任意の場所を追加することをお勧めします。

![](trusted-location.png)

注) Access / Excel / PowerPoint / Word / Visio / Project の各アプリケーションでそれぞれ異なる設定となります。また、Publisher には [信頼できる場所] 機能はないため、この対応方法は利用できません。

なお、本設定は以下グループ ポリシーでも追加 / 削除が可能です。

#### <グループポリシー設定 - 信頼できる場所>
- Access  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft Access 2016] - [アプリケーション設定] - [セキュリティ]  
\- [セキュリティ センター] - [信頼できる場所]  
\- [信頼できる場所 #1 - #20] : 有効 且つ [パス] に C:\temp などのフォルダーパスを登録します。
- Visio / Project  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft xxx 2016] - [xxx のオプション] -[セキュリティ]  
\- [セキュリティ センター]  
\- [信頼できる場所 #1 - #20] : 有効 且つ [パス] に C:\temp などのフォルダーパスを登録します。  
注) xxx は それぞれ Visio / Project です。 
- Excel / PowerPoint / Word  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft xxx 2016] - [xxx のオプション] -[セキュリティ]  
\- [セキュリティ センター] - [信頼できる場所]  
\- [信頼できる場所 #1 - #20] : 有効 且つ [パス] に C:\temp などのフォルダーパスを登録します。  
注) xxx は それぞれ Excel / PowerPoint / Word です。
<br>

なお、以下のポリシーで Office アプリケーション全体に一括設定することもできます。  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft Office 2016] - [セキュリティ] - [セキュリティ センター]  
\- [信頼できる場所] - [信頼できる場所 #1 - #20] : 有効 且つ [パス] に C:\temp などのフォルダーパスを登録します。

また、[信頼できる場所] 下部に表示されている [自分のネットワーク上にある信頼できる場所を許可する (推奨しません)] に対応するグループ ポリシーは以下となります。共有フォルダーなどネットワーク上のパスを指定する場合はこちらをご利用ください。(この設定は、Office 全体への一括設定はありません。)

#### <グループポリシー設定 - 自分のネットワーク上にある信頼できる場所を許可する>
- Access  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft Access 2016] - [アプリケーション設定] - [セキュリティ]  
\- [セキュリティ センター] - [信頼できる場所] - [ネットワーク上の信頼できる場所を許可する]
- Visio / Project  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft xxx 2016] - [xxx のオプション] - [セキュリティ]  
\- [セキュリティ センター] - [ネットワーク上の信頼できる場所を許可する]  
注) xxx は それぞれ Visio / Project です。
- Excel / PowerPoint / Word  
[ユーザーの構成] - [ポリシー] - [管理用テンプレート]  
\- [Microsoft xxx 2016] - [xxx のオプション] - [セキュリティ]  
\- [セキュリティ センター] - [信頼できる場所] - [ネットワーク上の信頼できる場所を許可する]  
注) xxx は それぞれ Excel / PowerPoint / Word です。

<br>

## <p id="EnableVBA-4">4) Web ・ネットワーク共有上のファイルを開くシナリオで、信頼済みサイトに登録する方法</p>
自社 Web サイトや社内ファイル サーバーなど、安全が保証されている場所の Office ファイルを開く場合は、Web サイトまたはサーバーを信頼済みサイトに登録すると、インターネット入手ファイルとみなされずマクロを有効にできます。

#### <信頼済みサイト登録手順>
1. タスクバーにある検索ボックスに "インターネット オプション" と入力します。
1. 表示された [インターネット オプション] をクリックし、[インターネットのプロパティ] を開きます。
1. [セキュリティ] タブ - [信頼済みサイト] を選択し [サイト] をクリックします。
1. マクロの実行を許可する Office ファイルが格納されたサイト・サーバーの URL を入力して [追加] をクリックします。 
   <登録内容>  
   Web サイト : URL (例: https://www.microsoft.com)  
   ネットワーク共有 : 接続時に指定している相手先サーバーの FQDN (例: \\\server.domain.com) または IP アドレス (例: 111.111.111.111)  
   ※ http:// で始まる URL や、ネットワーク共有を追加する場合は、下部の [このゾーンのサイトにはすべてサーバーの確認 (https:) を必要とする] のチェックを外します。
1. [閉じる] をクリックし、[信頼済みサイト] 画面を閉じます。
1. [インターネットのプロパティ] 画面で [OK] をクリックして画面を閉じます。

![](trusted-sites.png)

補足)  
Internet Explorer の場合は、Web サイトの応答ヘッダーが Content-Disposition: Attachment のときはインターネット一時フォルダにファイルがダウンロードされます。このフォルダに格納されたファイルは、Office ではすべてインターネット入手ファイルとみなすため、信頼済みサイトに登録してもマクロはブロックされます。

Edge ブラウザーの場合は、一時フォルダにダウンロードされたファイルも、サイトのゾーン設定に従って Web マークが付与され、それによってマクロのブロック有無が決定されるため信頼済みサイトへの登録でマクロを有効にできるようになります。


## ご参考
Office グループポリシー管理用テンプレートの導入手順について、以下に詳細が記載されていますので、こちらをご確認ください。

タイトル : Office の管理用テンプレートを使用してグループ ポリシー (GPO) で Office 365 ProPlus を制御する  
アドレス  : [https://answers.microsoft.com/ja-jp/msoffice/forum/all/office/3ec9d79c-44ec-4273-97e2-2a6f3a1fd8ef](https://answers.microsoft.com/ja-jp/msoffice/forum/all/office/3ec9d79c-44ec-4273-97e2-2a6f3a1fd8ef)  

また、上記以外に Microsoft 365 管理センターより設定可能な Office にサインインしているユーザーに適用するクラウド ポリシーや Microsoft Intune を利用したグループ ポリシーを構成する方法も存在します。ご利用の Office バージョンや環境に応じてこれらもご利用ください。以下に公開情報を記載します。

タイトル : Microsoft 365 Apps for enterprise 向け Office クラウド ポリシー サービスの概要  
アドレス  : [https://docs.microsoft.com/ja-jp/deployoffice/admincenter/overview-office-cloud-policy-service](https://docs.microsoft.com/ja-jp/deployoffice/admincenter/overview-office-cloud-policy-service) (日本語版 - 機械翻訳)    
アドレス  : [https://docs.microsoft.com/en-us/deployoffice/admincenter/overview-office-cloud-policy-service](https://docs.microsoft.com/en-us/deployoffice/admincenter/overview-office-cloud-policy-service) (英語版)  

タイトル : Windows 10/11 テンプレートを使用して、Microsoft Intune でグループ ポリシー設定を構成する  
アドレス  : [https://docs.microsoft.com/ja-jp/mem/intune/configuration/administrative-templates-windows](https://docs.microsoft.com/ja-jp/mem/intune/configuration/administrative-templates-windows) (日本語版 - 機械翻訳)    
アドレス  : [https://docs.microsoft.com/en-us/mem/intune/configuration/administrative-templates-windows](https://docs.microsoft.com/en-us/mem/intune/configuration/administrative-templates-windows) (英語版)    

<br>

今回の投稿は以上です。

<br>

**本情報の内容（添付文書、リンク先などを含む）は、作成日時点でのものであり、予告なく変更される場合があります。**



