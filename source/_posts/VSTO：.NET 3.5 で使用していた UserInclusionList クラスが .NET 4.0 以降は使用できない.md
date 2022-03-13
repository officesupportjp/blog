---
title: VSTO：.NET 3.5 で使用していた UserInclusionList クラスが .NET 4.0 以降は使用できない
date: 2019-03-01
id: cl0mdnodx002r5ovsfzvw5fh5
alias: /VSTO：.NET 3.5 で使用していた UserInclusionList クラスが .NET 4.0 以降は使用できない/
---

(※ 2017 年 2 月 16 日に Japan Office Developer Support Blog に公開した情報のアーカイブです。)

こんにちは、Office 開発 サポート チームの多田です。

今回は、.NET 3.5 で使用していた UserInclusionList クラスが .NET 4.0 以降では使用できない件、およびその対処策についてご案内します。

#### **はじめに**

VSTOソリューションをインストール時や、初回起動時にユーザーに信頼するかどうかの確認ダイアログを表示させたくない、といった理由から、VSTO ソリューションへの信頼をあらかじめ付与したい場合があります。信頼を付与する方法として、従来、VSTO では UserInclusionList クラスおよび AddInSecurityEntryクラスが用意されていました。

参考) 方法: 信頼のリストのエントリを追加または削除する  
[https://msdn.microsoft.com/ja-jp/library/bb398239.aspx](https://msdn.microsoft.com/ja-jp/library/bb398239.aspx)

UserIncrusionList クラスおよび AddInSecurityEntry クラスを含む Microsoft.VisualStudio.Tools.Office.Runtime.v10.0.dll は、.NET Framework 3.5 用の Office 拡張機能アセンブリです。

このため、.NET Framework 3.5 ターゲットの VSTO テンプレートを含まない Visual Studio 2012 以降がインストールされた開発環境で、かつ .NETFramework 3.5 が端末にインストールされていない場合は、Microsoft.VisualStudio.Tools.Office.Runtime.v10.0.dll アセンブリを利用できません。また、インストール先環境に .NET Framework 3.5 がインストールされていない場合も、同様に利用できません。 

参考)  
タイトル : Visual Studio Tools for OfficeRuntime のアセンブリ  
アドレス : [https://msdn.microsoft.com/ja-jp/library/ee712616.aspx](https://msdn.microsoft.com/ja-jp/library/ee712616.aspx)

##### **背景**

.NET Framework 4.0 以降をターゲットとする Office 2010 以降の VSTO では、ユーザーのコンピューターの Program Files ディレクトリは既定で信頼されます。また、.NET 4.0 以降の OfficeではClickOnce による配布を推奨しているため、セットアップ プロジェクトで InclusionList を追加するのではなく、ClickOnce のマニフェストに適切な証明書で署名を行ったり、ドキュメント カスタマイズ ソリューションの場合は Office のオプション設定から信頼された場所を追加することを想定しています。

これらの方法で信頼を付与するようになったため、.NET 4.0 以降においては、UserIncrusionList クラスおよび AddInSecurityEntry クラスが削除されています。 

参考)  
タイトル : Windows インストーラーを使用した Office ソリューションの配置  
アドレス : [https://msdn.microsoft.com/ja-jp/library/cc442767.aspx](https://msdn.microsoft.com/ja-jp/library/cc442767.aspx)

タイトル : ClickOnce を使用した Office ソリューションの配置  
アドレス : [https://msdn.microsoft.com/ja-jp/library/bb772100.aspx](https://msdn.microsoft.com/ja-jp/library/bb772100.aspx)

#### 

**対処策**

1\. Program Files 配下にインストールする

2\. 信頼できる証明書でマニフェストに署名する

3\. 信頼のリストの情報を保持するレジストリを直接登録する

4\. 開発環境および運用環境に .NET Framework 3.5 をインストールする

##### **1\. Program Files** **配下にインストールする**

先述の通り、Program Files ディレクトリ配下にインストールされる VSTO ソリューションは既定で信頼されます。インストール先を固定できる場合は、この方法を用いることで信頼のリストを追加する必要はありません。

##### 

**2\.** **信頼できる証明書でマニフェストに署名する**

VSTO ソリューションをビルドすると、マニフェストに署名が行われます。署名に使用する証明書をインストール先環境で \[信頼された発行元\] ストアに登録すると、確認ダイアログの表示無しに VSTO ソリューションをインストールできます。

参考)  
Office ソリューションへの信頼の付与  
[https://msdn.microsoft.com/ja-jp/library/bb772086.aspx](https://msdn.microsoft.com/ja-jp/library/bb772086.aspx)

**3\. 信頼のリストの情報を保持するレジストリを直接登録する**

UserInclusionList や AddInSecurityEntry クラスを用いて信頼のリストをインストーラー プログラムで追加した場合、最終的には、レジストリに情報が追加されます。このレジストリ情報が存在することにより、Office アプリケーションは当該 VSTO ソリューションを信頼します。

Windows インストーラーやグループポリシー等その他の方法で、以下のレジストリ情報を直接追加することで VSTO アプリケーションを信頼できます。

<レジストリ情報\>

キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\VSTO\\Security\\Inclusion\\\<GUID\>

※ \<GUID\> は他と重複しない任意の値です。検証端末などで信頼リストを設定せず Program Files 配下 "以外" にインストールすると、自動的に作成されます。この方法で自動的に作成されたレジストリ キーの GUID の値をコピーするなどして決定してください。

\- 設定値 1  
名前 : Url  
種類 : REG\_SZ  
値 : file:///\[インストール先ディレクトリ\]\<VSTO アプリケーションの配置マニフェスト名\>.vsto

例) [file:///C:/work/SampleExcelAddIn.vsto](../../../work/SampleExcelAddIn.vsto)

\- 設定値 2  
名前 : PublicKey  
種類 : REG\_SZ  
値 : マニフェスト ファイル (.manifest または .vsto) 内の \<RSAKeyValue\> セクションの値

イメージ) \<RSAKeyValue\>\<Modulus\>xxxxx\</Modulus\>\<Exponent\>AQAB\</Exponent\>\</RSAKeyValue\>

**4\. .NET Framework 3.5 をインストールする**

VSTO ソリューションのターゲット フレームワークが .NET Framework 4.0 以上であっても、開発環境およびインストール先環境に .NET Framework 3.5 がインストールされていれば、Microsoft.VisualStudio.Tools.Office.Runtime.v10.0.dll を参照して UserInclusionList / AddInSecurityEntryクラスを利用できます。

なお、Visual Studio 2010 Tools for Office Runtime のインストール時にアセンブリの登録を行いますので、これより先に .NET Framework 3.5 をインストールするようご注意ください。(Visual Studio 2010 Tools for Office Runtime は Office や Visual Studio のインストール時に合わせてインストールされます。)

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**