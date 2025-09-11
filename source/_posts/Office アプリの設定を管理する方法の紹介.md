---
title: Office アプリの設定を管理する方法の紹介
date: '2025-09-11'
id: cmff530z3000atsse9c4navwb
tags: 
 - Office アプリ全般
 - 運用管理
---

こんにちは、Office サポート チームの大津です。  

今回の記事では、特にお問い合わせの多い Windows 版の Office デスクトップアプリと Office for the Web を対象に、企業や組織の IT 管理者がユーザーの Office アプリの動作や機能を管理・制御する一般的な方法についてご紹介します。

# 1. 一般的な Office 向け設定の構成方法
## ■ Microsoft 365 クラウドポリシー
Microsoft 365 クラウド ポリシーは、Microsoft 365 Apps 管理センター画面から Microsoft 365 Apps へのサインイン ユーザーに対するポリシーを構成できます。
デバイスがドメインや Intune に参加していない場合でも Microsoft 365 Apps サインイン ユーザーに対してポリシーを構成することができます。構成方法や詳細情報は以下の公開情報をご確認ください。

参考：Microsoft 365 向けクラウドポリシーサービスの概要  
URL：[https://learn.microsoft.com/ja-jp/microsoft-365-apps/admin-center/overview-cloud-policy](https://learn.microsoft.com/ja-jp/microsoft-365-apps/admin-center/overview-cloud-policy)

## ■ Active Directory のグループポリシー

Active Directory (AD) のグループポリシー (GPO) では、ドメインに参加していて Windows OS を使用しているデバイスに対してポリシーを構成することができます。Office に対してポリシーを構成する場合、IT 管理者は Office の管理用テンプレートをダウンロードして構成する必要があります。詳細な手順は以下の公開情報をご覧ください。

参考：Office の管理用テンプレートを使用してグループポリシー (GPO) で Office 365 ProPrus を制御する  
URL：[https://learn.microsoft.com/ja-jp/answers/questions/4370544/office-(gpo)-office-365-proplus](https://learn.microsoft.com/ja-jp/answers/questions/4370544/office-(gpo)-office-365-proplus)

## ■ Intune 内 Microsoft 365 アプリのポリシー→ Microsoft 365 クラウドポリシーと同じ

Intune 管理センターでも Microsoft 365 クラウドポリシーを構成することができます。具体的な手順は以下の公開情報を参考にしてください。

参考：Microsoft 365 アプリのポリシー  
URL：[https://learn.microsoft.com/ja-jp/mem/intune/apps/app-office-policies](https://learn.microsoft.com/ja-jp/mem/intune/apps/app-office-policies)

## ■ Intune 設定カタログ

Intune 設定カタログでは、様々な機能や製品のポリシーを構成でき、Office のポリシーも含まれています。具体的な手順は以下の公開情報を参考にしてください。

参考：Microsoft Intune で設定カタログを使用してポリシーを作成する  
URL：[https://learn.microsoft.com/ja-jp/intune/intune-service/configuration/settings-catalog](https://learn.microsoft.com/ja-jp/intune/intune-service/configuration/settings-catalog)

<Intune について>  
Intune は企業や組織が従業員のデバイスを一元管理できるクラウドベースのエンドポイント管理サービスです。設定カタログを利用する場合は、ポリシーを構成したいユーザーのデバイスを Intune に登録しておく必要があります。

参考: 登録ガイド: Microsoft Intune登録  
[https://learn.microsoft.com/ja-jp/intune/intune-service/fundamentals/deployment-guide-enrollment](https://learn.microsoft.com/ja-jp/intune/intune-service/fundamentals/deployment-guide-enrollment)

## ■ Office 展開ツール (ODT) の構成ファイル AppSetting で「初期値」を設定

Microsoft 365 Apps インストール時や、または、インストール後に ODT を追加実行して、構成ファイルで指定した設定を、実行したデバイスに追加できます。他のポリシー設定とは異なり、既定値 (初期値) として構成され、強制することはできません。対応するオプション メニュー等からユーザーが変更できます。
具体的な手順は以下の公開情報を参考にしてください。

参考：Office 展開ツールの概要  
URL：[https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/overview-office-deployment-tool](https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/overview-office-deployment-tool)  
※「アプリケーションの基本設定をMicrosoft 365 Appsに適用する」セクションを参照

参考：Office 展開ツールのオプションの構成  
URL：[https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/office-deployment-tool-configuration-options](https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/office-deployment-tool-configuration-options)  
※ 「AppSettings 要素」セクションを参照

# 2. 構成方法の比較

各 Office 向け設定の構成方法について、特徴の違いを以下の表にまとめます。これを参考にどの方法を採用するかをご検討ください。

|  | Microsoft 365 クラウドポリシー | Active Directory のグループポリシー | Intune 設定カタログ | ODT の構成ファイルで初期設定 |
|------------|-------------------------------|--------------------------------------|------------------------|------------------------------|
|適用対象※1|Microsoft 365 Apps サインイン ユーザー|Windows デバイスまたはユーザー| Windows デバイスまたはユーザー| Windows デバイス|  
|前提| Microsoft 365 Apps にサインイン|デバイスやユーザーが AD の管理対象|デバイスを Intune に登録|特になし|  
|強制または任意|強制|強制|強制|任意|  
|優先順位|1 位|2 位|2 位|3 位|
|使用可能な製品<br>注：ポリシーはエンタープライズ向け製品でのみサポートされます|Microsoft 365 for enterprise|Microsoft 365 Apps for enterprise、ボリューム ライセンス版 Office| Microsoft 365 Apps for enterprise、ボリュームライセンス版 Office |クイック実行形式の Office (Microsoft 365 Apps、ボリュームライセンス版 Office 2019 以降、製品版 Office 2016 以降)|  
|使用可能なプラットフォーム|Windows 上の Microsoft 365 デスクトップ アプリ、Office for the web|Windows 上の Office デスクトップ アプリ|Windows 上の Office デスクトップ アプリ|Windows 上の Office デスクトップ アプリ|
|反映にかかる時間|最大 24 時間|90 ~ 120 分|約 8 時間ごとに更新|ODT 実行後 Office 起動時|  
|手動同期方法|なし|コンピューターのポリシー：クライアントの再起動<br>ユーザーのポリシー：Windows への再ログオン<br>または<br>gpupdate コマンド実行|管理者が Intune 管理画面から同期<br>または<br>クライアントで Windows の設定画面内 [アカウント] から同期|なし|  
|構成に必要な権限|Office Apps 管理者 (推奨)、セキュリティ管理者、グローバル管理者|Active Directory ドメイン管理者、エンタープライズ管理者|ポリシーとプロファイルマネージャーの役割|実行マシンのローカル管理者権限|
|インターネット環境要否|必要|必要なし|必要|必要なし|  

※1: 適用対象の動作詳細  
Office の動作設定のほとんどは、ユーザー向けに構成します。(セットアップ関連設定などのコンピューター向け設定も一部存在します。)  
Active Directory のグループポリシーは、ユーザー向けのポリシーは、ユーザーに対してのみ構成できます。  
Intune 設定カタログは、ユーザー向けのポリシーをデバイスに割り当てることもできます。この場合、そのデバイス上のすべてのサインイン ユーザーに反映されます。  
Office 展開ツールも、デバイス上で実行すると、そのデバイスにサインインするすべてのユーザーに反映されます。
<br>

今回の投稿は以上です。ご紹介した各種管理方法の中から、組織の環境や運用方針に最も適したものを選択いただくことで、Office アプリをより効率的に活用していただけるかと思います。  
今後も、皆様の業務に役立つ情報を継続的に発信してまいりますので、ぜひ当ブログをご覧ください。
<br>

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**

