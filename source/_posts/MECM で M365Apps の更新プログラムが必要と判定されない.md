---
title: MECM で M365Apps の更新プログラムが必要と判定されない
date: '2023-12-26'
lastupdate: '2023-12-26'
id:clqm07fq300005kkj5ww5bp4a
tags:
  - 更新

---

こんにちは、Office サポートの西川 (直)です。  

MECM (Microsoft Endpoint Configuration Manager) で M365Apps (Microsoft 365 Apps for enterprise) の更新プログラムが必要と判定されない現象について説明します。

<br>

現象
--
MECM で M365Apps の更新プログラムが必要と判定されず、更新が適用できない


<br>

対処方法
--
以下の 3 つの条件を全て満たしているかをご確認ください。

**1) Office の更新に使用する COM IF のレジストリキーが存在する**
`HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{B7F1785F-D69B-46F1-92FC-D2DE9C994F13}`

本レジストリは、[MECM から更新プログラムを受信できるようにする設定](https://learn.microsoft.com/ja-jp/deployoffice/updates/manage-microsoft-365-apps-updates-configuration-manager#enable-microsoft-365-apps-clients-to-receive-updates-from-configuration-manager)が反映された後、OS を再起動することで作成されます。

<br>

**2) VersionToReport より新しいビルドを展開しようとしている、または、UpdateChannelChanged が True である**
`HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuratio\VersionToReport`
ここには、16.0.12345.12345 のような 16.0 部分は固定のビルド番号が入ります。

`HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuratio\UpdateChannelChanged`
初期値は False となります。後述のように UpdateChannel が優先度により変更された際に True となります。

通常、チャネルの変更直後を除いて UpdateChannelChanged は False のため、VersionToReport の優劣にしたがい更新が行われます。

<br>

**3) UpdateChannel の URL に展開しているチャネルの識別子が含まれる**
`HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuratio\UpdateChannel`

***チャネルの識別子***
`最新チャネル : 492350f6-3a01-4f97-b9c0-c7c6ddf67d60`
`月次エンタープライズチャネル : 55336b82-a18d-4dd6-b5f6-9e5095c314a6`
`半期エンタープライズチャネル (プレビュー : b8f9b850-328d-4355-9145-c59439a0c4cf`
`半期エンタープライズチャネル : 7ffbc6bf-bc32-4f92-8982-f9dd17fd3114`

UpdateChannel は UpdateChannelChanged が False である場合、以下の優先度にしたがい値が評価されます。

優先度1) [GPO の "更新プログラムのパス"](https://learn.microsoft.com/ja-jp/deployoffice/updates/configure-update-settings-microsoft-365-apps)
`HKEY_LOCAL_MACHINE\software\policies\microsoft\office\16.0\common\OfficeUpdate\updatepath`

優先度2) [GPO の "チャネルの更新"](https://learn.microsoft.com/ja-jp/deployoffice/updates/configure-update-settings-microsoft-365-apps)
`HKEY_LOCAL_MACHINE\software\policies\microsoft\office\16.0\common\OfficeUpdate\updatebranch`

優先度3) [Office 展開ツールの Updates 要素の UpdatePath 属性](https://learn.microsoft.com/ja-jp/deployoffice/office-deployment-tool-configuration-options#updates-element)
`HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration\UpdateUrl`

優先度4) [Microsoft 365 管理センターで Microsoft 365 のインストール オプションを管理する](https://learn.microsoft.com/ja-jp/deployoffice/manage-software-download-settings-office-365) の取得頻度
`HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration\UnmanagedUpdateUrl`

優先度5) [Office 展開ツールの Updates 要素の Channel 属性](https://learn.microsoft.com/ja-jp/deployoffice/office-deployment-tool-configuration-options#updates-element)
`HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration\CDNBaseURL`

現在の UpdateChannel と値が異なる場合、評価された値に上書きされ、UpdateChannelChanged が True となります。

<br>

FAQ
--
**UpdateChannel が変更されるタイミングを教えてください**
タスクスケジューラーに登録されている Office Automatic Updates 2.0 が実行する OfficeC2RClient.exe の中で、前回の構成試行より 24 時間以上経過している場合に構成処理を行います。
このため、タスクスケジューラーを手動実行いただいても即座に実行することは出来ません。また、この時間を変更することは出来ません。

**優先度1 や 3 が入っている端末に MECM から更新プログラムを展開できますか？**
優先度1 は "未構成" に変更してください。優先度 3 の場合、UpdateUrl レジストリを削除していただくか、優先度 2 の GPO の "チャネルの更新" を代わりにご利用ください
なお、"チャネルの更新" を設定すると、MECM もしくはインターネット上の Office CDN から更新プログラムを取得するようになります。MECM で管理しない端末に設定する際はご留意ください。


**UpdateChannelChanged が既に True になっているため、優先度にしたがって設定を変更しても端末に反映されません**
UpdateChannelChanged を一度だけ False に書き換えてください。


**優先度4の UnmanagedUpdateUrl のチャネルを変更したいのですが、Microsoft 365 管理センターで半期エンタープライズチャネルを選択できません**
UnmanagedUpdateUrl レジストリを削除していただくか、優先度 2 の GPO の "チャネルの更新" を代わりにご利用ください
なお、"チャネルの更新" を設定すると、MECM もしくはインターネット上の Office CDN から更新プログラムを取得するようになります。MECM で管理しない端末に設定する際はご留意ください。

**MECM から更新しているのに、優先度4 の UnmanagedUpdateUrl レジストリが設定されているのはなぜですか？**
[方法3](https://learn.microsoft.com/ja-jp/deployoffice/updates/manage-microsoft-365-apps-updates-configuration-manager#enable-microsoft-365-apps-clients-to-receive-updates-from-configuration-manager) の Office 展開ツールでのインストール時に OfficeMgmtCOM を True で設定していない場合、端末に MECM から更新する設定が反映される前に M365Apps の自動更新が動作した可能性があります。


**COM IF のレジストリキーが作成されません**
何らかの影響により、端末で COM 登録が行えていない可能性があります。

対象の端末で、PowerShell を管理者権限で起動し、以下の 2 行を実行いただき、エラーが出るかをご確認ください。

```
$comCatalog = New-Object -ComObject COMAdmin.COMAdminCatalog
$appColl = $comCatalog.GetCollection("Applications")
```

**エラーが出る場合**
OS 観点で調査が必要であると考えられます。

**エラーが出ない場合**
以下の手順をご実施ください。
 
1. 以下の REG_DWORD レジストリに 0 を設定します。
`HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate\OfficeMgmtCOM`

2. スタートメニューの検索にて、「サービス」と検索し、サービスを起動します
3. "Microsoft Office クイック実行サービス" を右クリックし、"再起動" を実施します。

4. 以下の REG_DWORD レジストリに 1 を設定します。
`HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate\OfficeMgmtCOM`

5. "Microsoft Office クイック実行サービス" を右クリックし、"再起動" を実施します。
6. COM IF のレジストリキーが作成されるかをご確認ください。


<br>

参考資料
--
[How to manage Office 365 ProPlus Channels for IT Pros](https://techcommunity.microsoft.com/t5/microsoft-365-blog/how-to-manage-office-365-proplus-channels-for-it-pros/ba-p/795813)

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**