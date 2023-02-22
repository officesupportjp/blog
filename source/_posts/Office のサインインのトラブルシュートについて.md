---
title: Office のサインインのトラブルシュートについて
date: '2020-07-09'
lastupdate: '2021-05-10'
id: cl0m6umvg001gi4vsgppfcmk0
tags:
  - サインイン、認証

---

こんにちは、Office サポートの西川 (直) です。 

特定のユーザーだけ Office のサインインが成功しないとお問い合わせいただくことがございます。
この場合、以下のサインインおよびライセンス情報のリセットをお試しいただくことで、解消する可能性がございます。

[Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state)

ご利用されている製品によって、以下を指針としてご対応ください。

<br>

Office 2013、2016 の場合
--
1. Office からサインアウトし、Office を終了します。
2. [Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state) から "OLicenseCleanup.zip" をダウンロードして任意のフォルダに解凍します。
3. ダブルクリックで OLicenseCleanup.vbs を実行します。処理はバックグラウンドで終了いたします。
※ 管理者権限は不要です。他の管理者ユーザーで実行すると正しくキャッシュが削除出来ませんので事象が発生しているログオンユーザーの権限で実行してください。

<br>

Office 2019 以降の場合
--
1. Office からサインアウトし、Office を終了します。Teams、OneDrive 同期クライアントも全て終了します。
2. [Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state) から "OLicenseCleanup.zip" をダウンロードして任意のフォルダに解凍します。
3. ダブルクリックで OLicenseCleanup.vbs を実行します。処理はバックグラウンドで終了いたします。
※ 管理者権限は不要です。他の管理者ユーザーで実行すると正しくキャッシュが削除出来ませんので事象が発生しているログオンユーザーの権限で実行してください。

4. 以下の \[職場または学校にアクセスする\] からアカウントを切断してください。
※ Azure AD や AD に接続済みの項目を "切断" しますと、ドメインから外れてしまいますので、Azure AD や AD に接続済み の項目は切断しないようにご注意ください。

![](image01.png)  

5. 以下の [BrokerPlugin プロセスを確認する](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/sign-in-issues) の 1 から 5 の手順を実施します。

```
1. エクスプローラーを開き、アドレス バーに次の場所を入力します。%LOCALAPPDATA%\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts
2. Ctrl + A キーを押して、すべてを選択します。
3. 選択したファイルを右クリックして、[削除] を選択します。
4. エクスプローラーのアドレス バーに次の場所を入力します。%LOCALAPPDATA%\Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts
5. すべてのファイルを選択して削除します。
```
※ 事象が発生しているユーザーにてご実施ください。

6. OS からサインアウト、または、OS を再起動してください

<br>

Microsoft 365 Apps の場合
--
1. Office からサインアウトし、Office を終了します。Teams、OneDrive 同期クライアントも全て終了します。
2. [Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state) から "OLicenseCleanup.zip" をダウンロードして任意のフォルダに解凍します。
3. コマンドプロンプトを右クリック→管理者権限で起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。
ログオンユーザーに管理者権限がない場合、スクリプトを 2 回実行する必要があります。ログオンユーザーではダブルクリックで OLicenseCleanup.vbs を実行します。その後、コマンドプロンプトを右クリック→管理者権限として別の管理者ユーザーのアカウントで起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。
ログオンユーザーに管理者権限がなく別の管理者ユーザーも利用できない場合、まずはダブルクリックで OLicenseCleanup.vbs を実行して事象が改善するかをご確認ください。

4. 以下の \[職場または学校にアクセスする\] からアカウントを切断してください。
※ Azure AD や AD に接続済みの項目を "切断" しますと、ドメインから外れてしまいますので、Azure AD や AD に接続済み の項目は切断しないようにご注意ください。

![](image01.png)  

5. 以下の [BrokerPlugin プロセスを確認する](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/sign-in-issues) の 1 から 5 の手順を実施します。

```
1. エクスプローラーを開き、アドレス バーに次の場所を入力します。%LOCALAPPDATA%\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts
2. Ctrl + A キーを押して、すべてを選択します。
3. 選択したファイルを右クリックして、[削除] を選択します。
4. エクスプローラーのアドレス バーに次の場所を入力します。%LOCALAPPDATA%\Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts
5. すべてのファイルを選択して削除します。
```
6. OS からサインアウト、または、OS を再起動してください

<br>

**補足**
--
- Office Home & Business 2016 等、コンシューマー製品の場合には "ボリュームライセンス版の Office 2019 以降の場合" と同じ対処をご実施ください。

- 「職場または学校にアクセスする」に何らかの問題が生じている、または、コマンド等でアカウントを切断する場合、以下の手順をお試しください。
1. 以下のサイトから WPJCleanUp.zip を対象のクライアント端末にダウンロードします。  
[https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip](https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip)  
2. クライアント端末に ＜対象ユーザー＞ でログオンし、WPJCleanUp.zip を任意のフォルダに展開します。
3. CleanupTool に含まれている WPJCleanp.cmd を実行します。
※ダブルクリックにて実行時、\[Windows によって PC が保護されました\] と表示された場合は \[詳細情報\] をクリックし、表示された \[実行\] をクリックします。  

- [Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state) の Signoutofwamaccounts.ps1 が文字化けしている場合、以下をご利用ください。

```
if(-not [Windows.Foundation.Metadata.ApiInformation,Windows,ContentType=WindowsRuntime]::IsMethodPresent("Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager", "FindAllAccountsAsync"))
{
    throw "This script is not supported on this Windows version. Please, use CleanupWPJ.cmd."
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime

Function AwaitAction($WinRtAction) {
  $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod })[0]
  $netTask = $asTask.Invoke($null, @($WinRtAction))
  $netTask.Wait(-1) | Out-Null
}

Function Await($WinRtTask, $ResultType) {
  $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  $netTask.Result
}

$provider = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager,Windows,ContentType=WindowsRuntime]::FindAccountProviderAsync("https://login.microsoft.com", "organizations")) ([Windows.Security.Credentials.WebAccountProvider,Windows,ContentType=WindowsRuntime])

$accounts = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager,Windows,ContentType=WindowsRuntime]::FindAllAccountsAsync($provider, "d3590ed6-52b3-4102-aeff-aad2292ab01c")) ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult,Windows,ContentType=WindowsRuntime])

$accounts.Accounts | % { AwaitAction ($_.SignOutAsync("d3590ed6-52b3-4102-aeff-aad2292ab01c")) }
```

<br>

**\- 関連情報**
[Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について](https://officesupportjp.github.io/blog/cl0m75al4001gmcvse36w64dv/index.html)

<br>

<span style="color:#ff0000">**2020/10/28  Update**</span>  
<span style="color:#339966">\[職場または学校にアクセスする\] のアカウントの削除についてを追加しました</span>

<span style="color:#ff0000">**2021/5/10  Update**</span>  
<span style="color:#339966">\ログオンユーザーに管理者権限がない場合の手順を変更しました。</span>

<span style="color:#ff0000">**2022/3/16  Update**</span>  
<span style="color:#339966">キャッシュ削除手順を追加しました。</span>

<span style="color:#ff0000">**2023/2/22  Update**</span>  
<span style="color:#339966">記事を簡略化しました。</span>


**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**