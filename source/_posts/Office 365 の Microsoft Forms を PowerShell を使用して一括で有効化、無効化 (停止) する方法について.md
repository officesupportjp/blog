---
title: Office 365 の Microsoft Forms を PowerShell を使用して一括で有効化、無効化 (停止) する方法について
date: 2019-02-19
lastupdate: 2019-03-22
---

(※ 2017 年 6 月 29 日に Office Support Team Blog JAPAN に公開した情報のアーカイブです。)  
  
**2017/06/19 Update! Office 365 Enterprise パックのライセンスにも対応したコマンドに更新**  
  
こんにちは、Office サポートの佐村です。  
本記事では Office 365 にて公開されている Microsoft Forms (以下、Forms) について、PowerShell スクリプトを使用して一括でテナント全体のユーザーのライセンスを有効化、無効化する方法をご案内いたします。  
本方法をご利用いただくことで、別サービスの無効化状況には影響せずに、Forms だけを有効化、もしくは、無効化することが可能です。  
  
**目次**  
1\. Forms のライセンスについて  
2\. PowerShell を使用して一括で有効化する方法  
3\. PowerShell を使用して一括で無効化する方法  
4\. 関連情報  
  
**1\. Forms のライセンスについて**  
Forms のライセンスは、現在 Office 365 Education に含まれております。  
現在は Forms はプレビュー版として Education 契約の Office 365 に先行してリリースしておりますが、  
今後、数カ月以内にプレビュー対象を拡大し、その他のお客様にもご利用を開始していただく予定です。  
ライセンスの確認方法については、例えば Office 365 E5 for Faculty の場合は、下記の様に Office 365 管理センターのユーザー管理のページにてライセンスの詳細項目を表示することで確認できます。  
\[ユーザー\] - \[アクティブなユーザー\] - "任意のユーザーを選択" - \[製品ライセンス\] の \[編集\] を選択し、起動するダイアログ上の当該テナント名をクリックしてプルダウンします。  
  
手動でライセンスの変更を行う場合は、上記にて Microsoft Forms をオン、もしくは、オフにすることで、該当ユーザーの Forms の有効化、無効化を行うことができます。  
しかしながら、テナント全体でオン、オフにする、といったような設定方法はないため、下記 PowerShell を使用して一括で実行する方法をご案内いたします。  
  
**2\. PowerShell を使用して一括で有効化する方法**  
PowerShell を利用するには、以下のページで PowerShell をインストールし、事前設定を行う必要があります。  
タイトル : PowerShell の導入について  
アドレス : [https://social.msdn.microsoft.com/Forums/ja-JP/d3cf96bd-36e6-4cae-b3d1-24e210b9263d](https://social.msdn.microsoft.com/Forums/ja-JP/d3cf96bd-36e6-4cae-b3d1-24e210b9263d)  
Power Shell を導入後、下記の有効化のスクリプト内容をコピーし、ps1 ファイルとして保存します。(例 : EnableFormsLicense.ps1)  

```
connect-msolservice -credential $msolcred

# ライセンスが割り当てられた全ユーザーの列挙
$users = Get-MsolUser -All | where {$_.isLicensed -eq "True"}

# ユーザー毎のライセンス/サービスの確認と変更
foreach ($user in $users)
{
    write-host ("Processing " + $user.UserPrincipalName)
    $licensetype = $user | Select-Object -ExpandProperty Licenses | Sort-Object { $_.Licenses }

    foreach ($license in $user.Licenses) 
    {
        write-host (" " + $license.accountskuid)

        $includeFORMS = $false
        $disableplan = @()

    	foreach ($row in $($license.ServiceStatus))
    	{
            # FORMS_PLAN_ が含まれるライセンスを取得
            if ( $row.ServicePlan.ServiceName -like "*FORMS_PLAN_*" ) 
            {
                $includeFORMS=$true
                $officeFORMS=$row.ServicePlan.ServiceName
            }
            
            if ( $row.ProvisioningStatus -eq "Disabled" -and $row.ServicePlan.ServiceName -notlike "*FORMS_PLAN_*")
            {
                $disableplan += $row.ServicePlan.ServiceName
            }
	}

        # Forms が含まれる SKU の場合、現在の設定に加え、Forms を有効化 (Enable)
        if ($includeFORMS)
        {
            write-host ("      found FORMS service in " + $license.accountskuid)
            write-host ("      current disabled services : " + $disableplan )
            $x = New-MsolLicenseOptions -AccountSkuId $license.accountskuid -DisabledPlans $disableplan
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $x
        }
    }
}
```

PowerShell を起動し、EnableFormsLicense.ps1 ファイルを保存したフォルダーに移動し、下記の様に実行します。(下記は c:\\temp 以下に ps1 ファイルを保存した場合の例となります。)  

```
cd c:\temp
.\EnableFormsLicense.ps1
```

認証ダイアログが表示されますので、ご利用中のテナントの管理者アカウント (adminuser@tenant.onmicrosoft.com 等) およびパスワード情報を入力します。  
処理中の内容が表示されますので、プロンプトが戻るまで待ちます。  
処理完了後、ブラウザーにて Office 365 を表示していた場合は一度ブラウザーを終了させ、再度 Office 365 にサインインします。  
ユーザーのライセンス付与状況を確認し、Forms 製品のライセンスが期待通りとなっていることを確認します。  
また、アプリケーション ランチャー上の Forms のタイルの表示が期待通りとなっていることを確認します。  
  
**3\. PowerShell を使用して一括で無効化する方法**  
PowerShell を利用するには、以下のページで PowerShell をインストールし、事前設定を行う必要があります。  
タイトル : PowerShell の導入について  
アドレス : https://blogs.technet.microsoft.com/officesupportjp/2016/06/30/powershell\_installation/  
Power Shell を導入後、下記の無効化のスクリプト内容をコピーし、ps1 ファイルとして保存します。(例 : DisableFormsLicense.ps1)  

```
connect-msolservice -credential $msolcred

# ライセンスが割り当てられた全ユーザーの列挙
$users = Get-MsolUser -All | where {$\_.isLicensed -eq "True"}

# ユーザー毎のライセンス/サービスの確認と変更
foreach ($user in $users)
{
    write-host ("Processing " + $user.UserPrincipalName)
    $licensetype = $user | Select-Object -ExpandProperty Licenses | Sort-Object { $_.Licenses }

    foreach ($license in $user.Licenses) 
    {
        write-host (" " + $license.accountskuid)

        $includeFORMS = $false
        $disableplan = @()

    	foreach ($row in $($license.ServiceStatus))
    	{
            # FORMS_PLAN_ が含まれるライセンスを取得
            if ( $row.ServicePlan.ServiceName -like "*FORMS_PLAN_*" ) 
            {
                $includeFORMS=$true
                $officeFORMS=$row.ServicePlan.ServiceName
            }
            
            if ( $row.ProvisioningStatus -eq "Disabled" -and $row.ServicePlan.ServiceName -notlike "*FORMS_PLAN_*")
            {
                $disableplan += $row.ServicePlan.ServiceName
            }
	}

        # Forms が含まれる SKU の場合、現在の設定に加え、Forms を無効化 (Disable)
        if ($includeFORMS)
        {
            $disableplan += $officeFORMS
            write-host ("      found FORMS service in " + $license.accountskuid)
            write-host ("      disabled services : " + $disableplan )
            $x = New-MsolLicenseOptions -AccountSkuId $license.accountskuid -DisabledPlans $disableplan
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $x
        }
    }
}
```

PowerShell を起動し、DisableFormsLicense.ps1 ファイルを保存したフォルダーに移動し、下記の様に実行します。(下記は c:\\temp 以下に ps1 ファイルを保存した場合の例となります。)  

```
cd c:\temp
.\DisableFormsLicense.ps1
```

認証ダイアログが表示されますので、ご利用中のテナントの管理者アカウント (adminuser@tenant.onmicrosoft.com 等) およびパスワード情報を入力します。  
処理中の内容が表示されますので、プロンプトが戻るまで待ちます。  
処理完了後、ブラウザーにて Office 365 を表示していた場合は一度ブラウザーを終了させ、再度 Office 365 にサインインします。  
ユーザーのライセンス付与状況を確認し、Forms 製品のライセンスが期待通りとなっていることを確認します。  
また、アプリケーション ランチャー上の Forms のタイルの表示が期待通りとなっていることを確認します。  
  
**4\. 関連情報**   
タイトル : Enable and disable Microsoft Forms  
アドレス : [https://support.office.com/en-us/article/8dcbf3ab-f2d6-459a-b8be-8d9892132a43](https://support.office.com/en-us/article/8dcbf3ab-f2d6-459a-b8be-8d9892132a43)  

タイトル : Microsoft フォームについてよく寄せられる質問  
アドレス : [https://support.office.com/ja-jp/article/495c4242-6102-40a0-add8-df05ed6af61c](https://support.office.com/ja-jp/article/495c4242-6102-40a0-add8-df05ed6af61c)  

タイトル : Office 365 ライセンスと Windows PowerShell: ユーザーのライセンスの状態を確認する  
アドレス : [https://technet.microsoft.com/ja-jp/library/dn771772.aspx](https://technet.microsoft.com/ja-jp/library/dn771772.aspx)  

タイトル : ユーザーへのライセンスの割り当て  
アドレス : [https://technet.microsoft.com/ja-jp/library/dn771770.aspx](https://technet.microsoft.com/ja-jp/library/dn771770.aspx)  

タイトル : サービスのライセンス情報を表示する  
アドレス : [https://technet.microsoft.com/ja-jp/library/dn771771.aspx](https://technet.microsoft.com/ja-jp/library/dn771771.aspx)  

タイトル : サービスにアクセスを割り当てる  
アドレス : [https://technet.microsoft.com/ja-jp/library/dn771769.aspx](https://technet.microsoft.com/ja-jp/library/dn771769.aspx)  

今回の投稿は以上です。