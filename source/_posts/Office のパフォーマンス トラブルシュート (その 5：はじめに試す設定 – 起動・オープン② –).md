---
title: Office のパフォーマンス トラブルシュート (その 5：はじめに試す設定 – 起動・オープン② –)
date: 2020-02-26
lastupdate: 2020-05-13
---

こんにちは、Office サポート チームの中村です。

パフォーマンス改善効果が期待できる設定の第 3 弾です。今回は、前回に続き Office アプリケーションの起動やファイル オープンが遅い場合に効果が期待できる設定のうち、ネットワーク通信と関連する設定を紹介します。  
  

**目次**

1\. SMB 通信設定 (ファイル サーバー上の共有フォルダのファイルを開く場合)

2\. SharePoint からファイルを開くときのキャッシュ設定

3\. インターネット サービスに接続する機能

4\. 証明書有効性確認処理

補足: 適切なインターネット通信設定 (2. / 3. / 4. に関わる設定)  
  

**1\. SMB** **通信設定**

通常、ファイルサーバーの共有フォルダに格納された Office ファイルを開くときには、クライアントとサーバー間で SMB プロトコルで通信が行われます。SMB 2.1 を利用している場合、Lease 機能が利用できます。一般的にはこの機能によりクライアントにキャッシュが保持されることで SMB 通信によるネットワーク トラフィックを削減し、ファイル オープンのパフォーマンスは向上しますが、ウィルス対策ソフトの動作などの影響を受け、この機能が Office ファイルを開くパフォーマンスの低下を招く場合があります。このような場合は以下レジストリを設定して Lease 機能を無効化することが検討できます。(Lease 機能を無効化すると、Lease 機能より少し古い仕組みの Oplock 機能によるキャッシュが行われます。) また、現在無効な場合は有効にすると速くなる可能性があります。

参考) Client caching features: Oplock vs. Lease

[https://blogs.msdn.microsoft.com/openspecification/2009/05/22/client-caching-features-oplock-vs-lease/  
](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fblogs.msdn.microsoft.com%2Fopenspecification%2F2009%2F05%2F22%2Fclient-caching-features-oplock-vs-lease%2F&data=02%7C01%7Ckanakamu%40microsoft.com%7Cb92b5dc872414bd5e88e08d7a565dd88%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637159729243459967&sdata=eYDBezWpWECafLxmGvYpQmmsRuhUKzg6seCgGh71or0%3D&reserved=0)

<レジストリ\>

キー : HKEY\_LOCAL\_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\Lanmanserver\\Parameters

名前: DisableLeasing

種類: REG\_DWORD

値: (有効: 0 / 無効: 1)

\* 設定後、OS を再起動してください。  
  

**2\.** **SharePoint からファイルを開くときのキャッシュ設定**

SharePoint のドキュメント ライブラリから Office ファイルを開くときには、独自のキャッシュの仕組みなどを用いた通信が行われます。

プロキシを介している環境などにおいて、Office アプリケーションから SharePoint サーバーへの名前解決や HTTP 通信での CONNECT / OPTIONS / GET などの要求に対し、応答がすぐに返るようネットワーク通信周りが適切に構成されていない場合に、処理のタイムアウトを待つことで Office ファイルのオープンが遅延する場合があります。このような場合には、この処理を無効化することで起動パフォーマンスを改善できます。  

<グループポリシー\>

\[Microsoft Office 2016\] – \[Microsoft Office ドキュメント キャッシュ\] – \[最初に Office ドキュメント キャッシュからドキュメントを開く\]

設定値 : 無効  

<レジストリ\>

キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Internet

名前 : DisableServerReachability

種類 : REG\_DWORD

値 : 1

**3\.** **インターネット サービスに接続する機能**

Office には多数のオンライン サービスを利用する機能があります。これらの機能では、Office の起動時やファイル オープン時、リボンのタブ移動時など、各機能に適したタイミングでバックグラウンドでそのサービスのサイトへの通信が生じます。2\. のネットワークコスト確認処理と同様、ネットワーク構成によってはこの通信への応答が返るまでに時間を要し、パフォーマンスに影響を与えます。

オンライン サービスを利用する必要がない場合は、それらの機能を無効化する設定を行うことでパフォーマンスを改善できます。オンライン サービスの無効化を行う方法は、Office 365 バージョン 1904 でのプライバシー コントロールの新機能追加に伴い、この前後で大きく異なっています。  

<u>**Office 2016 / 2019 / Office 365** **バージョン 1903 以前**</u>

以下の過去投稿でオンライン サービスを無効化する方法をご案内しています。

Office 2016 が行うインターネット接続について

[https://social.msdn.microsoft.com/Forums/ja-JP/decfa528-6257-4b89-bca5-e0e613682865/office-2016](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fsocial.msdn.microsoft.com%2FForums%2Fja-JP%2Fdecfa528-6257-4b89-bca5-e0e613682865%2Foffice-2016&data=02%7C01%7Ckanakamu%40microsoft.com%7Cb92b5dc872414bd5e88e08d7a565dd88%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637159729243469961&sdata=pKo0W2r8dlkmw%2BPKXvK2BSsu0bYa95l5Cb5L8yIg2Ek%3D&reserved=0)

オプションについては上記記事で言及されていないため、以下に記載します。  

<オプション\>

いずれか 1 つのアプリケーションから設定を変更すると、他のアプリケーションの動作にも反映されます。

\[ファイル\] – \[オプション\] – \[セキュリティ センター\] – \[セキュリティ センターの設定\] – \[プライバシー オプション\] – \[Office を Microsoft のオンラインサービスに接続して、使用状況や環境設定に関連する機能を提供できるようにしますか?\] のチェックを外す

また、グループポリシーでは個別の機能ごとに無効化する設定が用意されているものもありますが、今回の記事では個々の機能の設定については割愛します。  

<u>**Office 365** **バージョン 1904 以降**</u>

バージョン 1904 でプライバシー コントロール機能が追加され、Office アプリケーションから送信する診断データ、および利用できるオンライン サービス (この変更以降、"接続エクスペリエンス" と呼ばれます) の整理と新たな制御方法が用意されました。詳細は以下の公開情報でご案内しています。

Office 365 ProPlus のプライバシー コントロールの概要

[https://docs.microsoft.com/ja-jp/DeployOffice/privacy/overview-privacy-controls](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2FDeployOffice%2Fprivacy%2Foverview-privacy-controls&data=02%7C01%7Ckanakamu%40microsoft.com%7Cb92b5dc872414bd5e88e08d7a565dd88%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637159729243479954&sdata=K03qU0PwWzaoM3QT3SEiANDyWFRrfOr7nCyYy1Kc4iE%3D&reserved=0)

ポリシーの設定を使用して、Office 365 ProPlus のプライバシー コントロールを管理する

[https://docs.microsoft.com/ja-jp/DeployOffice/privacy/manage-privacy-controls](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2FDeployOffice%2Fprivacy%2Fmanage-privacy-controls&data=02%7C01%7Ckanakamu%40microsoft.com%7Cb92b5dc872414bd5e88e08d7a565dd88%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637159729243479954&sdata=OKNN0OPC3H1qZ2Wr2HTkZutS2BZI4U2ejpB%2BPYz0wO4%3D&reserved=0)

Office 365 ProPlus の Office クラウドポリシーサービスの概要

[https://docs.microsoft.com/ja-jp/DeployOffice/overview-office-cloud-policy-service](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2FDeployOffice%2Foverview-office-cloud-policy-service&data=02%7C01%7Ckanakamu%40microsoft.com%7Cb92b5dc872414bd5e88e08d7a565dd88%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637159729243489949&sdata=A%2BXdhY5lbj%2B4XBATbWlKSCgyyFBVZbHMKJFwpvPnfWs%3D&reserved=0)

\* OS ログインユーザーに対する Active Directory グループポリシー制御の他に、Office 365 クラウドポリシーを用いて Office にログインしている Office 365 ユーザーに対する制御を行うこともできます。

状況によって細かな制御が必要になるため、上記資料を参照してご利用の環境、使い方に合った設定としていただければと思いますが、切り分けのため、最も厳しい設定としてすべてを無効化するには以下を設定します。  

<グループ ポリシー\>

\[Microsoft Office 2016\] – \[プライバシー\] – \[セキュリティ センター\] – \[Officeで接続エクスペリエンスの使用を許可する\]

設定値 : 無効  

<レジストリ\>

キー : HKEY\_CURRENT\_USER\\Software\\Policies\\Microsoft\\Office\\16.0\\Common\\Privacy

名前 : DisconnectedState

種類 : REG\_DWORD

値 : 2

**4\.** **証明書有効性確認処理**

Office アプリケーションに含まれるモジュールや、サードパーティ提供を含むアドインの多くでは、ファイルに対して証明書による署名が行われています。Office の起動時にはこれらのモジュール、アドインを読みこみますが、その際に署名が有効なものであるかを検証します。署名の検証プロセスでは、署名に使用されている証明書やそのルート証明書について有効性を確認する処理が行われます。このために、インターネットに証明書の情報を取得しに行きますが、2\. / 3. と同様、端末のネットワーク構成によってはこの処理でタイムアウト待ちが生じ、起動遅延に繋がります。

このような場合、この証明書の確認処理を行わないようにすることができます。本来これを無効化するとセキュリティ強度は低下しますが、タイムアウト待ちが生じている環境では元々確認は行えていないため、無効化しても状況に変わりはありません。

証明書の有効性確認では、「証明書失効確認処理 (CRL)」と「ルート証明書更新処理 (CTL)」があり、それぞれの無効化方法をご案内できるもののみ記載します。  

<設定\>

**証明書失効確認の無効化**

Internet Explorer の \[ツール\] – \[インターネット オプション\] – \[詳細設定\] - \[発行元証明書の取り消しを確認する\] のチェックをオフにする  

<グループポリシー\>

**ルート証明書更新プログラムの無効化**

\[コンピュータの構成\] – \[Windowsの設定\] - \[セキュリティの設定\] - \[公開キーのポリシー\] - \[証明書パス検証の設定\] - \[ネットワークの取得\] タブ - \[これらのポリシーの設定を定義する\] のチェックを有効にし、 \[Microsoft ルート証明書プログラムで証明書を自動更新する\] のチェックを外す  

<レジストリ\>

**ルート証明書更新プログラムの無効化**

キー： HKEY\_LOCAL\_MACHINE\\Software\\Policies\\Microsoft\\SystemCertificates\\AuthRoot

名前 : DisableRootAutoUpdate

種類 : REG\_DWORD

値： 1

**証明書失効確認の無効化**

キー : HKEY\_CURRENT\_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\WinTrust\\Trust Providers\\Software Publishing

名前 : State

種類 : REG\_DWORD

値 : 他の設定と合わせて 1 つのレジストリで管理するためユーザーの設定状況によって値は異なります。

IE が既定の設定に対して証明書失効確認の無効化のみを設定した場合は 0x23E00 (10進数: 146944) になります。

  

**補足: 適切なインターネット通信設定**

上記では、状況次第でタイムアウトにつながる機能の通信を抑止する方法を案内しましたが、機能を無効化しなくても、このような通信が行われたときに、即座に通信不可の応答が返れば問題ありません。このためには、一般的に以下の点にご注意ください。

**・名前解決が正常に行えるようにする、または解決できない旨の応答が速やかに返るようにする**

DNS サーバーのフォワード設定で不必要に転送が行われないようにしたり、Office の接続先アドレスを DNS サーバーの名前解決リストや、状況次第では hosts に追加します。これはダミーのアドレスでも構いません。

**・Office からの接続要求にすぐに応答を返す (正否は不問)**

名前解決はできた場合、その接続先への接続を行います。接続確立にタイムアウトが起きないよう、ネットワーク構成を確認します。

適切に応答が返るよう構成すべき Office の接続先情報は以下の公開情報を参照してください。Office 365 向けの情報ですが、Office 2016 / 2019 でもほとんど同じです。

Office 365 の URL と IP アドレスの範囲

[https://docs.microsoft.com/ja-jp/office365/enterprise/urls-and-ip-address-ranges](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fja-jp%2Foffice365%2Fenterprise%2Furls-and-ip-address-ranges&data=02%7C01%7Ckanakamu%40microsoft.com%7Cb92b5dc872414bd5e88e08d7a565dd88%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637159729243489949&sdata=0LhO6Mzjw3tJVmIXx84RQQVc%2FVtQjTit4zdRcjiwFhI%3D&reserved=0)

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**