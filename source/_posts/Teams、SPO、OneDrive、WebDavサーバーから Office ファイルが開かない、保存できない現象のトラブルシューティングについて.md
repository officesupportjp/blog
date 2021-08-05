---
title: Teams、SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象のトラブルシューティングについて
date: 2021-06-21
---

こんにちは、Office サポートの西川 (直) です。  
  
今回の投稿では、Teams、SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象のトラブルシューティングについて説明します。

以下、問題の切り分けに有用な情報を以下にご案内いたします。

これらの手順を上から順にご実施ください。

**<実施方法>**

1) Office への再サインイン

\--------------------------------

1\. Office からサインアウトし、Office を全て終了します。

2\. OS のスタートメニュー - 歯車 - "アカウント" - "職場または学校にアクセスする" をクリックし、"職場または学校アカウント" として表示されている項目がある場合、"切断" をクリックします

※ Azure AD や AD に接続済みの項目を "切断" しますと、ドメインから外れてしまいますので、Azure AD や AD に接続済み の項目は切断しないようにご注意ください。

3\. OS からログオフします。(可能であれば、再起動してください)

4\. Office を起動してサインインし、事象が解消するかをお試しください。

　  
2) Office アップロードセンターのキャッシュクリア

\---------------------------------------------------

1\. Office を全て終了します。

2\. \[スタート\]-\[すべてのプログラム\]-\[Microsoft Officeツール\]-\[Office アップロードセンター\] をクリックし、\[設定\] を押下後、\[キャッシュ ファイルの削除\] を実施してください。

3\. \[ファイルを閉じたときに Office ドキュメント キャッシュから削除する\] オプションでチェックを入れ有効に設定します。

4\. \[OK\] ボタンをクリックし、Microsoft Office アップロード センターの設定 ダイアログを閉じます。  

上記の手順でアップロードセンターが見つからない場合、以下の手順をご実施ください。

1\. Excel 等を開きます。

2\. \[ファイル\] タブ - "オプション" を開きます。

3\. "保存" より、"キャッシュ" の設定 - \[ファイルを閉じたときに Office ドキュメント キャッシュから削除する\] オプションでチェックを入れ有効に設定し、\[キャッシュ ファイルの削除\] を実施します。


　  
3) Office アップロードセンターの完全キャッシュクリア

\---------------------------------------------------

1\. Ctrl + Shift + Esc キーを押してタスク マネージャーを起動します。

2\. \[プロセス\] タブで \[MSOSYNC.EXE\] が存在する場合、クリックし、\[プロセスの終了\] をクリックしてから、タスク マネージャーを終了します。

3\. %userprofile%\\AppData\\Local\\Microsoft\\Office\\16.0\\ に移動します。

※ Office 2013 の場合は 15.0 となります。

4\. OfficeFileCache フォルダー配下のファイルを削除できるもののみ、削除して下さい。

5\. スタートメニューを右クリックして \[ファイル名を指定して実行\] から regedit と入力し、レジストリ エディタを起動します。

6\. 以下のキーを展開します。

HKEY\_CURRENT\_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Internet\\Server Cache

7\. Server Cache キーを選択し、削除します。

8\. OS を再起動します。  


　  
4) Office のサインインキャッシュのクリア

\---------------------------------------------------

以下の記事の手順により、サインインキャッシュのクリアをご実施ください。

Title : Office のサインインのトラブルシュートについて

URL : [https://social.msdn.microsoft.com/Forums/ja-JP/528e54cc-6e92-420d-b7ce-0e5a65fb6d3f/office?forum=officesupportteamja](https://social.msdn.microsoft.com/Forums/ja-JP/528e54cc-6e92-420d-b7ce-0e5a65fb6d3f/office?forum=officesupportteamja)

**\- 関連情報**

Title : MSI 版の Office 2016 で SPO、OneDrive、WebDavサーバーから Office ファイルが開かない、保存できない現象について

URL : [https://social.msdn.microsoft.com/Forums/ja-JP/cb8cb890-b30a-467b-a4fa-d7f7159dd373/msi-2925612398-office-2016-12391?forum=officesupportteamja](https://social.msdn.microsoft.com/Forums/ja-JP/cb8cb890-b30a-467b-a4fa-d7f7159dd373/msi-2925612398-office-2016-12391?forum=officesupportteamja)

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**