---
title: (WebAPI)Microsoft Graph - OneDrive API (C Sharp) を使ったサンプル コード
date: 2019-03-01
---

(※ 2016 年 12 月 29 日に Japan Office Developer Support Blog に公開した情報のアーカイブです。)

こんにちは、Office Developer サポートの森　健吾 (kenmori) です。

今回の投稿では、Microsoft Graph – OneDrive API を実際に C# で開発するエクスペリエンスをご紹介します。

ウォークスルーのような形式にしておりますので、慣れていない方も今回の投稿を一通り実施することで、プログラム開発を経験し理解できると思います。本投稿では、現実的な実装シナリオを重視するよりも、OneDrive API を理解するためになるべくシンプルなコードにすることを心掛けています。

例外処理なども含めていませんので、実際にコーディングする際には、あくまでこのコードを参考する形でご検討ください。

**事前準備**

[以前の投稿](https://officesupportjp.github.io/blog/(WebAPI)OAuth%20Bearer%20Token%20(Access%20Token)%20%E3%81%AE%E5%8F%96%E5%BE%97%E6%96%B9%E6%B3%95%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6/)をもとに、Azure AD にアプリケーションの登録を完了してください。少なくとも以下の 2 つのデリゲートされたアクセス許可が必要です。

・Have full access to all files user can access  
・Sign users in

その上で、クライアント ID とリダイレクト URI を控えておいてください。

**開発手順**

1\. Visual Studio を起動し、Windows フォーム アプリケーションを開始します。  
2\. ソリューション エクスプローラにて \[参照\] を右クリックし、\[NuGet パッケージの管理\] をクリックします。

![](image1.png)

3\. ADAL で検索し、Microsoft.IdentityMode.Clients.ActiveDirectory をインストールします。

![](image2.png)

4\. \[OK\] をクリックし、\[同意する\] をクリックします。  
5\. 次に同様の操作で Newtonsoft で検索し、Newtonsoft.Json をインストールします。  
6\. 次にフォームをデザインします。

![](image3.png)

**コントロール一覧**

*   OneDriveTestForm フォーム
*   fileListLB リスト ボックス
*   fileNameTB テキスト ボックス
*   uploadBtn ボタン
*   renameBtn ボタン
*   deleteBtn ボタン
*   openFileDialog1 オープン ファイル ダイアログ

7\. プロジェクトを右クリックし、\[追加\] – \[新しい項目\] をクリックします。  
8\. MyFile.cs を追加します。  
9\. 以下のような定義 (JSON 変換用) を記載します。  
 ※ 今回のサンプルでは使用しない定義も残しています。コードを書き換えてデータを参照したり、変更するなどしてみてください。

```
using Newtonsoft.Json;
using System.Collections.Generic;

namespace OneDriveDemo
{
    public class MyFile
    {
        public string name { get; set; }
        // 以下のプロパティは今回使用しませんが、デバッグ時に値を見ることをお勧めします。
        public string webUrl { get; set; }
        public string createdDateTime { get; set; }
        public string lastModifiedDateTime { get; set; }
    }

    public class MyFiles
    {
        public List<MyFile> value;
    }

    // ファイル移動時に使います。
    public class MyParentFolder
    {
        public string path { get; set; }
    }

    public class MyFileModify
    {
        public string name { get; set; }
        // ファイル移動時に使います。
        public MyParentFolder parentReference { get; set; }
    }
}
```

  

**補足**
JSON 変換時にオブジェクト プロパティ側に別名を付けたい場合は、下記のように JsonProperty 属性を指定することで可能です。

```
[JsonProperty("name")]
public string FileName { get; set; }
```
  

10\. フォームのコードに移動します。

11\. using を追記しておきます。

```
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OneDriveDemo; //9. で作成した MyFile.cs 内の名前空間
```

  

12\. フォームのメンバー変数に以下を加えます。  
 ※ clientid や redirecturi は Azure AD で事前に登録したものを使用ください。

```
const string resource = "https://graph.microsoft.com";
const string clientid = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
const string redirecturi = "urn:getaccesstokenfordebug";
// ADFS 環境で SSO ドメイン以外のテナントのユーザーを試す場合はコメント解除
// const string loginname = "admin@tenant.onmicrosoft.com";

string AccessToken;
```
  

13\. フォームのデザインでフォームをダブルクリックし、ロード時のイベントを実装します。

```
private async void Form1_Load(object sender, EventArgs e)
{
    AccessToken = await GetAccessToken(resource, clientid, redirecturi);
    DisplayFiles();
}

// アクセス トークン取得
private async Task<string> GetAccessToken(string resource, string clientid, string redirecturi)
{
    AuthenticationContext authenticationContext = new AuthenticationContext("https://login.microsoftonline.com/common");
    AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenAsync(
        resource,
        clientid,
        new Uri(redirecturi),
        new PlatformParameters(PromptBehavior.Auto, null)
        // ADFS 環境で SSO ドメイン以外のテナントのユーザーを試す場合はコメント解除
        //, new UserIdentifier(loginname, UserIdentifierType.RequiredDisplayableId)
        );
    return authenticationResult.AccessToken;
}

// ファイル一覧表示
private async void DisplayFiles()
{
    using (HttpClient httpClient = new HttpClient())
    {
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);
        HttpRequestMessage request = new HttpRequestMessage(
            HttpMethod.Get,
            new Uri("https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name,weburl,createdDateTime,lastModifiedDateTime")
        );
        var response = await httpClient.SendAsync(request);
        MyFiles files = JsonConvert.DeserializeObject<MyFiles>(response.Content.ReadAsStringAsync().Result);

        fileListLB.Items.Clear();
        foreach (MyFile file in files.value)
        {
            fileListLB.Items.Add(file.name);
        }
    }
    if (!string.IsNullOrEmpty(fileNameTB.Text))
    {
        fileListLB.SelectedItem = fileNameTB.Text;
    }
}
```

  

14\. fileListLB の SelectedIndexChanged イベントをダブルクリックして、処理を実装します。

```
// リスト ボックスで選択したファイルをテキスト ボックスに同期
private void fileListLB_SelectedIndexChanged(object sender, EventArgs e)
{
    fileNameTB.Text = ((ListBox)sender).SelectedItem.ToString();
}
```

  

15\. uploadBtn をダブルクリックして、ボタンクリックイベントを実装します。  
※ この方式でアップロードできるファイル サイズには制限があります。

```
// ファイル アップロード処理
private async void uploadBtn_Click(object sender, EventArgs e)
{
    if (openFileDialog1.ShowDialog() == DialogResult.OK)
    {
        fileNameTB.Text = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf("\\") + 1);
        using (HttpClient httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "octet-stream");
            HttpRequestMessage request = new HttpRequestMessage(
                HttpMethod.Put,
                new Uri(string.Format("https://graph.microsoft.com/v1.0/me/drive/root:/{0}:/content", fileNameTB.Text))
            );
            request.Content = new ByteArrayContent(ReadFileContent(openFileDialog1.FileName));
            var response = await httpClient.SendAsync(request);
            MessageBox.Show(response.StatusCode.ToString());
        }
        DisplayFiles();
    }
}

// ローカル ファイルの読み取り処理
private byte[] ReadFileContent(string filePath)
{
    using (FileStream inStrm = new FileStream(filePath, FileMode.Open))
    {
        byte[] buf = new byte[2048];
        using (MemoryStream memoryStream = new MemoryStream())
        {
            int readBytes = inStrm.Read(buf, 0, buf.Length);
            while (readBytes > 0)
            {
                memoryStream.Write(buf, 0, readBytes);
                readBytes = inStrm.Read(buf, 0, buf.Length);
            }
            return memoryStream.ToArray();
        }
    }
}
```

  

16\. renameBtn をダブルクリックして、クリックイベントを実装します。

```
// ファイル名の変更処理
private async void renameBtn_Click(object sender, EventArgs e)
{
    foreach (string fileLeafRef in fileListLB.SelectedItems)
    {
        using (HttpClient httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");

            HttpRequestMessage request = new HttpRequestMessage(
                new HttpMethod("PATCH"),
                new Uri(string.Format("https://graph.microsoft.com/v1.0/me/drive/root:/{0}", fileLeafRef))
            );

            MyFileModify filemod = new MyFileModify();
            filemod.name = fileNameTB.Text;
            request.Content = new StringContent(JsonConvert.SerializeObject(filemod), Encoding.UTF8, "application/json");

            var response = await httpClient.SendAsync(request);
            MessageBox.Show(response.StatusCode.ToString());
        }
    }
    DisplayFiles();
}
```
  

17\. deleteBtn をダブルクリックして、クリックイベントを実装します。

```
// ファイル削除処理
private async void deleteBtn_Click(object sender, EventArgs e)
{
    foreach (string fileLeafRef in fileListLB.SelectedItems)
    {
        using (HttpClient httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AccessToken);

            HttpRequestMessage request = new HttpRequestMessage(
                HttpMethod.Delete,
                new Uri(string.Format("https://graph.microsoft.com/v1.0/me/drive/root:/{0}", fileLeafRef))
            );
            var response = await httpClient.SendAsync(request);
            MessageBox.Show(response.StatusCode.ToString());
        }
    }
    fileNameTB.Text = "";
    DisplayFiles();
}
```

18\. ソリューションをビルドして、動作を確認しましょう。

 最初にログイン画面が表示され、ログインしたユーザーでアクセス トークンを取得します。  
※ ディレクトリ同期されたドメインでダイアログを表示すると、ユーザー名・パスワードの入力画面が表示されず、自動的にログインされる場合があります。

![](image4.png)

該当ユーザーの OneDrive for Business のルート ディレクトリのファイル一覧が表示されます。

![](image5.png)

該当フォルダーで、ファイルのアップロードや名前変更、削除などの操作をお試しください。

**参考情報**

以下は、Microsoft Graph における OneDrive API の参考サイトです。

タイトル : Microsoft Graph でのファイルの作業  
アドレス : [https://graph.microsoft.io/ja-jp/docs/api-reference/v1.0/resources/onedrive](https://graph.microsoft.io/ja-jp/docs/api-reference/v1.0/resources/onedrive)

OneDrive API のリファレンスについては、以下のページをご確認いただき、OneDrive APIが持つ様々なメソッドをお試しください。

タイトル : Develop with the OneDrive API  
アドレス : [https://dev.onedrive.com/README.htm#](https://dev.onedrive.com/README.htm)

現時点では、OneDrive API で実装できる範囲について、OneDrive コンシューマ版と OneDrive for Business には相違があります。詳細は以下をご確認ください。

タイトル : Release notes for using OneDrive API with OneDrive for Business and SharePoint  
アドレス : [https://dev.onedrive.com/sharepoint/release-notes.htm](https://dev.onedrive.com/sharepoint/release-notes.htm)

以下は、OneDrive API を紹介した Channel9 の動画になります。

タイトル : Office Dev Show - Episode 38 - OneDrive APIs in the Microsoft Graph  
アドレス : [https://channel9.msdn.com/Shows/Office-Dev-Show/Office-Dev-Show-Episode-38-OneDrive-APIs-in-the-Microsoft-Graph](https://channel9.msdn.com/Shows/Office-Dev-Show/Office-Dev-Show-Episode-38-OneDrive-APIs-in-the-Microsoft-Graph)

Json.NET に関するドキュメントは以下をご参考にしてください。

タイトル : Json.NET Documentation  
アドレス : [http://www.newtonsoft.com/json/help/html/Introduction.htm](http://www.newtonsoft.com/json/help/html/Introduction.htm)

タイトル : Serializing and Deserializing JSON  
アドレス : [http://www.newtonsoft.com/json/help/html/SerializingJSON.htm](http://www.newtonsoft.com/json/help/html/SerializingJSON.htm)

別投稿に記載した通り、開発工数削減のため、アプリケーション開発前に Graph Explorer, Fiddler や Postman などを使用して、あらかじめ使用する REST を確立しておくことをお勧めします。デバッグ方法については、以下をご参考にしてください。

タイトル : \[WebAPI\]Microsoft Graph を使用した開発に便利なツール群  
アドレス : [https://officesupportjp.github.io/blog/(WebAPI)Microsoft Graph を使用した開発に便利なツール群/](https://officesupportjp.github.io/blog/(WebAPI)Microsoft%20Graph%20%E3%82%92%E4%BD%BF%E7%94%A8%E3%81%97%E3%81%9F%E9%96%8B%E7%99%BA%E3%81%AB%E4%BE%BF%E5%88%A9%E3%81%AA%E3%83%84%E3%83%BC%E3%83%AB%E7%BE%A4/)

今回の投稿は以上です。 

   
**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**