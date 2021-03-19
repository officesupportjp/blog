---
title: template # ファイル名と同じ
date: 2020-09-15 # 日付
tags: # オプション（複数指定可能）
- タグ
- Excel
---

これはサンプルファイルです。
<!-- コメントアウトはこのように書きます。 -->

# 見出し1
## 見出し2
### 見出し3
#### 見出し4
##### 見出し5

<!-- 
画像を指定する際、実際には "template/" などのフォルダのパスは必要ありません。
[ブログのタイトル] フォルダを作成し、その中に "image.png" を配置して、Markdown には "image.png" のみを記述します。
 -->
![](template/image.png)

```PowerShell
npm run getpage コマンドやソースコードの記述
```

`print('短いコードなど');`

[リンク（GitHubへ）](https://github.com/officesupportjp/blog)

- 箇条書き
  - Tabキーで階層化

1. 段落番号
2. 段落番号

**太字**

*斜体*

~~取り消し線~~

> 引用
> 改行したい場合は \
> バックスラッシュ
>> 2重も可能

---
**上記で生成されるブログページ**
![](template/visible.png)
