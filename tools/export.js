const fs = require('fs');
const path = require('path');
const yaml = require("js-yaml");
const XLSX = require('xlsx')

// このjsファイルのディレクトリ
const thisFileDir = path.dirname(process.argv[1]);

// 記事があるディレクトリ source/_posts
const postsDir = path.join(thisFileDir, '..', 'source','_posts');

// xlsx のディレクトリ
const xlsxDir = path.join(thisFileDir,'out.xlsx');

// xlsx ファイルに差し込む 二次元配列
let newData = [['title', 'id', 'tags', 'filename', 'user', 'link']];

// postsDir 内のファイルを取得して ループ処理
const allNames = fs.readdirSync(postsDir);
const files = allNames.filter(file => /.*\.md$/.test(file));
for(let i in files){

    // ファイルの読み込み
    let data = fs.readFileSync(path.join(postsDir, files[i]), 'utf8');

    // 記事の header を json 形式で取得
    let header = data.match(/^---[\s\S]+?---/) ?? [''];
    let tags_json = yaml.load(header[0].replace(/^-+|-+$/g, ''), {
        schema: yaml.JSON_SCHEMA
    });

    // newData に push する一次元配列
    let row = [];
    row.push(tags_json['title'] ?? '');
    row.push(tags_json['id'] ?? '');
    row.push((tags_json['tags'] ?? ['']).toString());
    row.push(files[i]);
    row.push((data.match(/こんにちは.+?です/) ?? [''])[0].replace(/こんにちは、*|です/g, ''));
    row.push('https://officesupportjp.github.io/blog/' + tags_json['id'] + '/');
    newData.push(row);
}


// ファイルに書き込み
const workbook = XLSX.utils.book_new();
const sheet = XLSX.utils.aoa_to_sheet(newData);
XLSX.utils.book_append_sheet(workbook, sheet, "BlogInfo");
XLSX.writeFile(workbook, xlsxDir);
