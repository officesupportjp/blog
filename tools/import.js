const fs = require('fs');
const path = require('path');
const yaml = require("js-yaml");
const XLSX = require('xlsx');

// このjsファイルのディレクトリ
const thisFileDir = path.dirname(process.argv[1]);

// 記事があるディレクトリ source/_posts
const postsDir = path.join(thisFileDir, '..', 'source','_posts');

// xlsx のディレクトリ
const xlsxDir = path.join(thisFileDir,'out.xlsx');

// xlsx ファイルの読み込み
const workbook = XLSX.readFile(xlsxDir);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const jsonHeaders = XLSX.utils.sheet_to_json(sheet);

// postsDir 内のファイルを取得して ループ処理
const allNames = fs.readdirSync(postsDir);
const files = allNames.filter(file => /.*\.md$/.test(file));
for(var i in files){

    // ファイルの読み込み
    let data = fs.readFileSync(path.join(postsDir, files[i]), 'utf8');

    // json の取得
    let header = data.match(/^---[\s\S]+?---/) ?? [''];
    let tags_json = yaml.load(header[0].replace(/^-+|-+$/g, ''), {
        schema: yaml.JSON_SCHEMA
    });

    // id が同じもを読み込んだxlsxファイルから取得
    let target = jsonHeaders.find((v) => v['id'] == tags_json['id']);
    if(target !== null && target !== undefined){
        
        // alias の削除
        delete tags_json['alias'];

        // target の中のタグは文字列のため配列に変更
        const tagsArr = target['tags'].split(',');
        tagsArr.forEach((v,i,arr) => {
            arr[i] = v.replace(/^\s+|\s+$/g, '');//最初と最後のスペースを削除
        });

        // 代入
        tags_json['tags'] = tagsArr;

        // 記事へのheaderの変更
        data = data.replace(/^---[\s\S]+?---/, '---\n' + yaml.dump(tags_json) + '\n---');
        fs.writeFileSync(path.join(postsDir, files[i]), data, 'utf8');    
    }
}