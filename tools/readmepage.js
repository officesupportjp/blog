const fs = require('fs');
const path = require('path');
const yaml = require("js-yaml");

// このjsファイルのディレクトリ
const thisFileDir = path.dirname(process.argv[1]);

// 記事があるディレクトリ source/_posts
const postsDir = path.join(thisFileDir, '..', 'source','_posts');

// 書き込み先ファイル
const readme = path.join(thisFileDir, 'articles_README.md');

// postsDir 内のファイルを取得して ループ処理
const allNames = fs.readdirSync(postsDir);
const files = allNames.filter(file => /.*\.md$/.test(file));

let contents = '';
for(var i in files){

    // ファイルの読み込み
    let data = fs.readFileSync(path.join(postsDir, files[i]), 'utf8');

    // 記事の header を json 形式で取得
    let header = data.match(/^---[\s\S]+?---/) ?? [''];
    let tags_json = yaml.load(header[0].replace(/^-+|-+$/g, ''), {
        schema: yaml.JSON_SCHEMA
    });

    contents += '- [' + tags_json['title'] + '](./source/_posts/' + files[i].replace(/ |　/g, '%20') + ')\n';
}

fs.writeFileSync(readme, "# Articles\n" + contents);