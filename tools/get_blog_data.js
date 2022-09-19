const fs = require('fs');
const path = require('path');
const yaml = require("js-yaml");
const csv = require("csv");
const {parse} = require('csv-parse/sync');

// このファイルのディレクトリ
const thisFileDir = path.dirname(process.argv[1]);

// 記事があるディレクトリ source/_posts
const postsDir = path.join(thisFileDir, '..', 'source','_posts');

// csv のディレクトリ
const csvDir = path.join(thisFileDir,'out.csv');



function csv_export(){

    //変換後の配列を格納
    let newData = [];
    newData.push(['title', 'id', 'tags', 'filename', 'user']);
    const allNames = fs.readdirSync(postsDir);
    const files = allNames.filter(file => /.*\.md$/.test(file));  
    for(var i in files){

        var data = fs.readFileSync(path.join(postsDir, files[i]), 'utf8');

        // 操作
        console.log(path.join(postsDir, files[i]));

        // json の取得
        var header = data.match(/^---[\s\S]+?---/) ?? [''];
        var tags_json = yaml.load(header[0].replace(/^-+|-+$/g, ''), {
            schema: yaml.JSON_SCHEMA
        });

        var row = [];
        row.push(tags_json['title'] ?? '');
        row.push(tags_json['id'] ?? '');
        row.push((tags_json['tags'] ?? ['']).toString());
        row.push(files[i]);
        row.push((data.match(/こんにちは.+?です/) ?? [''])[0].replace(/こんにちは、*|です/g, ''))
        newData.push(row);
    }
    console.log(newData);
    //write
    csv.stringify(newData,(error,output)=>{
        fs.writeFileSync( path.join(thisFileDir,'out.csv'), output, 'utf8');
    })
}

function csv_import(){
    
    const csvData = read_csv();
    const allNames = fs.readdirSync(postsDir);
    const files = allNames.filter(file => /.*\.md$/.test(file));
    for(var i in files){
        var data = fs.readFileSync(path.join(postsDir, files[i]), 'utf8');

        // json の取得
        var header = data.match(/^---[\s\S]+?---/) ?? [''];
        var tags_json = yaml.load(header[0].replace(/^-+|-+$/g, ''), {
            schema: yaml.JSON_SCHEMA
        });

        var target = csvData.find((v) => v['id'] == tags_json['id']);
        if(target !== null && target !== undefined){
            console.log(target);
            delete tags_json['alias'];
            tags_json['tags'] = target['tags']

            data = data.replace(/^---[\s\S]+?---/, '---\n' + yaml.dump(tags_json) + '\n---');

            fs.writeFileSync(path.join(postsDir, files[i]), data, 'utf8');    
        }
    }
}

function get_alias(){

    let aliasData = [];
    const allNames = fs.readdirSync(postsDir);
    const files = allNames.filter(file => /.*\.md$/.test(file));  
    for(var i in files){

        var data = fs.readFileSync(path.join(postsDir, files[i]), 'utf8');

        // json の取得
        var header = data.match(/^---[\s\S]+?---/) ?? [''];
        var tags_json = yaml.load(header[0].replace(/^-+|-+$/g, ''), {
            schema: yaml.JSON_SCHEMA
        });
        
        if(tags_json['alias'] !== null){
            var title = tags_json['title'];
            var id = tags_json['id'];
            aliasData.push(title+'/index.html: '+id+'/index.html');
        }
    }
    fs.writeFileSync( path.join(thisFileDir,'alias_config.txt'), aliasData.join('\n'), 'utf8');
}

function read_csv(){
    // csvファイルの内容を読み込み
    const fileData = fs.readFileSync(csvDir);
    // csvファイルをパース
    const records = parse(fileData, { columns: true });
    // tags 要素の整形
    for (var j in records) {
        records[j]['tags'] =  records[j]['tags'].split(',')
    }
    return records;
}

csv_import();

//get_alias();