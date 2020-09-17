'use strict';

var TurndownService = require('turndown');
var DomParser = require('dom-parser');
var request = require('request');
var fs = require('fs');
const { stdout, stderr } = require('process');

if (require.main === module) {
  main();
}

function main() {
  var turndownService = new TurndownService();
  var url = getUrl();

  // Get old page
  request.get(url, function(err, res, html) {
    if (err) {
      console.log('Error: ' + err.message);
      return;
    }

    var obj = getBody(html);

    // 10進数の数値文字参照を通常テキストへデコードするために、本来とは違う用途で turndown を使用
    obj.title = turndownService.turndown(obj.title);

    var markdown = getYamlInfo(obj) + turndownService.turndown(obj.body);
    var imageUrls = getImageUrls(markdown);
    if (imageUrls.length > 0) {
      getImages(imageUrls, obj.title);
      markdown = replaceImageUrls(imageUrls, markdown);
    }

    var outputFileName = obj.title + '.md';
    writeFile('source/_posts/', outputFileName, markdown);
  });

  // Generate new page
  require('child_process').exec('hexo generate', (err, stdout) => {
    if (err) {
      console.log('err:', err);
    }
    else {
      console.log(stdout);
    }
  });
}

function replaceImageUrls(urls, body) {
  for (var i = 0; i < urls.length; i++) {
    body = body.replace(urls[i], 'image' + (i+1) + '.png');
  }
  return body;
}

function getImages(urls, title) {
  var folderName = 'source/_posts/' + title;
  if (!fs.existsSync(folderName)) {
    fs.mkdirSync(folderName)
  }

  var imageNum = 1;
  for (var i = 0; i < urls.length; i++) {
    var url = urls[i];
    request(
      {method: 'GET', url: url, encoding: null},
      function (error, response, body){
        if(!error && response.statusCode === 200){
          fs.writeFileSync(folderName + '/image' + (imageNum++) + '.png', body, 'binary');
        }
      }
    );
  }
}

function getImageUrls(body) {
  var urls = [];
  var urlStartIndex = body.indexOf('https://social.msdn.microsoft.com/Forums/getfile/');

  while (urlStartIndex != -1) {
    var urlEndIndex = body.indexOf(')', urlStartIndex);
    var url = body.substring(urlStartIndex, urlEndIndex);
    urls.push(url);
    urlStartIndex = body.indexOf('https://social.msdn.microsoft.com/Forums/getfile/', urlEndIndex);
  }

  return urls;
}

function getUrl() {
  if (process.argv.length == 3) {
    return process.argv[2];
  }
  else {
    console.error("Argument error: Please input 1 argument (target page url).");
    process.exit(1);
  }
}

function getYamlInfo(obj) {
  var yamlText = '---\n';
  yamlText += 'title: ' + obj.title + '\n';
  yamlText += 'date: ' + obj.date + '\n';
  yamlText += '---\n\n';
  return yamlText;
}

function writeFile(path, fileName, data) {
  fs.writeFile(path + fileName, data, function (err) {
    if (err) {
        throw err;
    }
    else {
      console.log('[Generated] ' + fileName);
    }
  });
}

function getBody(htmlStrings) {
  var domParser = new DomParser();
  var obj = new Object();
  var doc = domParser.parseFromString(htmlStrings, 'text/html');
  obj.title = doc.getElementsByTagName('title')[0].innerHTML.replace('/', '-').replace(' : ', '：');
  obj.body = doc.getElementsByClassName('body')[0].innerHTML;
  obj.date = formatDate(doc.getElementsByClassName('date')[0].textContent);
  return(obj);
}

// yaml に記述可能な形式に変換 ("2020年8月24日 15:29" -> "2020-08-24")
function formatDate(date) {
  date = date.split(' ')[0]
  date = date.replace('年', '-').replace('月', '-').replace('日', '');
  var dates = date.split('-');
  if (dates[1].length == 1) {
    dates[1] = '0' + dates[1];
  }
  if (dates[2].length == 1) {
    dates[2] = '0' + dates[2];
  }
  return dates.join('-');
}
