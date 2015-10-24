var xlsx = require('xlsx');
var builder = require('xmlbuilder');
var config = require('./config.json');
var workbook = xlsx.readFile(config.source);
var sheet1 = workbook.SheetNames[0];
var workSheet = workbook.Sheets[sheet1];
var fs = require('fs');

var transformToJsonList = function(workSheet){
    var termColumn = 'B',definitionColumn = 'C', imageUrlColumn = 'D', audioColumn = 'E';
    var rowIndex = 2;
    var list = [];

    while(workSheet[termColumn+rowIndex]){
        list.push(
            {
                value: workSheet[termColumn+rowIndex].v,
                description: workSheet[definitionColumn+rowIndex].v,
                audioUrl: workSheet[audioColumn+rowIndex].v,
                imageUrl: workSheet[imageUrlColumn+rowIndex]?workSheet[imageUrlColumn+rowIndex]:''
            }
        );
        rowIndex++;
    }
    return list;
};
var groupAlphabetOrder = function(jsonList){
    var mainObject = {};
    jsonList.forEach(function(data){
        var firstLetter = data.value && data.value.charAt(0).toUpperCase();
        if(!mainObject[firstLetter]){
            mainObject[firstLetter] = {
                words:[]
            };
        }
        mainObject[firstLetter].words.push(data);
    });
    return mainObject;
};
var createRenderList = function(groupResult){
    var list=[];
    for (var key in groupResult) {
        if (groupResult.hasOwnProperty(key)) {
            list.push({
                value:key,
                words:groupResult[key].words
            });
        }
    }

    //Sort List to Ascending Order
    list.sort(function(a,b){
        if(a.value < b.value) return -1;
        if(a.value > b.value) return 1;
        return 0;
    });
    return list;
};
var buildXml = function(renderList){
    var glossary = builder.create('glossary');
    renderList.forEach(function(alphabet){
        var words = glossary.ele('alphabet').ele('value').dat(alphabet.value).insertAfter('words');
        alphabet.words.forEach(function(word){
            var tmpWord = words.ele('word')
                    .ele('value',{'type': 'String'}).dat(word.value);
            if(word.description){
                tmpWord.insertAfter('description',{'type': 'String'}).dat(word.description);
            }
            if(word.audioUrl){
                tmpWord.insertAfter('audioURL',{'type': 'String'}).dat(word.audioUrl);
            }
            if(word.imageUrl){
                tmpWord.insertAfter('imageURL',{'type': 'String'}).dat(word.imageUrl);
            }
        });
    });
    return glossary.doc().end({ pretty: true});
};

var result = transformToJsonList(workSheet);
var groupResult = groupAlphabetOrder(result);
var renderList = createRenderList(groupResult);
var xmlData = buildXml(renderList);

//Write File
fs.writeFile(config.destination,xmlData,function(err){
    if(err){
        throw err;
    }else{
        console.log('XML File Generated Successfully');
    }
});



