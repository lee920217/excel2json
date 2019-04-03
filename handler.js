const { ipcRenderer } = require("electron");
const xlsx = require("xlsx");

var fs = require("fs"),
  drag = $("#drag");

drag.on("dragenter dragover", function(event) {
  // 重写ondragover 和 ondragenter 使其可放置
  event.preventDefault();

  drag.addClass("drag-ondrag");
  drag.text("Release Mouse");
});

drag.on("dragleave", function(event) {
  event.preventDefault();

  drag.removeClass("drag-ondrag");
  drag.text("Please Drag sth. in here");
});

drag.on("drop", function(event) {
  // 调用 preventDefault() 来避免浏览器对数据的默认处理（drop 事件的默认行为是以链接形式打开）
  event.preventDefault();

  // 原生语句如下，但引进jquery后要更改
  // var file=event.dataTransfer.files[0];
  // 原因：
  // 在jquery中，最终传入事件处理程序的 event 其实已经被 jQuery 做过标准化处理，
  // 其原有的事件对象则被保存于 event 对象的 originalEvent 属性之中，
  // 每个 event 都是 jQuery.Event 的实例
  // 应该这样写:
  var efile = event.originalEvent.dataTransfer.files[0];
  drag.removeClass("drag-ondrag");
  const workbook = xlsx.readFile(efile.path);
  const sheetNames = workbook.SheetNames;
  const worksheet = workbook.Sheets[sheetNames[0]];
  const header = {};
  const data = [];
  const keys = Object.keys(worksheet);
  const keysList = {};
  let tmpKey;
  keys
    .filter(k => k[0] !== "!")
    .forEach((k, v) => {
      //console.log(k);
      let col = k.substr(0, 1);
      let row = parseInt(k.substr(1));
      let value = worksheet[k].v;
      if (row == 1 && col != "A") {
        header[value] = {};
        keysList[col] = (value);
      }
      if (row != 1) {
        const innerKey = Object.keys(header);
        if (col == "A") {
          //获取key
          tmpKey = value;
        } else {
          let headerKey = keysList[col];
          header[headerKey][tmpKey] = value;
        } 
      }
    });
  const textarea = $('textarea');
  textarea.val(JSON.stringify(header, null, '\t'));
  //console.log('keyList => ', keysList);
  // fs.readFile(efile.path,"utf8",function(err,data){
  //     if(err) throw err;
  //     console.log(data);

  //     // drag.text(data);
  //     //读取文件的内容，如果使用jquery的text()或html()，val()都不能保留格式
  //     //而使用Obj.innerText 则可以。那么把drag的jquery对象转为dom对象
  //     drag[0].innerText=data;
  // });
  console.log();
  return false;
});
