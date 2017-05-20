var X = XLSX;
var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
var wtf_mode = false;
var resultJson = null;

function to_json(workbook) {
  var result = {};
  workbook.SheetNames.forEach(function(sheetName) {
    var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName]);
    if (roa.length > 0) {
      result[sheetName] = roa;
    }
  });
  return result;
}


function process_wb(wb) {
  var output = "";
  resultJson = to_json(wb);
  //output = JSON.stringify(resultJson, 2, 2);
  console.log(resultJson);
  loadChart(resultJson);
  if (typeof console !== 'undefined') console.log("output", new Date());
}

var testExcel = document.getElementById('testExcel');

function handleDrop(e) {
  e.stopPropagation();
  e.preventDefault();
  rABS = true;
  use_worker = false;
  var files = e.dataTransfer.files;
  var f = files[0]; {
    var reader = new FileReader();
    //var name = f.name;
    reader.onload = function(e) {
      if (typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
      var data = e.target.result;
      wb = X.read(data, {
          type: 'binary'
      });
      process_wb(wb);
    };
    if (rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
  }
}

function handleDragover(e) {
  e.stopPropagation();
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
}

if (testExcel.addEventListener) {
  testExcel.addEventListener('dragenter', handleDragover, false);
  testExcel.addEventListener('dragover', handleDragover, false);
  testExcel.addEventListener('drop', handleDrop, false);
}


var xlf = document.getElementById('xlf');

function handleFile(e) {
  rABS = true;
  use_worker = false;
  var files = e.target.files;
  var f = files[0]; {
    var reader = new FileReader();
    //var name = f.name;
    reader.onload = function(e) {
      if (typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
      var data = e.target.result;
      var wb;
      if (rABS) {
          wb = X.read(data, {
              type: 'binary'
          });
      }
      process_wb(wb);
    };
    if (rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
  }
}

if (xlf.addEventListener) xlf.addEventListener('change', handleFile, false);

var myChart = echarts.init($('#testExcel')[0]);


function loadChart(data){
   var option = {
    title:{
      text: '专业分类'
    },
    tooltip:{
      trigger: 'axis',
      axisPointer: {
        type: 'shadow'
      }
    },
    legend: {
      data:[]
    },
    grid: {
        left: '3%',
        right: '4%',
        bottom: '3%',
        containLabel: true
    },
    xAxis : [
      {
        type : 'category',
        data : []
      }
    ],
    yAxis : [
        {
            type : 'value'
        }
    ],
    series:[
    ]
  };

  $.each(data.Sheet1,function(idx,val){
    var vals = [];
    var name = '';
    $.each(val,function(i,v){
      if(idx==0&&i!='类别'){
        option.xAxis[0].data.push(i);
        vals.push(v);
      }else if(i!='类别'){
        vals.push(v);
      }else if(i=='类别'){
        name = v;
        option.legend.data.push(v);
      }
    });
    option.series.push({
      name:name,
      type:'bar',
      data:vals
    })
  });
  console.log(option);
  myChart.setOption(option);
}

