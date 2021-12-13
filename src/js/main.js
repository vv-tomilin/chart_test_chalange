
import '../css/style.css';

var XLSX = require('xlsx');
var Highcharts = require('highcharts');

var inputFile = document.getElementById('select-file');

inputFile.addEventListener('change', handleFile);

//* функция парсинга EXCEL файла
function handleFile(e) {
  var files = e.target.files;
  var i, f;

  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();

    reader.onload = function (e) {
      var data = e.target.result;

      var workbook = XLSX.read(data, { type: 'binary' });

      //* таблица с нужными даннными
      var sheets = Object.entries(workbook.Sheets['Сетевой график']);

      //* данные из таблицы "Сетевой график"
      var bottomholeDepth = []; //* глубина забоя
      var dates = [];           //* даты
      var operationsNames = []; //* названия операций

      var yPointsData = []; //* данные для точек графика

      var subtractionTime = 0; //* первая дата для расчета суток на оси X

      var regExp = /[B,H,J]1[9-9]|[B,H,J][2-9][0-9][0-9]+|[B,H,J][2-9][0-9]+|[B,H,J]1[0-9][0-9]+|[B,H,J][2-9][0-9][0-9]+/;

      for (var j = 0; j < sheets.length; j++) {
        var current = sheets[j];

        if (regExp.test(current[0])) {
          switch (current[0].split('')[0]) {
            case 'B':
              //* записываю названия операций
              operationsNames.push(current[1].v);
              break;

            case 'H':
              //* записываю глубины забоя
              bottomholeDepth.push(current[1].v);
              break;

            case 'J':
              let date = current[1].w;

              //* конвертирую дату в формат ISO
              var dateDay = '20' + date.split(' ')[0].split('/').reverse().join('-');
              var dateTime = date.split(' ')[1] && (date.split(' ')[1].length) === 5 ? date.split(' ')[1] : '0' + date.split(' ')[1];
              var fullTime = dateDay + 'T' + dateTime;

              //* записываю дату и время начала для расчета суток
              if (j === 107) {
                subtractionTime = fullTime;
              }

              //* записываю дату для отображения во всплывающем окне графика 
              //* и рассчитываю количество прошедших суток на выбранную дату
              dates.push(
                [current[1].w.replaceAll('/', '.'),
                ((Date.parse(fullTime) - Date.parse(subtractionTime)) / 1000 / 60 / 60 / 24).toFixed(2)]
              );
              break;
          }
        }
      }

      //* формирую объект с данными для всплывающего окошка 
      //* с данными для каждой точки графика
      for (var k = 0; k < dates.length; k++) {
        yPointsData.push({
          y: bottomholeDepth[k],
          date: dates[k][0],
          bottomhole: bottomholeDepth[k],
          operation: operationsNames[k],
          duration: dates[k][1],
        });
      }

      //* рисую график на основе распарсенных данных из таблицы
      Highcharts.chart('container', {
        chart: {
          type: 'line'
        },
        title: {
          text: ''
        },
        xAxis: {
          title: {
            text: '<b>Продолжительность бурения (сут)</b>'
          },
          categories: dates.map(function (elem) {
            return elem[1];
          })
        },
        yAxis: {
          reversed: true,
          title: {
            text: 'Глубина по стволу (м)'
          }
        },
        tooltip: {
          formatter: function () {
            return (
              '<b style="font-size: 15px; color: green">План</b><br>'
              + '<b>Дата/время: </b>'
              + this.point.date + '<br>'
              + '<b>Забой: </b>' + this.point.bottomhole + '<br>'
              + '<b>Операция: </b>' + this.point.operation + '<br>'
              + '<b>Прошло суток: </b>' + this.point.duration)
          }
        },
        series: [{
          name: 'Сетевой график "Глубина-день"',
          data: yPointsData
        }]
      });
    };
    reader.readAsBinaryString(f);
  }
}





