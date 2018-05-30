'use strict';

//import * as functions from  '/Functions/FunctionFile.js';





(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
                console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
            }

            formatControls();
            ModelViz();


            document.getElementById("help").onclick = function () {
               location.href = "/help.html";
            };

            

            $("#run").click(run);
            $("#createTable").click(createTable);
            $("#createGraph").click(createGraph);

        });
    };

    function formatControls() {

        var TextFieldElements = document.querySelectorAll(".ms-TextField");
        for (var i = 0; i < TextFieldElements.length; i++) {
            new fabric['TextField'](TextFieldElements[i]);
        }

        var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
        for (var i = 0; i < DropdownHTMLElements.length; ++i) {
            var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
        }

        var PivotElements = document.querySelectorAll(".ms-Pivot");
        for (var i = 0; i < PivotElements.length; i++) {
            new fabric['Pivot'](PivotElements[i]);
        }

       
    }

    function run() {
        Excel.run(function (context) {

            var range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            var s = Number($('#StockPrice').val());
            var x = Number($('#StrikePrice').val());
            var t = Number($('#YearsTM').val());
            var r = Number($('#RiskFreeRate').val());
            var v = Number($('#Vol').val());
            var flag = $('#flag').val();


            var optionVal = Number(BlackScholes(flag, s, x, t, r, v));
            optionVal = Number(optionVal.toFixed(2));
            var test = addtwoValues(s, x);
            console.log(test);
            console.log(optionVal);

            range.values = [[optionVal]];

            range.load("address");
            return context.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    function createTable() {
        Excel.run(function (context) {


            const sheets = context.workbook.worksheets;

            const sheet = sheets.add();
            sheet.name = "Data";
            sheets.load("items/name");
            context.sync();



            const currentWorksheet = context.workbook.worksheets.getItem("Data");
            const OptionDataTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
            OptionDataTable.name = "OptionsDataTable";

            OptionDataTable.getHeaderRowRange().values =
                [["Date", "Stock Price", "Strike Price", "Option Value"]];

            OptionDataTable.rows.add(null /*add at the end*/, [
                ["1/1/2017", "61", "65", "120"],
                ["1/1/2017", "61", "65", "120"],
                ["1/1/2017", "61", "65", "120"],
                ["1/1/2017", "61", "65", "120"],
                ["1/1/2017", "61", "65", "120"],
                ["1/1/2017", "61", "65", "120"]
            ]);

            OptionDataTable.columns.getItemAt(1).getRange().numberFormat = [['$#,##0.00']];
            OptionDataTable.columns.getItemAt(2).getRange().numberFormat = [['$#,##0.00']];
            OptionDataTable.columns.getItemAt(3).getRange().numberFormat = [['$#,##0.00']];
            OptionDataTable.getRange().format.autofitColumns();
            OptionDataTable.getRange().format.autofitRows();

            return context.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

     function createGraph() {
         Excel.run(function (context) {


            const sheet = context.workbook.worksheets.getItem("Data");
            var dataRange = sheet.tables.getItem("OptionsDataTable");
            var loadedvalues = dataRange.getDataBodyRange();

            context.sync();


            var chart = sheet.charts.add("Line", loadedvalues, "auto");

            chart.title.text = "Options Data";
            chart.legend.position = "right"
            chart.legend.format.fill.setSolidColor("white");
            chart.dataLabels.format.font.size = 15;
            chart.dataLabels.format.font.color = "black";
            chart.series.getItemAt(0).name = "Stock Price";
            
            return context.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }


     function addtwoValues(a, b) {

         return a + b;
     }





   

    function BlackScholes(PutCallFlag, S, X, T, r, v) {

        var d1, d2;
        d1 = (Math.log(S / X) + (r + v * v / 2.0) * T) / (v * Math.sqrt(T));
        d2 = d1 - v * Math.sqrt(T);

        if (PutCallFlag == "Call")
            return S * CND(d1) - X * Math.exp(-r * T) * CND(d2);
        else
            return X * Math.exp(-r * T) * CND(-d2) - S * CND(-d1);
    }

    function CND(x) {

        var a1, a2, a3, a4, a5, k;

        a1 = 0.31938153, a2 = -0.356563782, a3 = 1.781477937, a4 = -1.821255978, a5 = 1.330274429;

        if (x < 0.0)
            return 1 - CND(-x);
        else
            k = 1.0 / (1.0 + 0.2316419 * x);
        return 1.0 - Math.exp(-x * x / 2.0) / Math.sqrt(2 * Math.PI) * k
            * (a1 + k * (-0.356563782 + k * (1.781477937 + k * (-1.821255978 + k * 1.330274429))));

    }

    function Gdelta(CallPutFlag,S,X,T,r,b,v) {

        
        var dl; 
        dl = (Log(S / X) + (b + v ^ 2 / 2) * T) / (v * Sqr(T));

        if (CallPutFlag == "Call")
        return  Exp((b-r) * T) * CND(d1);
        else
        return  -Exp((b-r) * T) * CND(-d1);
        


    }

    function ModelViz() {

        // Create and populate a data table.
        var data = new vis.DataSet();
        // create some nice looking data with sin/cos
        var counter = 0;
        var steps = 50;  // number of datapoints will be steps*steps
        var axisMax = 314;
        var axisStep = axisMax / steps;
        for (var x = 0; x < axisMax; x += axisStep) {
            for (var y = 0; y < axisMax; y += axisStep) {

                //var value = (Math.sin(x / 50) * Math.cos(y / 50) * 50 + 50);
                var value = (Gdelta("Call",Number(50),Number(55),x,Number(0.035),Number(0.10),y));
                data.add({ id: counter++, x: x, y: y, z: value, style: value });
            }
        }

        // specify options
        var options = {
            width: '100%',
            height: '400px',
            style: 'surface',
            showPerspective: true,
            showGrid: true,
            showShadow: false,
            keepAspectRatio: true,
            verticalRatio: 0.5,
            tooltip: true
        };

        // Instantiate our graph object.
        var container = document.getElementById('Model');
        var graph3d = new vis.Graph3d(container, data, options);

    }


})();