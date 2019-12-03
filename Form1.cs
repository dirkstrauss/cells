using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace asposeCells
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CreateDocument();
        }

        private void btnCreateDocument_Click(object sender, EventArgs e)
        {
            CreateDocument();
        }

        private void CreateDocument()
        {
            var wbook = new Workbook();

            var pth = "C:/temp/aspose/";

            #region Formulas Sheet 1
            //SetupFormulaWorkbookData(wbook.Worksheets[0]);

            //var ws = wbook.Worksheets[0];

            //ws.Name = "Formulas";            

            //wbook.Save($"{pth}1-Formulas.xlsx", SaveFormat.Xlsx);
            #endregion

            // CHARTS START

            //#region BoxWhisker
            //SetupBoxWhiskerChart(wbook.Worksheets[0]);
            //var ws = wbook.Worksheets[0];
            //ws.Name = "Charts";
            //wbook.Save($"{pth}2-BoxWhisker.xlsx", SaveFormat.Xlsx);
            //#endregion

            //#region Funnel
            ////SetupFunnelChart(wbook.Worksheets[0]);
            //SetupChart(wbook.Worksheets[0], ChartType.Funnel);
            //var ws = wbook.Worksheets[0];
            //ws.Name = "Charts";
            //wbook.Save($"{pth}3-Funnel.xlsx", SaveFormat.Xlsx);
            //#endregion

            //#region Pareto
            ////SetupParetoLineChart(wbook.Worksheets[0]);
            //SetupChart(wbook.Worksheets[0], ChartType.ParetoLine);
            //var ws = wbook.Worksheets[0];
            //ws.Name = "Charts";
            //wbook.Save($"{pth}4-Pareto.xlsx", SaveFormat.Xlsx);
            //#endregion

            //#region Sunburst
            ////SetupSunburstChart(wbook.Worksheets[0]);
            //SetupChart(wbook.Worksheets[0], ChartType.Sunburst);
            //var ws = wbook.Worksheets[0];
            //ws.Name = "Charts";
            //wbook.Save($"{pth}5-Sunburst.xlsx", SaveFormat.Xlsx);
            //#endregion

            //#region Treemap
            ////SetupTreeMapChart(wbook.Worksheets[0]);
            //SetupChart(wbook.Worksheets[0], ChartType.Treemap);
            //var ws = wbook.Worksheets[0];
            //ws.Name = "Charts";
            //wbook.Save($"{pth}6-Treemap.xlsx", SaveFormat.Xlsx);
            //#endregion

            //#region Waterfall
            ////SetupWaterfallChart(wbook.Worksheets[0]);
            //SetupChart(wbook.Worksheets[0], ChartType.Waterfall);
            //var ws = wbook.Worksheets[0];
            //ws.Name = "Charts";
            //wbook.Save($"{pth}7-Waterfall.xlsx", SaveFormat.Xlsx);
            //#endregion

            //#region Map
            SetupMapChart(wbook.Worksheets[0]);
            var ws = wbook.Worksheets[0];
            ws.Name = "Charts";
            wbook.Save($"{pth}8-Map.xlsx", SaveFormat.Xlsx);
            //#endregion


            // CHARTS END





        }


        private void SetupChart(Worksheet ws, ChartType chrtType)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(chrtType, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B3:B18";
        }

        private void SetupFormulaWorkbookData(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Name");
            ws.Cells["C2"].PutValue("Semester");
            ws.Cells["D2"].PutValue("Score");
            ws.Cells["E2"].PutValue("Grade");

            #region Add Names
            ws.Cells["B3"].PutValue("John");
            ws.Cells["B4"].PutValue("Lidia");
            ws.Cells["B5"].PutValue("Mark");
            ws.Cells["B6"].PutValue("Anne");
            ws.Cells["B7"].PutValue("Hayley");
            ws.Cells["B8"].PutValue("Lane");
            ws.Cells["B9"].PutValue("Peter");
            ws.Cells["B10"].PutValue("James");
            ws.Cells["B11"].PutValue("Mary");
            ws.Cells["B12"].PutValue("John");
            ws.Cells["B13"].PutValue("Lidia");
            ws.Cells["B14"].PutValue("Mark");
            ws.Cells["B15"].PutValue("Anne");
            ws.Cells["B16"].PutValue("Hayley");
            ws.Cells["B17"].PutValue("Lane");
            ws.Cells["B18"].PutValue("Peter");
            ws.Cells["B19"].PutValue("James");
            ws.Cells["B20"].PutValue("Mary");
            #endregion

            #region Add Semesters
            ws.Cells["C3"].PutValue(1);
            ws.Cells["C4"].PutValue(2);
            ws.Cells["C5"].PutValue(1);
            ws.Cells["C6"].PutValue(2);
            ws.Cells["C7"].PutValue(1);
            ws.Cells["C8"].PutValue(1);
            ws.Cells["C9"].PutValue(2);
            ws.Cells["C10"].PutValue(1);
            ws.Cells["C11"].PutValue(2);
            ws.Cells["C12"].PutValue(2);
            ws.Cells["C13"].PutValue(1);
            ws.Cells["C14"].PutValue(2);
            ws.Cells["C15"].PutValue(1);
            ws.Cells["C16"].PutValue(2);
            ws.Cells["C17"].PutValue(2);
            ws.Cells["C18"].PutValue(1);
            ws.Cells["C19"].PutValue(2);
            ws.Cells["C20"].PutValue(1);
            #endregion

            #region Add Scores
            ws.Cells["D3"].PutValue(75);
            ws.Cells["D4"].PutValue(65);
            ws.Cells["D5"].PutValue(15);
            ws.Cells["D6"].PutValue(75);
            ws.Cells["D7"].PutValue(95);
            ws.Cells["D8"].PutValue(56);
            ws.Cells["D9"].PutValue(72);
            ws.Cells["D10"].PutValue(88);
            ws.Cells["D11"].PutValue(24);
            ws.Cells["D12"].PutValue(61);
            ws.Cells["D13"].PutValue(72);
            ws.Cells["D14"].PutValue(97);
            ws.Cells["D15"].PutValue(17);
            ws.Cells["D16"].PutValue(63);
            ws.Cells["D17"].PutValue(84);
            ws.Cells["D18"].PutValue(48);
            ws.Cells["D19"].PutValue(65);
            ws.Cells["D20"].PutValue(68);
            #endregion

            // IFS Function
            // eg: =IFS(D3<60;"Fail";D3<70;"C";D3<80;"B";D3<90;"A";D3>=90;"A+")
            for (var i = 3; i <=20; i++)
            {
                ws.Cells[$"E{i}"].Formula = $"=IFS(D{i}<60,\"Fail\",D{i}<70,\"C\",D{i}<80,\"B\",D{i}<90,\"A\",D{i}>=90,\"A +\")";
            }

            // SWITCH Function
            // eg: =SWITCH(E3;"Fail";"Try harder";"C";"Ok";"B";"Good";"A";"Great";"A+";"Excellent")
            for (var i = 3; i <= 20; i++)
            {
                ws.Cells[$"F{i}"].Formula = $"=SWITCH(E{i},\"Fail\",\"Try harder\",\"C\",\"Ok\",\"B\",\"Good\",\"A\",\"Great\",\"A +\",\"Excellent\")";
            }

            // CONCAT Function
            // eg: =CONCAT(B3;" - Result: "; E3; " - "; F3)
            for (var i = 3; i <= 20; i++)
            {
                ws.Cells[$"G{i}"].Formula = $"=CONCAT(B{i},\" - Your result: \", E{i}, \" - \", F{i})";
            }


            #region Results
            ws.Cells["J2"].PutValue("Semester");
            ws.Cells["K2"].PutValue("Highest");
            ws.Cells["L2"].PutValue("Lowest");

            ws.Cells["J3"].PutValue(1);
            ws.Cells["J4"].PutValue(2);
            #endregion

            var maxFirstSemesterCell = ws.Cells["K3"];
            maxFirstSemesterCell.Formula = "=MAXIFS(D3:D20,C3:C20,\"1\")";

            var maxSecondSemesterCell = ws.Cells["K4"];
            maxSecondSemesterCell.Formula = "=MAXIFS(D3:D20,C3:C20,\"2\")";

            var minFirstSemesterCell = ws.Cells["L3"];
            minFirstSemesterCell.Formula = "=MINIFS(D3:D20,C3:C20,\"1\")";

            var minSecondSemesterCell = ws.Cells["L4"];
            minSecondSemesterCell.Formula = "=MINIFS(D3:D20,C3:C20,\"2\")";

            #region Name Details
            ws.Cells["B23"].PutValue("First name");
            ws.Cells["C23"].PutValue("Middle name");
            ws.Cells["D23"].PutValue("Last name");
            ws.Cells["E23"].PutValue("Full name");

            ws.Cells["B24"].PutValue("John");
            ws.Cells["B25"].PutValue("Lidia");
            ws.Cells["B26"].PutValue("Mark");
            ws.Cells["B27"].PutValue("Anne");
            ws.Cells["B28"].PutValue("Hayley");
            ws.Cells["B29"].PutValue("Lane");
            ws.Cells["B30"].PutValue("Peter");
            ws.Cells["B31"].PutValue("James");
            ws.Cells["B32"].PutValue("Mary");

            ws.Cells["C24"].PutValue("Reginald");
            ws.Cells["C27"].PutValue("Mary");
            ws.Cells["C28"].PutValue("Lindy");
            ws.Cells["C30"].PutValue("Lee");

            ws.Cells["D24"].PutValue("Van Zandt");
            ws.Cells["D25"].PutValue("Cunningham");
            ws.Cells["D26"].PutValue("Lester");
            ws.Cells["D27"].PutValue("Joseph");
            ws.Cells["D28"].PutValue("Miller");
            ws.Cells["D29"].PutValue("Bower");
            ws.Cells["D30"].PutValue("Sanders");
            ws.Cells["D31"].PutValue("Williams");
            ws.Cells["D32"].PutValue("Davis");
            #endregion

            // TEXTJOIN Function
            // eg: =TEXTJOIN(" "; TRUE; B24:D24)
            for (var i = 24; i <= 32; i++)
            {
                ws.Cells[$"E{i}"].Formula = $"=TEXTJOIN(\" \", TRUE, B{i}:D{i})";
            }           
        }

        #region ...
        private void SetupBoxWhiskerChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(ChartType.BoxWhisker, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B3:B18";
        }

        private void SetupFunnelChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(ChartType.Funnel, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B3:B18";
        }

        private void SetupParetoLineChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(ChartType.ParetoLine, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B3:B18";

            //chart.SetChartDataRange("B2:E18", true);
        }

        private void SetupSunburstChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(ChartType.Sunburst, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B3:B18";
        }

        private void SetupTreeMapChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(ChartType.Treemap, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            //var srs = chart.NSeries[0];
            //srs.DataLabels.ShowCategoryName = true;

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B2:B18";


        }

        private void SetupWaterfallChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Produce");
            ws.Cells["C2"].PutValue("Year 2014");
            ws.Cells["D2"].PutValue("Year 2015");
            ws.Cells["E2"].PutValue("Year 2016");

            for (var i = 3; i <= 18; i++)
            {
                if (i == 3 || i == 7 || i == 11 || i == 15)
                    ws.Cells[$"B{i}"].PutValue("Oranges");

                if (i == 4 || i == 8 || i == 12 || i == 16)
                    ws.Cells[$"B{i}"].PutValue("Apples");

                if (i == 5 || i == 9 || i == 13 || i == 17)
                    ws.Cells[$"B{i}"].PutValue("Pears");

                if (i == 6 || i == 10 || i == 14 || i == 18)
                    ws.Cells[$"B{i}"].PutValue("Grapes");
            }

            var rnd = new Random();

            for (var i = 3; i <= 18; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"D{i}"].PutValue(rnd.Next(10000, 70000));
                ws.Cells[$"E{i}"].PutValue(rnd.Next(10000, 70000));
            }

            var chartIndex = ws.Charts.Add(ChartType.Waterfall, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C18", true);
            _ = chart.NSeries.Add("=D3:D18", true);
            _ = chart.NSeries.Add("=E3:E18", true);

            chart.NSeries.CategoryData = "=B3:B18";
        } 
        #endregion

        private void SetupMapChart(Worksheet ws)
        {
            ws.Cells["B2"].PutValue("Country");
            ws.Cells["C2"].PutValue("Sales");

            ws.Cells[$"B3"].PutValue("South Africa");
            ws.Cells[$"B4"].PutValue("Canada");
            ws.Cells[$"B5"].PutValue("India");
            ws.Cells[$"B6"].PutValue("France");
                        
            var rnd = new Random();

            for (var i = 3; i <= 6; i++)
            {
                ws.Cells[$"C{i}"].PutValue(rnd.Next(50000, 70000));                
            }

            var chartIndex = ws.Charts.Add(ChartType.Map, 6, 6, 25, 15);
            var chart = ws.Charts[chartIndex];

            _ = chart.NSeries.Add("=C3:C6", true);

            chart.NSeries.CategoryData = "=B3:B6";
        }

    }
}
