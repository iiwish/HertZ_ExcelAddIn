using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace HertZ_ExcelAddIn
{
    class AsyncCode
    {
        public async Task SheetToValueAsync()
        {
            FunCtion FunC = new FunCtion();
            await Task.Run(() => {
                Excel.Application ExcelApp = Globals.ThisAddIn.Application;
                Excel.Workbook WBK = ExcelApp.ActiveWorkbook;
                Excel.Worksheet wst;
                Excel.Range Rg;
                int AllRows;
                int AllColumns;

                double OneStep = Math.Round(1 / double.Parse(WBK.Worksheets.Count.ToString()) * 100, 6);//WBK.Worksheets.Count;
                double i = 0;

                //遍历工作表
                var SheetE = WBK.Worksheets.GetEnumerator();
                while (SheetE.MoveNext())
                {
                    wst = (Excel.Worksheet)SheetE.Current;
                    AllRows = Math.Max(FunC.AllRows(wst, "A", 13), 2);
                    AllColumns = Math.Max(FunC.AllColumns(wst, 1, 13), 2);
                    Rg = wst.Range[String.Format("A1:{0}{1}", FunC.CName(AllColumns), AllRows)];
                    Rg.Value2 = Rg.Value2;
                    ExcelApp.StatusBar = "已完成 " + Math.Round(i += OneStep, 2) + "%";
                }

                ExcelApp.StatusBar = false;
            });
        }

    }
}
