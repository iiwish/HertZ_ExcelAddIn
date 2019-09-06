using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace HertZ_ExcelAddIn
{
    public class B_TableProcessing
    {
        //protected string SubjectCoding { get; set; } //配置文件名（要包含后缀名）

        Excel.Application ExcelApp;
        Worksheet WST;

        //public int AllRows(String ColumnIndex)
        //{

        //    ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        //    WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
        //    return ((Range)(WST.Cells[WST.Rows.Count, "A"])).End[Excel.XlDirection.xlUp].Row;

        //}

        public void testbox()
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Worksheet)ExcelApp.ActiveSheet;
            MessageBox.Show(((Range)(WST.Cells[WST.Rows.Count, "A"])).End[XlDirection.xlUp].Row.ToString());
        }

    }
}
