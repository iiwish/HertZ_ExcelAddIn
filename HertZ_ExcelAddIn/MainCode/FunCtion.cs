using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace HertZ_ExcelAddIn
{
    public class FunCtion
    {
        Excel.Application ExcelApp;

        public bool SheetExist(string SheetName)
        {
            bool returnValue = false;
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            try
            {
                String Cell1 = ExcelApp.Worksheets[SheetName].Cells[1,1].Value;
                returnValue = true;
            }
            catch (Exception)
            {
                returnValue = false;
            }
            return returnValue;
        }

        public bool SelectSheet(string SheetName)
        {
            bool returnValue = false;
            if (SheetExist("余额表") == false)
            {
                string msg = "未发现“往来款明细”表，是否将当前工作表重命名为“往来款明细”并继续？";
                if ((int)MessageBox.Show(msg, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == 1)
                {
                    ExcelApp.ActiveSheet.Name = "往来款明细";
                    returnValue = true;
                    return returnValue;
                }
                else
                {
                    return returnValue;
                }
            }
            else
            {
                ExcelApp.ActiveWorkbook.Worksheets[SheetName].Select();
                returnValue = true;
                return returnValue;
            }
        }
    }       
}
