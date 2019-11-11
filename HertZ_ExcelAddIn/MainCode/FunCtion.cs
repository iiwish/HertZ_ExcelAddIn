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
        Excel.Application ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        Excel.Worksheet WST;

        /// <summary>
        /// 数字转列字母
        /// </summary>
        private string CName(int ColumnNumber)
        {
            //if (ColumnNumber < 1) { throw new Exception("invalid parameter"); }

            int dividend = ColumnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// 列名转换数字
        /// </summary>
        private int CNumber(string ColumnName)
        {
            //if (!System.Text.RegularExpressions.Regex.IsMatch(ColumnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

            int index = 0;
            char[] chars = ColumnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index;
        }

        /// <summary>
        /// 判断工作表是否存在
        /// </summary>
        public bool SheetExist(string SheetName)
        {
            bool returnValue = false;

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

        /// <summary>
        /// 将目标工作表重命名，并选中
        /// </summary>
        public bool SelectSheet(string SheetName)
        {
            bool returnValue = false;

            if (SheetExist(SheetName) == false)
            {
                string msg = "未发现“" + SheetName + "”表，是否将当前工作表重命名为“" + SheetName + "”并继续？";

                if ((int)MessageBox.Show(msg, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == 1)
                {
                    ExcelApp.ActiveSheet.Name = SheetName;
                    returnValue = true;
                }
            }
            else
            {
                ExcelApp.ActiveWorkbook.Worksheets[SheetName].Select();
                returnValue = true;
            }

            return returnValue;
        }

        /// <summary>
        /// 返回指定列的行数
        /// </summary>
        public int AllRows(string ColumnName = "A",int ColumnsTotal = 1)
        {
            
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            int returnValue = 0;
            int StartColumn = CNumber(ColumnName);
            int NewRows;
            String Column;

            for (int i = StartColumn; i < StartColumn + ColumnsTotal; i++)
            {
                Column = CName(i);
                NewRows = ((Excel.Range)(WST.Cells[WST.Rows.Count, Column])).End[Excel.XlDirection.xlUp].Row;
                returnValue = Math.Max(returnValue, NewRows);
            }

            return returnValue;
        }

        /// <summary>
        /// 返回指定行的列数
        /// </summary>
        public int AllColumns(int RowName = 1,int RowsTotal = 1)
        {
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            int returnValue = 0;
            int NewColumns;

            for (int i = RowName; i < RowName + RowsTotal; i++)
            {
                NewColumns = ((Excel.Range)(WST.Cells[i, "IV"])).End[Excel.XlDirection.xlToLeft].Column;
                returnValue = Math.Max(returnValue, NewColumns);
            }

            return returnValue;
        }

        public bool RangeIsStandard()
        {
            bool returnValue = false;

            if (AllRows("A",10) == AllRows("A"))
            {
                if (AllColumns(1, 10) == AllColumns(1))
                {
                    returnValue = true;
                }
            }

            return returnValue;
        }
    } 
}
