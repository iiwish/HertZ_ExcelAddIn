using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Drawing;

namespace HertZ_ExcelAddIn
{
    public class FunCtion
    {
        private Excel.Application ExcelApp;
        Excel.Worksheet WST;
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void ClearBackExcel()
        {
            //int Rows;
            //for (int i = 1; i < 13; i++)
            //{
            //    ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            //    WST = (Excel.Worksheet)ExcelApp.ActiveSheet;

            //    try
            //    {
            //        Rows = ((Excel.Range)(WST.Cells[WST.Rows.Count, 1])).End[Excel.XlDirection.xlUp].Row;
            //    }
            //    catch
            //    {
            //        MessageBox.Show("后台有未清理的Excel程序，请检查并清理");
            //    }
            //}
        }

        /// <summary>
        /// 数字转列字母
        /// </summary>
        public string CName(int ColumnNumber)
        {
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
        public int CNumber(string ColumnName)
        {
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
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            bool returnValue;

            try
            {
                String Cell1 = ExcelApp.Worksheets[SheetName].Cells[1,1].Value.ToString();
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
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
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
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
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
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = ExcelApp.ActiveSheet;
            int returnValue = 0;
            int NewColumns;

            for (int i = RowName; i < RowName + RowsTotal; i++)
            {
                NewColumns = ((Excel.Range)(WST.Cells[i, "IV"])).End[Excel.XlDirection.xlToLeft].Column;
                returnValue = Math.Max(returnValue, NewColumns);
            }

            return returnValue;
        }

        /// <summary>
        /// 检查Sheet中的数据区域是否规范
        /// </summary>
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

        /// <summary>
        /// 判断字符串是否是字母
        /// </summary>
        public bool IsLetter(string str)
        {
            bool returnValue;
            if (System.Text.RegularExpressions.Regex.IsMatch(str, @"(?i)^[A-Za-z]+$"))
            {
                returnValue = true;
            }
            else
            {
                returnValue = false; 
            }
            return returnValue;
        }

        /// <summary>
        /// 判断字符串是否是数字
        /// </summary>
        public bool IsNumber(string str)
        {
            bool returnValue;
            try
            {
                double OutN = double.Parse(str);
                returnValue = true;
            }
            catch
            {
                returnValue = false;
            }
            return returnValue;
        }

        /// <summary>
        /// 选择指定列名
        /// </summary>
        public int SelectColumn(List<string> ColumnName,List<string> OName,bool MustSelect)
        {
            int returnValue = 0;
            
            //匹配现有列名和目标列名
            for (int i = 1; i <= ColumnName.Count(); i++)
            {
                for (int i1 = 1; i1 <= OName.Count(); i1++)
                {
                    if (ColumnName[i - 1] == OName[i1 - 1])
                    {
                        returnValue = i1;
                        return returnValue;
                    }
                }
            }

            //如果未匹配到该列，弹出窗体选择
            string PromptText = "请选择“" + ColumnName[0] + "”列";

            if (MustSelect == false)
            {
                PromptText = PromptText + Environment.NewLine + "如果不需要该列，请直接点击取消";
            }
            //捕获用户 直接点击取消 的情况
            try
            {
                returnValue = ExcelApp.InputBox(Prompt: PromptText, Type: 8).Column;
            }
            catch
            {
                returnValue = 0;
            }
            //判断选区是否超出有效区域
            if(returnValue > OName.Count())
            {
                MessageBox.Show("所选区域超出数据有效区域，请检查并重新选择");
                returnValue = 0;
            }

            return returnValue;
        }
        
        /// <summary>
        ///移动range数组中的列到新数组
        /// </summary>
        public void TrColumn(object[,] ORG, object[,] NRG,int AllRows, int OColumn, int NColumn)
        {
            //移动列
            for (int i = 1; i <= AllRows; i++)
            {
                try
                {
                    NRG[i - 1, NColumn - 1] = ORG[i, OColumn];
                }
                catch
                {
                    NRG[i - 1, NColumn - 1] = "";
                }
            }
        }

        /// <summary>
        ///检查某列是否全部为数字,NRG初始单元格为0,0
        /// </summary>
        public bool IsNumColumn(object[,] NRG, int Column, int StartRow, int EndRow)
        {
            bool returnValue = true;
            for (int i = StartRow; i < EndRow; i++)
            {
                if (NRG[i, Column] == null)
                {

                }
                else if (!IsNumber(NRG[i, Column].ToString()))
                {
                    MessageBox.Show("所选列第“" + (i + 1).ToString() + "”行不是数字格式，请检查");
                    returnValue = false;
                    return returnValue;
                }
            }
            return returnValue;
        }

        /// <summary>
        /// 创建工作表
        /// </summary>
        public void NewSheet(string SheetName)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            if (SheetExist(SheetName))
            {
                WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[SheetName];
                ExcelApp.DisplayAlerts = false;//关闭弹窗
                WST.Delete();
                ExcelApp.DisplayAlerts = true;//打开弹窗
            }
            WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets.Add(Type.Missing, (Excel.Worksheet)ExcelApp.ActiveSheet);
            WST.Name = SheetName;
            WST.Select();
        }

        /// <summary>
        /// 添加往来科目单sheet
        /// </summary>
        /// <param name="ORG">原始数组</param>
        /// <param name="AllRows">全部行数</param>
        /// <param name="AccountName">科目名称</param>
        /// <param name="OtherName">对方科目</param>
        /// <returns></returns>
        public bool AddCASheet(object[,] ORG, int AllRowsC, string AccountName, string OtherName)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            //定义新数组NRG
            object[,] NRG = new object[AllRowsC + 6, 14];
            bool returnValue = true;
            //读取首行列名
            NRG[0, 0] = "[客户编号]";
            NRG[0, 1] = "[客户名称]";
            NRG[0, 2] = "[一级科目]";
            NRG[0, 3] = "[明细科目]";
            NRG[0, 4] = "[期初余额]";
            NRG[0, 5] = "[期初重分类]";
            NRG[0, 6] = "[期初审定数]";
            NRG[0, 7] = "[本期借方]";
            NRG[0, 8] = "[本期贷方]";
            NRG[0, 9] = "[期末余额]";
            NRG[0, 10] = "[期末重分类]";
            NRG[0, 11] = "[期末审定数]";
            NRG[0, 12] = ORG[1, 9];
            NRG[0, 13] = "[函证]";

            //读取当前科目的行
            int i3 = 1;
            for (int i = 2; i <= AllRowsC; i++)
            {
                if (ORG[i, 3].ToString() == AccountName)
                {
                    //读入前4列
                    for (int i1 = 0; i1 < 5; i1++)
                    {
                        NRG[i3, i1] = ORG[i, i1 + 1];
                    }
                    //第5、6列
                    if (double.Parse(NRG[i3, 4].ToString()) < 0) { NRG[i3, 5] = -double.Parse(NRG[i3, 4].ToString()); }
                    NRG[i3, 6] = "=E" + (i3 + 1).ToString() + "+F" + (i3 + 1).ToString();
                    //读入7-9列
                    for (int i1 = 7; i1 < 10; i1++)
                    {
                        NRG[i3, i1] = ORG[i, i1 - 1];
                    }
                    //第10、11列
                    if (double.Parse(NRG[i3, 9].ToString()) < 0) { NRG[i3, 10] = -double.Parse(NRG[i3, 9].ToString()); }
                    NRG[i3, 11] = "=J" + (i3 + 1).ToString() + "+K" + (i3 + 1).ToString();
                    //读入12列
                    NRG[i3, 12] = ORG[i, 9];

                    i3 += 1;
                }
            }
            //读取重分类至当前科目的行
            int i4 = i3 + 1;
            for (int i = 2; i <= AllRowsC; i++)
            {
                //取期末余额小于0的对方科目行
                if (ORG[i, 3].ToString() == OtherName)
                {
                    if (double.Parse(ORG[i, 8].ToString()) < 0)
                    {
                        //读入前3列
                        for (int i1 = 0; i1 < 4; i1++)
                        {
                            NRG[i4, i1] = ORG[i, i1 + 1];
                        }

                        //第5、6列
                        if (double.Parse(ORG[i, 5].ToString()) < 0) { NRG[i4, 5] = -double.Parse(ORG[i, 5].ToString()); }
                        NRG[i4, 6] = "=F" + (i4 + 1).ToString();

                        //第10、11列
                        NRG[i4, 10] = -double.Parse(ORG[i, 8].ToString());
                        NRG[i4, 11] = "=K" + (i4 + 1).ToString();

                        i4 += 1;
                    }
                    else if(double.Parse(ORG[i, 5].ToString()) < 0)
                    {
                        //读入前3列
                        for (int i1 = 0; i1 < 4; i1++)
                        {
                            NRG[i4, i1] = ORG[i, i1 + 1];
                        }

                        //第5、6列
                        NRG[i4, 5] = -double.Parse(ORG[i, 5].ToString());
                        NRG[i4, 6] = "=F" + (i4 + 1).ToString();

                        i4 += 1;
                    }
                }
            }

            //添加末尾合计行
            if (i4 != i3 + 1)
            {
                i4 += 1;
            }
            NRG[i4, 1] = "合计";
            for (int i1 = 4; i1 < 12; i1++)
            {
                NRG[i4, i1] = "=SUM(" + CName(i1 + 1) + "2:" + CName(i1 + 1) + i4.ToString() + ")";
            }
            NRG[i4 + 1, 1] = "重分类至" + OtherName;
            NRG[i4 + 1, 4] = "=SUMIF(C:C,\"" + AccountName + "\",F:F)";
            NRG[i4 + 1, 9] = "=SUMIF(C:C,\"" + AccountName + "\",K:K)";
            NRG[i4 + 2, 1] = "从" + OtherName + "重分类";
            NRG[i4 + 2, 4] = "=SUMIF(C:C,\"" + OtherName + "\",F:F)";
            NRG[i4 + 2, 9] = "=SUMIF(C:C,\"" + OtherName + "\",K:K)";
            NRG[i4 + 3, 1] = "报表数";
            NRG[i4 + 4, 1] = "差异";
            NRG[i4 + 4, 6] = "=G" + (i4 + 1) + "-G" + (i4 + 4);
            NRG[i4 + 4, 11] = "=L" + (i4 + 1) + "-L" + (i4 + 4);

            //新建Sheet，并赋值
            NewSheet(AccountName);
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            WST.Range["A1:N" + (AllRowsC + 6).ToString()].Value2 = NRG;

            //更新总行数
            AllRowsC = AllRows("B");
            //改格式
            //定义rg为有效区域
            Excel.Range rg = WST.Range["A1:N" + AllRowsC.ToString()];
            //加框线
            rg.Borders.LineStyle = 1;
            //设置数字格式
            WST.Range["E2:L" + AllRowsC.ToString()].NumberFormatLocal = "#,##0.00 ";
            //自动列宽
            rg.EntireColumn.AutoFit();
            //函证列
            rg = WST.Range["N2:N" + (AllRowsC - 6).ToString()];
            AddData(rg, "函,补,");
            //首行颜色设置为灰色
            rg = WST.Range["A1:N1"];
            rg.Interior.ColorIndex = 15;
            //冻结行和列
            ExcelApp.ActiveWindow.SplitColumn = 2;
            ExcelApp.ActiveWindow.SplitRow = 1;
            ExcelApp.ActiveWindow.FreezePanes = true;

            //如果未选择辅助列，则删除这列
            if (AllRows("M") < 2)
            {
                WST.Range["M:M"].Delete();
            }

            return returnValue;
        }

        /// <summary>
        /// 添加数据验证
        /// </summary>
        /// <param name="rg">单元格区域</param>
        /// <param name="DateList">有效的验证序列</param>
        /// <returns></returns>
        public void AddData(Excel.Range rg, string DataList)
        {
            
            try
            {
                rg.Validation.Delete();
                rg.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlNotBetween, DataList, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            rg.Validation.InCellDropdown = true;
            rg.Validation.IgnoreBlank = true;
            rg.Value2 = "";
        }

        /// <summary>
        /// 将object转换为double,保留两位小数
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public double TD(object Value)
        {
            double returnValue = 0d;
            if (Value == null)
            {
                return Math.Round(returnValue, 2);
            }
            string inputValue = Value.ToString();
            double.TryParse(inputValue, out returnValue);
            returnValue = Math.Round(returnValue, 2);
            return returnValue;
        }

        /// <summary>
        /// 将不是数字的单元格标注黄色
        /// </summary>
        /// <param name="SelectRange"></param>
        public void ColorNotNum(string SelectRange)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            Excel.Range rg = WST.Range[SelectRange];
            object[,] ORG = rg.Value2;
            int StartRow = rg.Row;
            int StartColumn = rg.Column;

            //清除选区颜色
            rg.Interior.ColorIndex = 0;
            //寻找非数字单元格
            string CellsStr = "0";
            for (int i = 1; i <= ORG.GetLength(0);i++)
            {
                for (int i1 = 1; i1 <= ORG.GetLength(1); i1++)
                {
                    if (ORG[i, i1] == null) { break; }
                    if (!IsNumber(ORG[i, i1].ToString()))
                    {
                        CellsStr = CellsStr + "," + CName(StartColumn + i1 - 1) + (StartRow + i - 1);
                    }
                }
            }
            //修改颜色
            if (CellsStr != "0")
            {
                WST.Range[CellsStr.Remove(0, 2)].Interior.Color = Color.Yellow;
            }
        }
    }
}
