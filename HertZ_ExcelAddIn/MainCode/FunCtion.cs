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
using System.IO;
using System.Collections;
using System.Data;

namespace HertZ_ExcelAddIn
{
    public class FunCtion
    {
        private Excel.Application ExcelApp;
        Excel.Worksheet WST;
        
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
            ExcelApp = Globals.ThisAddIn.Application;
            bool returnValue;

            try
            {
                WST = ExcelApp.Worksheets[SheetName];
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
            ExcelApp = Globals.ThisAddIn.Application;
            bool returnValue = false;

            if (!SheetExist(SheetName))
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
            ExcelApp = Globals.ThisAddIn.Application;
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
            ExcelApp = Globals.ThisAddIn.Application;
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
            ExcelApp = Globals.ThisAddIn.Application;
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
            ExcelApp = Globals.ThisAddIn.Application;
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
                    if (TD(NRG[i3, 4]) < 0) { NRG[i3, 5] = -TD(NRG[i3, 4]); }
                    NRG[i3, 6] = "=E" + (i3 + 1).ToString() + "+F" + (i3 + 1).ToString();
                    //读入7-9列
                    for (int i1 = 7; i1 < 10; i1++)
                    {
                        NRG[i3, i1] = ORG[i, i1 - 1];
                    }
                    //第10、11列
                    if (TD(NRG[i3, 9]) < 0) { NRG[i3, 10] = -TD(NRG[i3, 9]); }
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
                    if (TD(ORG[i, 8]) < 0)
                    {
                        //读入前3列
                        for (int i1 = 0; i1 < 4; i1++)
                        {
                            NRG[i4, i1] = ORG[i, i1 + 1];
                        }

                        //第5、6列
                        if (TD(ORG[i, 5]) < 0) { NRG[i4, 5] = -TD(ORG[i, 5]); }
                        NRG[i4, 6] = "=F" + (i4 + 1).ToString();

                        //第10、11列
                        NRG[i4, 10] = -TD(ORG[i, 8]);
                        NRG[i4, 11] = "=K" + (i4 + 1).ToString();

                        i4 += 1;
                    }
                    else if(TD(ORG[i, 5]) < 0)
                    {
                        //读入前3列
                        for (int i1 = 0; i1 < 4; i1++)
                        {
                            NRG[i4, i1] = ORG[i, i1 + 1];
                        }

                        //第5、6列
                        NRG[i4, 5] = -TD(ORG[i, 5]);
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
                return Math.Round(returnValue, 4);
            }
            string inputValue = Value.ToString();
            double.TryParse(inputValue, out returnValue);
            returnValue = Math.Round(returnValue, 4);
            return returnValue;
        }

        /// <summary>
        /// 将object转换为string
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public string TS(object Value)
        {
            string returnValue = "";
            if (Value == null)
            {
                return "";
            }
            returnValue = Value.ToString();
            return returnValue;
        }

        /// <summary>
        /// 将不是数字的单元格标注黄色
        /// </summary>
        /// <param name="SelectRange"></param>
        public void ColorNotNum(string SelectRange)
        {
            ExcelApp = Globals.ThisAddIn.Application;
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

        /// <summary>
        ///生成函证表功能中，将抽取的函证加入字典keyDic
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="PrKey"></param>
        /// <param name="KeyDic"></param>
        /// <returns></returns>
        public Dictionary<string, string> ConfirmationAddPrKey(string SheetName,string PrKey,Dictionary<string, string> KeyDic)
        {
            if (!SheetExist(SheetName)) { return KeyDic; }
            WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[SheetName];
            WST.Select();
            int AllRows1 = AllRows();
            int AllColumns1 = AllColumns();
            //主键列号
            int ColumnNumber1 = 0;
            //[函证]列号
            int ColumnNumber2 = 0;

            //将表格读入数组ORG
            object[,] ORG = WST.Range["A1:" + CName(AllColumns1) + AllRows1.ToString()].Value2;

            //寻找key列和函证列
            for (int i = 1; i <= AllColumns1; i++)
            {
                if (ORG[1, i].ToString() == "[" + PrKey +"]")
                {
                    ColumnNumber1 = i;
                }
                else if (ORG[1, i].ToString() == "[函证]")
                {
                    ColumnNumber2 = i;
                }
            }
            //如果找到了[函证]列和主键列
            if (ColumnNumber1 != 0 && ColumnNumber2 != 0)
            {
                for (int i = 2; i <= AllRows1; i++)
                {
                    if (ORG[i, ColumnNumber2] != null && ORG[i, ColumnNumber1] != null)
                    {
                        if (ORG[i, ColumnNumber2].ToString() == "函" && !KeyDic.ContainsKey(ORG[i, ColumnNumber1].ToString()))
                        {
                            KeyDic.Add(ORG[i, ColumnNumber1].ToString(), "函");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(SheetName + "表中未发现" + PrKey + "列或[函证]列，请检查");
            }
            return KeyDic;
        }
    
        /// <summary>
        /// 生成函证表功能中，补充同一客户不同科目的函证
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="PrKey"></param>
        /// <param name="KeyDic"></param>
        /// <param name="NRG"></param>
        /// <returns></returns>
        public object[,] ConfirmationAddCon(string SheetName, string PrKey, Dictionary<string, string> KeyDic, object[,] NRG)
        {
            if (!SheetExist(SheetName)) { return NRG; }
            WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[SheetName];
            WST.Select();
            int AllRows1 = AllRows();
            int AllColumns1 = AllColumns();
            int ColumnNumber1 = 0;//key
            int ColumnNumber2 = 0;//[函证]

            int ColumnNumber3 = 0;//客户名称或者客户编号（与key相反的另一个）
            int ColumnNumber4 = 0;//往来明细科目
            int ColumnNumber5 = 0;//审定期末余额

            //赋值ColumnName3
            string ColumnName3;
            if (PrKey == "客户编号")
            {
                ColumnName3 = "[客户名称]";
            }
            else
            {
                ColumnName3 = "[客户编号]";
            }


            //将表格读入数组ORG
            object[,] ORG = WST.Range["A1:" + CName(AllColumns1) + AllRows1.ToString()].Value2;

            //寻找指定列号
            for (int i = 1; i <= AllColumns1; i++)
            {
                if (ORG[1, i].ToString() == "[" + PrKey + "]")
                {
                    ColumnNumber1 = i;
                }
                else if (ORG[1, i].ToString() == "[函证]")
                {
                    ColumnNumber2 = i;
                }
                else if (ORG[1, i].ToString() == ColumnName3)
                {
                    ColumnNumber3 = i;
                }
                else if (ORG[1, i].ToString() == "[明细科目]")
                {
                    ColumnNumber4 = i;
                }
                else if (ORG[1, i].ToString() == "[期末审定数]")
                {
                    ColumnNumber5 = i;
                }
            }
            //审定余额列
            if (ColumnNumber5 == 0)
            {
                MessageBox.Show(SheetName + "表中未发现[期末审定数]列，请检查");
                return NRG;
            }

            //如果找到了主键列和[函证]列
            if (ColumnNumber1 != 0 && ColumnNumber2 != 0)
            {
                int i3 = int.Parse(NRG[0, 0].ToString());
                for (int i = 2; i <= AllRows1; i++)
                {
                    if(ORG[i, ColumnNumber2] != null && ORG[i, ColumnNumber2].ToString() == "补")
                    {
                        ORG[i, ColumnNumber2] = null;
                    }

                    if (ORG[i, ColumnNumber1]  != null && KeyDic.ContainsKey(ORG[i, ColumnNumber1].ToString()))
                    {
                        if(ORG[i, ColumnNumber2] != null && ORG[i, ColumnNumber2].ToString() == "函")
                        {
                            NRG[i3, 5] = "函"; 
                        }
                        else
                        {
                            ORG[i, ColumnNumber2] = "补";
                            //如果审定余额为0，则提前进行下一个循环
                            if (Math.Abs(TD(ORG[i, ColumnNumber5])) < 0.001d) { continue; }
                            NRG[i3, 5] = "补";
                        }

                        if (PrKey == "客户编号")
                        {
                            NRG[i3, 0] = ORG[i, ColumnNumber1];
                            NRG[i3, 1] = ORG[i, ColumnNumber3];
                        }
                        else
                        {
                            NRG[i3, 0] = ORG[i, ColumnNumber3];
                            NRG[i3, 1] = ORG[i, ColumnNumber1];
                        }

                        if (ColumnNumber4 != 0) { NRG[i3, 3] = ORG[i, ColumnNumber4]; }
                        if (ColumnNumber5 != 0) { NRG[i3, 4] = ORG[i, ColumnNumber5]; }
                        NRG[i3, 2] = SheetName;

                        if (i3 < 5999)
                        {
                            i3 += 1;
                        }
                        else
                        {
                            MessageBox.Show("设置为最多支持6000行函证，如有更大需求请自行修改代码");
                            return NRG;
                        }
                    }
                }

                NRG[0, 0] = i3;
                WST.Range["A1:" + CName(AllColumns1) + AllRows1.ToString()].Value2 = ORG;
            }
            else
            {
                MessageBox.Show(SheetName + "表中未发现[" + PrKey + "]列或[函证]列，请检查");
            }


            return NRG;
        }

        /// <summary>
        /// 得到一个汉字的拼音第一个字母，如果是一个英文字母则直接返回大写字母
        /// </summary>
        /// <param name="CnChar">单个汉字</param>
        /// <returns>单个大写字母</returns>
        public string GetSpellCode(string CnChar)
        {
            long iCnChar;
            byte[] arrCN = System.Text.Encoding.Default.GetBytes(CnChar);

            //如果是字母，则直接返回
            if (arrCN.Length == 1)
            {
                CnChar = CnChar.ToUpper();
            }
            else
            {
                int area = (short)arrCN[0];
                int pos = (short)arrCN[1];
                iCnChar = (area << 8) + pos;

                // iCnChar match the constant
                string letter = "ABCDEFGHJKLMNOPQRSTWXYZ";
                int[] areacode = { 45217, 45253, 45761, 46318, 46826, 47010, 47297, 47614, 48119, 49062, 49324, 49896, 50371, 50614, 50622, 50906, 51387, 51446, 52218, 52698, 52980, 53689, 54481, 55290 };
                for (int i = 0; i < 23; i++)
                {
                    if (areacode[i] <= iCnChar && iCnChar < areacode[i + 1])
                    {
                        CnChar = letter.Substring(i, 1);
                        break;
                    }
                }
            }
            return CnChar;
        }

        /// <summary>
        /// CheckBAJ的双击事件
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        /// <param name="Cancel"></param>
        public void CheckDoubleClick(object Sh,Excel.Range Target, ref bool Cancel)
        {
            WST = (Excel.Worksheet)Sh;
            object[,] rg;
            Excel.Worksheet WST2;

            int AllRows1 = AllRows("A",2);
            int AllColumns1 = AllColumns();
            int ColumnNum = 0;
            string i3;
            int i4;

            if (Target.Row > AllRows1 || Target.Row == 1) { return; }
            if (Target.Column > AllColumns1) { return; }
            if(Target.Count != 1) { return; }

            //检查是否已加工余额表序时账
            if (WST.Name == "余额表")
            {
                if (!SheetExist("序时账"))
                {
                    MessageBox.Show("请将序时账放入当前工作簿！");
                    return;
                }
                else
                {
                    WST = (Excel.Worksheet)Sh;
                    rg = WST.Range["A1:B1"].Value2;
                    if (rg[1, 1] == null || rg[1, 2] == null) { MessageBox.Show("请先加工余额表！"); return; }
                    if (rg[1, 1].ToString() != "[显示]" || rg[1, 2].ToString() != "[科目编码]")
                    {
                        MessageBox.Show(rg[1, 1].ToString() + rg[1, 2].ToString() +"请先加工余额表");
                        return;
                    }
                    else
                    {
                        //双击列超过第四列，跳转至序时账
                        if (Target.Column > 4)
                        {
                            WST2 = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets["序时账"];

                            rg = WST2.Range["A1:E1"].Value2;
                            if (rg[1, 1] == null || rg[1, 1].ToString() != "[辅助]") { MessageBox.Show("请先加工序时账！"); return; }
                            for (int i = 1; i < 8; i++)
                            {
                                if (rg[1, i] != null && rg[1, i].ToString() == "[科目编码]")
                                {
                                    ColumnNum = i;
                                    break;
                                }
                            }
                            if(ColumnNum == 0) { MessageBox.Show("未发现序时账的[科目编码]列！"); return; }

                            ExcelApp.ScreenUpdating = false;//关闭屏幕刷新

                            //取消筛选
                            if (WST2.AutoFilterMode) { WST2.AutoFilterMode = false; }

                            rg = WST.Range["A1:B" + AllRows1].Value2;
                            i3 = rg[Target.Row, 2].ToString();
                            WST2.Select();
                            AllRows1 = AllRows("A",2);
                            AllColumns1 = AllColumns();
                            rg = WST2.Range["A1:A" + AllRows1].Value2;

                            for(int i = 2; i <= AllRows1; i++)
                            {
                                rg[i, 1] = string.Format("=Left({0}{1},{2})", CName(ColumnNum), i, i3.Length);
                            }


                            //赋值
                            WST2.Range["A1:A" + AllRows1].Value2 = rg;

                            //筛选[显示]列
                            WST2.Range["A1:" + CName(AllColumns1) + AllRows1].AutoFilter(1, i3);
                        }
                        else//双击前四列，展开下级科目
                        {
                            ExcelApp.ScreenUpdating = false;//关闭屏幕刷新

                            //取消筛选
                            if (WST.AutoFilterMode) { WST.AutoFilterMode = false; }

                            AllRows1 = AllRows("A", 2);

                            rg = WST.Range["A1:B" + (AllRows1 + 1)].Value2;
                            rg[AllRows1 + 1, 2] = "1";
                            if(rg[Target.Row + 1, 2] == null) { return; }
                            if (rg[Target.Row, 2] == null) { return; }
                            if (rg[Target.Row + 1, 1] == null) { return; }
                            i4 = rg[Target.Row + 1, 2].ToString().Length;
                            if (i4 <= rg[Target.Row, 2].ToString().Length)
                            {
                                return;
                            }
                            
                            //修改下级科目[显示]列
                            if (rg[Target.Row+1, 1].ToString() == "0")
                            {
                                for (int i = Target.Row + 1; i <= AllRows1; i++)
                                {
                                    if (rg[i, 2] == null || rg[i, 2].ToString().Length == i4)
                                    {
                                        rg[i, 1] = 1;
                                    }
                                    else if (rg[i, 2].ToString().Length < i4)
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                for (int i = Target.Row + 1; i <= AllRows1; i++)
                                {
                                    if (rg[i, 2].ToString().Length < i4)
                                    {
                                        break;
                                    }
                                    rg[i, 1] = 0;
                                }
                            }
                            
                            
                            WST.Range["A1:B" + AllRows1].Value2 = rg;

                            //筛选[显示]列
                            WST.Range["A1:" + CName(AllColumns1) + AllRows1].AutoFilter(1, 1);
                        }
                    }
                }
            }
            else if(WST.Name == "序时账")
            {
                rg = WST.Range["A1:B1"].Value2;
                if (rg[1, 1] == null || rg[1, 1].ToString() != "[辅助]"){ MessageBox.Show("请先加工序时账"); return; }
                else if(rg[1,2]==null || rg[1,2].ToString() != "[日期&凭证号]") { MessageBox.Show("请先加工序时账"); return; }
                
                ExcelApp.ScreenUpdating = false;//关闭屏幕刷新
                i3 = WST.Range["B" + Target.Row].Value2;

                if (!SheetExist("联查凭证"))
                {
                    object[,] ORG = WST.Range["A1:" + CName(AllColumns1) +AllRows1].Value2;
                    NewSheet("联查凭证");
                    WST2 = (Excel.Worksheet)ExcelApp.ActiveSheet;
                    WST2.Range["A1:" + CName(AllColumns1) + AllRows1].Value2 = ORG;
                    
                    //调整表格格式

                    //首行颜色
                    WST2.Range[string.Format("A1:{0}1", CName(AllColumns1))].Interior.Color = Color.LightGray;
                    //加框线
                    WST2.Range["A1:" + CName(AllColumns1) + AllRows1].Borders.LineStyle = 1;
                    //设置数字格式
                    for(int i = 1; i < AllColumns1; i++)
                    {
                        if(ORG[1,i] != null && ORG[1, i].ToString() == "[借方金额]")
                        {
                            WST2.Range[string.Format("{0}2:{1}{2}", CName(i), CName(i + 1), AllRows1)].NumberFormatLocal = "#,##0.00 ";
                            break;
                        }
                    }
                   
                    //设置日期格式
                    WST2.Range["C2:C" + AllRows1].NumberFormatLocal = @"yyyy/m/d";
                    //ABC列靠左显示
                    WST2.Range["B2:M" + AllRows1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    //设置自动列宽
                    WST2.Columns["B:B"].EntireColumn.AutoFit();
                    WST2.Columns[string.Format("{0}:{1}", CName(AllColumns1 - 2), CName(AllColumns1))].EntireColumn.AutoFit();
                    //隐藏A、D列
                    WST2.Columns["A:A"].Hidden = true;
                    WST2.Columns["D:D"].Hidden = true;

                    ORG = null;

                    //冻结行和列
                    ExcelApp.ActiveWindow.SplitColumn = 2;
                    ExcelApp.ActiveWindow.SplitRow = 1;
                    ExcelApp.ActiveWindow.FreezePanes = true;
                    WST2.Tab.Color = Color.Blue;
                }
                else
                {
                    WST2 = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets["联查凭证"];
                    WST2.Select();
                }

                //取消筛选
                if (WST2.AutoFilterMode) { WST2.AutoFilterMode = false; }

                //筛选[显示]列
                WST2.Range["A1:" + CName(AllColumns1) + AllRows1].AutoFilter(2, i3);
            }
            else { return; }

            Cancel = true;

            ExcelApp.ScreenUpdating = true;//打开屏幕刷新
        }

        /// <summary>
        /// 判断公式是否需要增加括号
        /// </summary>
        /// <param name="FormulaStr"></param>
        /// <returns></returns>
        public bool AddParen(string FormulaStr)
        {
            bool returnValue = false;
            string InAParenStr;

            if (FormulaStr.Substring(0,2) == "=-")
            {
                FormulaStr = FormulaStr.Substring(2);
            }
            else
            {
                FormulaStr = FormulaStr.Substring(1);
            }

            int TempInt = FormulaStr.IndexOf("(");

            //如果包含括号
            if (TempInt != -1)
            {
                int StartInt = 0;
                //取去掉最外层括号的值
                InAParenStr = FormulaStr.Substring(TempInt+1, FormulaStr.LastIndexOf(")")- TempInt-1);
                
                //如果内侧括号不能成对出现，就加括号
                foreach(char c in InAParenStr)
                {
                    if (c == '(') { StartInt += 1; }
                    else if (c == ')')
                    {
                        if (StartInt == 0) { return true; }
                        else { StartInt -= 1; }
                    }
                }

                //第一个左括号前的字符串
                if(TempInt != 0) 
                { 
                    InAParenStr = FormulaStr.Substring(0, TempInt);
                    if (InAParenStr.IndexOf("+") != -1 || InAParenStr.IndexOf("-") != -1)
                    {
                        return true;
                    }
                }

                //最后一个右括号后的字符串
                TempInt = FormulaStr.LastIndexOf(")");
                if (TempInt != FormulaStr.Length-1)
                {
                    InAParenStr = FormulaStr.Substring(TempInt+1);
                    if (InAParenStr.IndexOf("+") != -1 || InAParenStr.IndexOf("-") != -1)
                    {
                        return true;
                    }
                }

            }
            else
            {
                //如果出现加减运算就加括号
                if (FormulaStr.IndexOf("+") != -1 || FormulaStr.IndexOf("-") != -1)
                {
                    return true;
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public DataTable OpenCSV(string filePath)
        {
            //UTF8Encoding encoding = Common.GetType(filePath); //Encoding.ASCII;//
            //Encoding encoding = GetType(filePath);
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            StreamReader sr = new StreamReader(fs, Encoding.UTF8);
            //StreamReader sr = new StreamReader(fs, encoding);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                //strLine = Common.ConvertStringUTF8(strLine, encoding);
                //strLine = Common.ConvertStringUTF8(strLine);

                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            fs.Close();
            return dt;
        }

        /// <summary>
        /// object转int
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public int TI(object Value)
        {
            int returnValue = 0;
            if (Value == null)
            {
                return 0;
            }
            string inputValue = Value.ToString();
            int.TryParse(inputValue, out returnValue);
            return returnValue;
        }

        /// <summary>
        /// object转decimal
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public decimal TDM(object Value)
        {
            decimal returnValue = 0;
            if (Value == null)
            {
                return 0;
            }
            string inputValue = Value.ToString();
            decimal.TryParse(inputValue, out returnValue);
            return returnValue;
        }

        /// <summary>
        /// 修改久其表格式
        /// </summary>
        /// <param name="RG"></param>
        /// <param name="ColumnWide"></param>
        public void JQChangeFont(string RG,List<decimal> ColumnWide)
        {
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            Excel.Range rg = ((Excel.Worksheet)ExcelApp.ActiveSheet).Range[RG];
            Excel.Range rg2;

            rg.WrapText = true;//自动换行

            rg.Interior.Pattern = Excel.XlPattern.xlPatternNone;
            rg.Interior.TintAndShade = 0;
            rg.Interior.PatternTintAndShade = 0;
            rg.Font.Name = "仿宋_GB2312";
            rg.Font.Size = 9;
            rg.Font.Name = "Arial Narrow";
            
            rg.Borders.get_Item(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;

            rg.VerticalAlignment = Excel.Constants.xlCenter;//垂直居中
            object[,] ORG = rg.Value2;
            if(ORG.GetLength(0) > 1)
            {
                if (TS(ORG[2, 1]) == "") 
                {
                    if(ORG.GetLength(0)>2 && TS(ORG[3, 1]) == "")
                    {
                        rg2 = WST.Range[string.Format("A4:{0}6", CName(ORG.GetLength(1)))];
                    }
                    else
                    {
                        rg2 = WST.Range[string.Format("A4:{0}5", CName(ORG.GetLength(1)))];
                    }
                }
                else
                {
                    rg2 = WST.Range[string.Format("A4:{0}4", CName(ORG.GetLength(1)))];
                }

                foreach(Excel.Range rgcell in rg2)
                {
                    if(TS(rgcell.Value2) != "")
                    {
                        rgcell.Value2 = TS(rgcell.Value2).Replace("\n", string.Empty);
                    }
                }

                rg2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                rg2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;
                rg2.Font.Bold = true;
                rg2.HorizontalAlignment = Excel.Constants.xlCenter;//表头水平居中

                if (TS(ORG[ORG.GetLength(0), 1]).Contains("合") && TS(ORG[ORG.GetLength(0), 1]).Contains("计"))
                {
                    rg2 = WST.Range[string.Format("A{0}:{1}{0}", ORG.GetLength(0) + 3, CName(ORG.GetLength(1)))];
                    rg2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                    rg2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;
                    rg2.Font.Bold = true;
                }

                //左右对齐
                if (TS(ORG[2, 1]) == "")
                {
                    if (ORG.GetLength(0) > 2 && TS(ORG[3, 1]) == "")
                    {
                        WST.Range[string.Format("A6:A{0}", ORG.GetLength(0) + 3)].HorizontalAlignment = Excel.Constants.xlLeft;
                        WST.Range[string.Format("B6:{0}{1}", CName(ORG.GetLength(1)), ORG.GetLength(0) + 3)].HorizontalAlignment = Excel.Constants.xlRight;
                    }
                    else
                    {
                        WST.Range[string.Format("A5:A{0}", ORG.GetLength(0) + 3)].HorizontalAlignment = Excel.Constants.xlLeft;
                        WST.Range[string.Format("B5:{0}{1}", CName(ORG.GetLength(1)), ORG.GetLength(0) + 3)].HorizontalAlignment = Excel.Constants.xlRight;
                    }
                }
                else
                {
                    WST.Range[string.Format("A4:A{0}", ORG.GetLength(0) + 3)].HorizontalAlignment = Excel.Constants.xlLeft;
                    WST.Range[string.Format("B4:{0}{1}", CName(ORG.GetLength(1)), ORG.GetLength(0) + 3)].HorizontalAlignment = Excel.Constants.xlRight;
                }
                rg2 = null;
            }

            for(int i = 1;i<= ColumnWide.Count; i++)
            {
                WST.Range[string.Format("{0}:{0}", CName(i))].ColumnWidth = Math.Round(ColumnWide[i - 1],2);
            }

            WST.Rows["4:" + (ORG.GetLength(0) + 3)].EntireRow.AutoFit();
        }

        /// <summary>
        /// 判断文件是否被占用
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool IsFileInUse(string fileName)
        {
            bool inUse = true;

            FileStream fs = null;
            try
            {

                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read,

                FileShare.None);

                inUse = false;
            }
            catch
            {

            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            return inUse;//true表示正在使用,false没有使用
        }

        /// <summary>
        /// 返回Link域中的表格名称
        /// </summary>
        /// <param name="CodeText"></param>
        /// <returns></returns>
        public string LinkSheet(string CodeText)
        {
            string TempStr = CodeText.Split('!')[0];//CodeText.Split('"')[3];
            TempStr = TempStr.Substring(TempStr.LastIndexOf('"') + 1);
            return TempStr;
        }
    }
}
