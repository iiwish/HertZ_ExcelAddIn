using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace HertZ_ExcelAddIn.MyForm.Tool
{

    public partial class RegForm : Form
    {
        private Excel.Application ExcelApp;
        private Excel.Worksheet WST;
        //引用函数模块
        private readonly FunCtion FunC = new FunCtion();

        public RegForm()
        {
            InitializeComponent();
        }

        private void ConfirmBtn_Click(object sender, EventArgs e)
        {
            string TempStr = comboBox.Text;
            if (TempStr == null) { return; }
            if(TempStr.Length < 1) { return; }

            //如果是选择的下拉框，拆分内容
            if(TempStr.Contains(": ")) { TempStr = TempStr.Split(':')[1].Trim(); }
            if (TempStr == null) { return; }
            if (TempStr.Length < 1) { return; }

            try
            {
                Regex regex = new Regex(TempStr);
            }
            catch
            {
                MessageBox.Show("输入内容不符合正则表达式规范，请检查！");
                return;
            }

            ExcelApp = Globals.ThisAddIn.Application;
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;

            //读取选中区域
            Excel.Range rg;
            try
            {
                rg = ExcelApp.Selection;
            }
            catch
            {
                return;
            }

            //如果只选中一个单元格
            if (rg.Count == 1)
            {
                if (rg.Value2 != null)
                {
                    if (ChangeBtn.Text == "Replace")
                    {
                        rg.Value2 = Regex.Replace(FunC.TS(rg.Value2), TempStr, "");
                    }
                    else
                    {
                        rg.Value2 = FunC.TS(Regex.Match(FunC.TS(rg.Value2), TempStr));
                    }
                }
                return;
            }

            //如果选中了一个区域
            int AllRows;
            int AllColumns;
            object[,] ORGv;//原始数组ORGv 读取值
            object[,] NRG;//新数组NRG

            ORGv = rg.Value2;

            //限制列数，防止选择整行时多余的计算
            AllColumns = FunC.AllColumns(rg.Row, FunC.AllRows(FunC.CName(rg.Column)) + 10) - rg.Column + 1;//坑
            AllColumns = Math.Min(AllColumns, ORGv.GetLength(1));
            AllColumns = Math.Max(1, AllColumns);

            //限制行数

            AllRows = FunC.AllRows(FunC.CName(rg.Column), AllColumns) - rg.Row + 1;
            AllRows = Math.Min(AllRows, ORGv.GetLength(0));
            AllRows = Math.Max(1, AllRows);

            //定义新数组
            NRG = new object[AllRows, AllColumns];

            for (int i = 1; i <= AllColumns; i++)
            {
                for (int i1 = 1; i1 <= AllRows; i1++)
                {
                    if (ORGv[i1, i] != null)
                    {
                        if (ChangeBtn.Text == "Replace")
                        {
                            NRG[i1 - 1, i - 1] = Regex.Replace(FunC.TS(ORGv[i1, i]), TempStr, "");
                        }
                        else
                        {
                            NRG[i1 - 1, i - 1] = FunC.TS(Regex.Match(FunC.TS(ORGv[i1, i]), TempStr));
                        }
                    }
                }
            }

            //赋值
            WST.Range[FunC.CName(rg.Column) + rg.Row + ":" + FunC.CName(rg.Column + AllColumns - 1) + (rg.Row + AllRows - 1)].Value2 = NRG;

            ORGv = null;
            NRG = null;
            this.Close();
        }

        /// <summary>
        /// 改变匹配方式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChangeBtn_Click(object sender, EventArgs e)
        {
            if(ChangeBtn.Text == "Replace")
            {
                ChangeBtn.Text = "Match";
                string[] comboBoxList = new string[]
                {
                    @"匹配帐号是否合法(字母开头,允许5-16字节,允许字母数字下划线): ^[a-zA-Z][a-zA-Z0-9_]{4,15}$",
                    @"匹配国内电话号码(如0511-4405222或021-87888822): \d{3}-\d{8}|\d{4}-\d{7}",
                    @"匹配腾讯QQ号: [1-9][0-9]{4,}",@"匹配中国邮政编码: [1-9]\d{5}(?!\d)",
                    @"匹配身份证: \d{15}|\d{18}",@"匹配ip地址: \d+\.\d+\.\d+\.\d+",@"匹配正整数: [1-9]\d*$",
                    @"匹配负整数: -[1-9]\d*$",@"匹配整数: -?[1-9]\d*$",@"匹配非负整数（正整数+0）: [1-9]\d*|0$",
                    @"匹配非正整数（负整数+0）: -[1-9]\d*|0$",@"匹配正浮点数: [1-9]\d*\.\d*|0\.\d*[1-9]\d*$",
                    @"匹配负浮点数: -([1-9]\d*\.\d*|0\.\d*[1-9]\d*)$",@"匹配浮点数: -?([1-9]\d*\.\d*|0\.\d*[1-9]\d*|0?\.0+|0)$",
                    @"匹配非负浮点数（正浮点数+0）: [1-9]\d*\.\d*|0\.\d*[1-9]\d*|0?\.0+|0$",
                    @"匹配非正浮点数（负浮点数+0）: (-([1-9]\d*\.\d*|0\.\d*[1-9]\d*))|0?\.0+|0$",
                    @"匹配由26个英文字母组成的字符串: [A-Za-z]+$",@"匹配由26个英文字母的大写组成的字符串: [A-Z]+$",
                    @"匹配由26个英文字母的小写组成的字符串: [a-z]+$",@"匹配由数字和26个英文字母组成的字符串: [A-Za-z0-9]+$",
                    @"匹配由数字26个英文字母或者下划线组成的字符串: \w+$"
                };
                comboBox.Items.Clear();
                comboBox.Items.AddRange(comboBoxList);
            }
            else
            {
                ChangeBtn.Text = "Replace";
                string[] comboBoxList = new string[]
                {
                    @"去中文: [\u4e00-\u9fa5]",@"留中文: [^\u4e00-\u9fa5]",@"去字母: [A-Za-z]",
                    @"留字母: [^A-Za-z]",@"去数字: \d+(\.\d)?",@"留数字: ^\d+(\.\d)?",@"去数字字符: \d",
                    @"去非数字字符: \D",@"去换页字符: \f",@"去换行字符: \n",@"去回车符字符: \r",
                    @"去任何空白: \s",@"去任何非空白字符: \S",@"去任何非空白字符: [^\f\n\r\t\v]",
                    @"去制表字符: \t",@"去垂直制表符: \v",@"去包括下划线在内的任何字字符: \w",
                    @"去包括下划线在内的任何字字符: [A-Za-z0-9_]",@"去任何非字字符: \W",
                    @"去任何非字字符: [^A-Za-z0-9_]",@"去双字节字符(包括汉字在内): [^\x00-\xff]",
                    @"去HTML标记: <(\S*?)[^>]*>.*?</\1>|<.*?/>",@"去首尾空白字符: ^\s*|\s*$",
                    @"去Email地址: \w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
                };
                comboBox.Items.Clear();
                comboBox.Items.AddRange(comboBoxList);
            }
        }
    }
}
