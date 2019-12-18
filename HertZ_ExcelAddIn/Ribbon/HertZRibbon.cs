using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Drawing;

namespace HertZ_ExcelAddIn
{
    public partial class HertZRibbon
    {
        private Excel.Application ExcelApp;
        private Excel.Worksheet WST;

        //引用函数模块
        private readonly FunCtion FunC = new FunCtion();

        //判断浮点数是否等于0的参数
        public const double PRECISION = 0.0001d;

        private void HertZRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //清除后台没关干净的excel软件
            FunC.ClearBackExcel();
        }

        //加工余额表
        private void BalanceSheet_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;

            int AllRows;
            int AllColumns;
            int ColumnNumber;
            List<string> ColumnName;
            //原始表格数组ORG
            object[,] ORG;
            //目标新数组NRG
            object[,] NRG;

            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //选中科目余额表并继续
            if (FunC.SelectSheet("余额表") == false) { return; };
            WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets["余额表"];
            WST.Select();
            AllRows = FunC.AllRows();
            AllColumns = FunC.AllColumns();

            //规范原始数据
            if (FunC.RangeIsStandard() == false)
            {
                MessageBox.Show("请规范数据格式，保证数据内容不超出首行和首列");
                return;
            }

            //将表格读入数组ORG
            ORG = WST.Range["A1:" + FunC.CName(AllColumns) + AllRows.ToString()].Value2;
            //创建目标新数组NRG
            NRG = new object[AllRows, 9];

            //将列名读入List
            List<string> OName = new List<string> { };
            for (int i = 1; i <= AllColumns; i++)
            {
                OName.Add(ORG[1, i].ToString());
            }

            //选择[科目编码]列
            ColumnName = new List<string> { "[科目编码]", "科目编码", "科目编号", "科目号" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 1);
            NRG[0, 0] = ColumnName[0];
            ColumnName.Clear();

            //选择[科目名称]列
            ColumnName = new List<string> { "[科目名称]", "科目名称" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 2);
            if (ColumnNumber == 0) { return; }
            NRG[0, 1] = ColumnName[0];
            ColumnName.Clear();

            //自动匹配是否按方向和金额列示
            ColumnName = new List<string> { "[期初余额]", "期初余额", "期初金额", "审定期初数"};
            //匹配现有列名和目标列名
            DialogResult dr = DialogResult.None; //是否以方向和金额列示
            bool br = false;//是否退出循环
            for (int i = 1; i <= ColumnName.Count(); i++)
            {
                for (int i1 = 1; i1 <= OName.Count(); i1++)
                {
                    if (ColumnName[i - 1] == OName[i1 - 1])
                    {
                        dr = DialogResult.Yes;
                        br = true;
                        break;
                    }
                }
                if (br) { break; }
            }

            //如果未匹配到余额，自动匹配是否按借贷分别列示余额
            if (!br)
            {
                ColumnName = new List<string> { "[期初借方]", "期初借方", "期初借方金额", "期初借方余额" };
                //匹配现有列名和目标列名
                for (int i = 1; i <= ColumnName.Count(); i++)
                {
                    for (int i1 = 1; i1 <= OName.Count(); i1++)
                    {
                        if (ColumnName[i - 1] == OName[i1 - 1])
                        {
                            dr = DialogResult.No;
                            br = true;
                            break;
                        }
                    }
                    if (br) { break; }
                }
            }

            //如果还没有匹配到，弹出窗口让用户选择
            if (!br)
            {
                //选择期初期末余额列示方式
                dr = MessageBox.Show("期初期末余额是否按[借贷方向][金额]列示？" + Environment.NewLine + "若分别[借方余额][贷方余额]按列示，请选否", "请选择", MessageBoxButtons.YesNo);
            }
            
            if (dr == DialogResult.Yes)
            {
                //选择[方向]列,可选列
                ColumnName = new List<string> { "[方向]", "借贷方向" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, false);
                if (ColumnNumber != 0) { FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 3); }
                NRG[0, 2] = ColumnName[0];
                ColumnName.Clear();

                //选择[期初余额]列
                ColumnName = new List<string> { "[期初余额]", "期初余额", "期初金额", "期初数", "审定期初数" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 4);
                if (ColumnNumber == 0) { return; }
                NRG[0, 3] = ColumnName[0];
                ColumnName.Clear();

                //选择[期末余额]列
                ColumnName = new List<string> { "[期末余额]", "期末余额", "期末金额", "期末数", "审定期末数" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 7);
                if (ColumnNumber == 0) { return; }
                NRG[0, 6] = ColumnName[0];
                ColumnName.Clear();
            }
            else if(dr == DialogResult.No)
            {
                //选择[期初借方]列，先借用NRG的第5列存放数据
                ColumnName = new List<string> { "[期初借方]", "期初借方", "期初借方金额", "期初借方余额" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 5);
                if (ColumnNumber == 0) { return; }
                ColumnName.Clear();

                //选择[期初贷方]列，先借用NRG的第6列存放数据
                ColumnName = new List<string> { "[期初贷方]", "期初贷方", "期初贷方金额", "期初贷方余额" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 6);
                if (ColumnNumber == 0) { return; }
                ColumnName.Clear();

                //赋值[方向]列,[期初余额]列
                NRG[0, 2] = "[方向]";
                NRG[0, 3] = "[期初余额]";
                for (int i = 1; i < AllRows; i++)
                {
                    //规范[期初借方]列数据
                    if (string.IsNullOrWhiteSpace(NRG[i, 4].ToString()))
                    {
                        NRG[i, 4] = 0;
                    }
                    else
                    {
                        if(!FunC.IsNumber(NRG[i, 4].ToString()))
                        {
                            MessageBox.Show("所选[期初借方]列,第" + (i + 1) + "行存在非数值内容，请检查");
                            return;
                        }
                    }

                    //规范[期初贷方]列数据
                    if (string.IsNullOrWhiteSpace(NRG[i, 5].ToString()))
                    {
                        NRG[i, 5] = 0;
                    }
                    else
                    {
                        if (!FunC.IsNumber(NRG[i, 5].ToString()))
                        {
                            MessageBox.Show("所选[期初贷方]列,第" + (i + 1) + "行存在非数值内容，请检查");
                            return;
                        }
                    }

                    //计算[方向]列
                    if(double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString()) > 0)
                    {
                        NRG[i, 2] = "借";
                        NRG[i, 3] = double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString());
                    }
                    else if(double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString()) < 0)
                    {
                        NRG[i, 2] = "贷";
                        NRG[i, 3] = double.Parse(NRG[i, 5].ToString()) - double.Parse(NRG[i, 4].ToString());
                    }
                    else
                    {
                        NRG[i, 2] = "平";
                    }
                }

                //选择[期末借方]列，先借用NRG的第5列存放数据
                ColumnName = new List<string> { "[期末借方]", "期末借方", "期末借方金额", "期末借方余额" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 5);
                if (ColumnNumber == 0) { return; }
                ColumnName.Clear();

                //选择[期末贷方]列，先借用NRG的第6列存放数据
                ColumnName = new List<string> { "[期末贷方]", "期末贷方", "期末贷方金额", "期末贷方余额" };
                ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 6);
                if (ColumnNumber == 0) { return; }
                ColumnName.Clear();

                //赋值[期末余额]列
                NRG[0, 6] = "[期末余额]";
                for (int i = 1; i < AllRows; i++)
                {
                    //规范[期末借方]列数据
                    if (string.IsNullOrWhiteSpace(NRG[i, 4].ToString()))
                    {
                        NRG[i, 4] = 0;
                    }
                    else
                    {
                        if (!FunC.IsNumber(NRG[i, 4].ToString()))
                        {
                            MessageBox.Show("所选[期末借方]列,第" + (i + 1) + "行存在非数值内容，请检查");
                            return;
                        }
                    }

                    //规范[期末贷方]列数据
                    if (string.IsNullOrWhiteSpace(NRG[i, 5].ToString()))
                    {
                        NRG[i, 5] = 0;
                    }
                    else
                    {
                        if (!FunC.IsNumber(NRG[i, 5].ToString()))
                        {
                            MessageBox.Show("所选[期初贷方]列,第" + (i + 1) + "行存在非数值内容，请检查");
                            return;
                        }
                    }

                    //计算[期末余额]列
                    if (NRG[i, 2].ToString() == "借")
                    {
                        NRG[i, 6] = double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString());
                    }
                    else if (NRG[i, 2].ToString() == "贷")
                    {
                        NRG[i, 6] = double.Parse(NRG[i, 5].ToString()) - double.Parse(NRG[i, 4].ToString());
                    }
                    else
                    {
                        if (double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString()) > 0)
                        {
                            NRG[i, 2] = "借";
                            NRG[i, 6] = double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString());
                        }
                        else if (double.Parse(NRG[i, 4].ToString()) - double.Parse(NRG[i, 5].ToString()) < 0)
                        {
                            NRG[i, 2] = "贷";
                            NRG[i, 6] = double.Parse(NRG[i, 5].ToString()) - double.Parse(NRG[i, 4].ToString());
                        }
                    }
                }

            }
            else { return; }

            //选择[本年借方]列
            ColumnName = new List<string> { "[本年借方]", "本年借方", "本年借方累计", "借方金额累计", "审定借方发生额" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 5);
            if (ColumnNumber == 0) { return; }
            NRG[0, 4] = ColumnName[0];
            ColumnName.Clear();

            //选择[本年贷方]列
            ColumnName = new List<string> { "[本年贷方]", "本年贷方", "本年贷方累计", "贷方金额累计", "审定贷方发生额" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 6);
            if (ColumnNumber == 0) { return; }
            NRG[0, 5] = ColumnName[0];
            ColumnName.Clear();

            //规范[科目编码]列
            //使用长度区分科目层级
            Dictionary<int, string> CodeLen = new Dictionary<int, string> { };
            for (int i = 1; i < AllRows; i++)
            {
                try
                {
                    CodeLen.Add(NRG[i, 0].ToString().Length, NRG[i, 0].ToString());
                }
                catch { }
            }

            //字典排序
            CodeLen = CodeLen.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);

            //字典转list
            int[] CodeList = (from val in CodeLen select val.Key).ToArray<int>();
            CodeLen.Clear();

            //添加是否显示列
            NRG[0, 7] = "[显示]";
            for(int i = 1; i < AllRows; i++)
            {
                if(NRG[i,0].ToString().Length == CodeList[0])
                {
                    NRG[i, 7] = 1;
                }
                else
                {
                    NRG[i, 7] = 0;
                }
            }
            //添加科目层级列
            NRG[0, 8] = "[科目层级]";
            for (int i = 1; i < AllRows; i++)
            {
                for(int i1 = 1;i1 <= CodeList.Count();i1++)
                {
                    if(NRG[i, 0].ToString().Length == CodeList[i1-1])
                    {
                        NRG[i, 8] = i1;
                    }
                }
            }

            //删除sheet中的原始数据
            WST.Range["A:" + FunC.CName(AllColumns)].Delete();

            //写入数据
            WST.Range["A1:I" + AllRows.ToString()].Value2 = NRG;

            //释放数组
            ORG = null;

            ExcelApp.ScreenUpdating = false;//关闭Excel视图刷新
            //调整格式
            WST.Range["A1:I1"].Interior.Color = Color.LightGray;
            //按科目层级修改颜色
            Excel.Range rg;//定义单元格区域对象
            for (int i = 2; i <= AllRows; i++)
            {
                rg = WST.Range["A" + i + ":I" + i];
                switch (NRG[i-1, 8])
                {
                    case 1:
                        rg.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                        rg.Interior.TintAndShade = -0.249977111117893;
                        break;
                    case 2:
                        rg.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                        rg.Interior.TintAndShade = -0.149998474074526;
                        break;
                    case 3:
                        rg.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                        rg.Interior.TintAndShade = -4.99893185216834E-02;
                        break;
                    case 4:
                        rg.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
                        rg.Interior.TintAndShade = 0.799981688894314;
                        break;
                    case 5:
                        rg.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent6;
                        rg.Interior.TintAndShade = 0.799981688894314;
                        break;
                }
            }
            //释放数组
            NRG = null;

            //设置数字格式
            WST.Range["D2:G" + AllRows].NumberFormatLocal = "#,##0.00 ";
            //ABC列靠左显示
            WST.Range["A2:B" + AllRows].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //设置自动列宽
            WST.Columns["A:I"].EntireColumn.AutoFit();
            //筛选[显示]列
            WST.Range["A1:I" + AllRows].AutoFilter(8, 1);
            //隐藏[显示]列
            WST.Columns["H:H"].Hidden = true;

            ExcelApp.ScreenUpdating = true;//打开Excel视图刷新
            WST.Tab.Color = 3;//设置tab颜色为红色
        }

        //加工序时账
        private void JournalSheet_Click(object sender, RibbonControlEventArgs e)
        {

        }

        //账表加工设置
        private void BalanceAndJournalSetting_Click(object sender, RibbonControlEventArgs e)
        {
            //Form BAJSetting = new BAJSetting();
            //BAJSetting.StartPosition = FormStartPosition.CenterScreen;
            //BAJSetting.Show();
        }

        //加工往来款
        private void EditCurrentAccount_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;

            int AllRows;
            int AllColumns;
            int ColumnNumber;
            List<string> ColumnName;
            //原始表格数组ORG
            object[,] ORG;
            //目标新数组NRG
            object[,] NRG;

            //选中往来款明细表并继续
            if (FunC.SelectSheet("往来款明细") == false) { return; };
            WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets["往来款明细"];
            WST.Select();
            AllRows = FunC.AllRows();
            AllColumns = FunC.AllColumns();

            //规范原始数据
            if (FunC.RangeIsStandard() == false)
            {
                MessageBox.Show("请规范数据格式，保证数据内容不超出首行和首列");
                return;
            }

            //将表格读入数组ORG
            ORG = WST.Range["A1:" + FunC.CName(AllColumns) + AllRows.ToString()].Value2;
            //创建目标新数组NRG
            NRG = new object[AllRows,9];

            //将列名读入List
            List<string> OName = new List<string> { };
            for (int i = 1;i <= AllColumns; i++)
            {
                if(ORG[1, i] != null)
                {
                    OName.Add(ORG[1, i].ToString());
                }
                else
                {
                    OName.Add("0");
                }
            }

            //选择[客户编号]列
            ColumnName = new List<string> { "[客户编号]", "客户编号", "客户编码", "供应商编码" };
            ColumnNumber = FunC.SelectColumn(ColumnName,OName,true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG,NRG,AllRows, ColumnNumber, 1);
            NRG[0, 0] = ColumnName[0];
            ColumnName.Clear();

            //选择[客户名称]列
            ColumnName = new List<string> { "[客户名称]", "客户名称","供应商名称" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 2);
            if (ColumnNumber == 0) { return; }
            NRG[0, 1] = ColumnName[0];
            ColumnName.Clear();

            //选择[一级科目]列
            ColumnName = new List<string> { "[一级科目]", "一级科目" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 3);
            if (ColumnNumber == 0) { return; }
            NRG[0, 2] = ColumnName[0];
            ColumnName.Clear();

            //选择[明细科目]列,可选列
            ColumnName = new List<string> { "[明细科目]", "明细科目" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, false);
            if (ColumnNumber != 0) { FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 4); }
            NRG[0, 3] = ColumnName[0];
            ColumnName.Clear();

            //选择[期初余额]列
            ColumnName = new List<string> { "[期初余额]", "期初余额", "去年余额" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 5);
            if (ColumnNumber == 0) { return; }
            if (!FunC.IsNumColumn(NRG,4,1,AllRows)) { return; }
            NRG[0, 4] = ColumnName[0];
            ColumnName.Clear();

            //选择[本期借方]列
            ColumnName = new List<string> { "[本期借方]", "本期借方", "本年借方", "本年累计借方" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 6);
            if (ColumnNumber == 0) { return; }
            if (!FunC.IsNumColumn(NRG, 5, 1, AllRows)) { return; }
            NRG[0, 5] = ColumnName[0];
            ColumnName.Clear();

            //选择[本期贷方]列
            ColumnName = new List<string> { "[本期贷方]", "本期贷方", "本年贷方", "本年累计贷方" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 7);
            if (ColumnNumber == 0) { return; }
            if (!FunC.IsNumColumn(NRG, 6, 1, AllRows)) { return; }
            NRG[0, 6] = ColumnName[0];
            ColumnName.Clear();

            //选择[期末余额]列
            ColumnName = new List<string> { "[期末余额]", "期末余额" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 8);
            if (ColumnNumber == 0) { return; }
            if (!FunC.IsNumColumn(NRG, 7, 1, AllRows)) { return; }
            NRG[0, 7] = ColumnName[0];
            ColumnName.Clear();

            //选择[辅助项目]列，可选列
            ColumnName = new List<string> { "[辅助项目]", "辅助项目" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, false);
            if (ColumnNumber != 0)
            { 
                FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 9);
            }
            else
            {
                NRG[0, 8] = ColumnName[0];
            }
            ColumnName.Clear();

            //删除期初借贷余均为零的行
            DialogResult dr = MessageBox.Show("是否删除期初借贷余均为零的行？" , "请选择", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                //释放数组
                ORG = null;
                ORG = NRG;
                NRG = null;
                NRG = new object[AllRows, 9];

                for (int i1 = 0; i1 < 9; i1++)
                {
                    NRG[0, i1] = ORG[0, i1];
                }

                int i3 = 1;
                for (int i = 1; i < AllRows; i++)
                {
                    if ( FunC.TD(ORG[i, 4]) != 0d || FunC.TD(ORG[i, 5]) != 0d || FunC.TD(ORG[i, 6]) != 0d || FunC.TD(ORG[i, 7]) != 0d)
                    {
                        for (int i1 = 0; i1 < 9; i1++)
                        {
                            NRG[i3, i1] = ORG[i, i1];
                        }
                        i3 += 1;
                    }
                }
            }

            //删除sheet中的原始数据
            WST.Range["A:" + FunC.CName(AllColumns)].Delete();

            //写入数据
            WST.Range["A1:I" + AllRows.ToString()].Value2 = NRG;
            
            //释放数组
            ORG = null;
            NRG = null;
            AllRows = FunC.AllRows();
            //重新将表格读入数组ORG
            ORG = WST.Range["A1:I" + AllRows.ToString()].Value2;

            //对一级科目去重
            List<string> SheetsName0 = new List<string> { "[一级科目]" };
            for (int i = 2; i <= AllRows; i++)
            {
                try
                {
                    if (ORG[i, 3].ToString() != SheetsName0[SheetsName0.Count - 1])
                    {
                        SheetsName0.Add(ORG[i, 3].ToString());
                    }
                }
                catch(NullReferenceException)
                {
                    MessageBox.Show("一级科目列第" + i.ToString() + "行存在空值，请检查");
                    ORG = null;
                    return;
                }
                
            }
            List<string> SheetsName1 = new List<string> { };
            SheetsName0.Distinct().ToList().ForEach(s => SheetsName1.Add(s));

            //定义往来款字典
            Dictionary<string, string> SheetsName = new Dictionary<string, string>
            {
                { "应收账款", "借" },
                { "预付账款", "借" },
                { "其他应收款", "借" },
                { "应付账款", "贷" },
                { "预收账款", "贷" },
                { "其他应付款", "贷" }
            };

            //检查一级科目是否规范
            for (int i = 1; i < SheetsName1.Count; i++)
            {
                if(!SheetsName.ContainsKey(SheetsName1[i]))
                {
                    MessageBox.Show("一级科目列存在非往来款科目，请检查");
                    ORG = null;
                    return;
                }
            }

            //是否修改贷方往来款发生额方向和贷方发生额方向
            dr = MessageBox.Show("是否修改贷方科目以及贷方发生额列的正负号？"+ Environment.NewLine + "如果是SAP导出的往来款，请选择“是”", "请选择", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                for (int i = 2; i <= AllRows; i++)
                {
                    if (SheetsName[ORG[i,3].ToString()] == "贷")
                    {
                        ORG[i, 5] = -double.Parse(ORG[i, 5].ToString());
                        ORG[i, 8] = -double.Parse(ORG[i, 8].ToString());
                    }
                }
            }

            ExcelApp.ScreenUpdating = false;//关闭Excel视图刷新

            //应收账款表
            if (!FunC.AddCASheet(ORG, AllRows, "应收账款", "应付账款")) { return; }
            //预付账款表
            if (!FunC.AddCASheet(ORG, AllRows, "预付账款", "预收账款")) { return; }
            //其他应收款表
            if (!FunC.AddCASheet(ORG, AllRows, "其他应收款", "其他应付款")) { return; }
            //应付账款表
            if (!FunC.AddCASheet(ORG, AllRows, "应付账款", "应收账款")) { return; }
            //预收账款表
            if (!FunC.AddCASheet(ORG, AllRows, "预收账款", "预付账款")) { return; }
            //其他应付款表
            if (!FunC.AddCASheet(ORG, AllRows, "其他应付款", "其他应收款")) { return; }

            ExcelApp.ScreenUpdating = true;//打开Excel视图刷新

        }

        //拆分账龄
        private void AgeOfAccount_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;

            int AllRows;
            int AllColumns;
            int ColumnNumber;
            string ProjectName;
            List<string> ColumnName;
            //原始表格数组ORG
            object[,] ORG;
            //目标新数组NRG
            object[,] NRG;

            //选择科目名称
            using (var form = new SelectCA())
            {
                var result = form.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ProjectName = form.ReturnValue;
                }
                else
                {
                    return;
                }
            }

            //检查科目名称是否有效
            if (!FunC.SheetExist(ProjectName) )
            {
                MessageBox.Show("请将上年账龄表与加工完的往来款表放在同一工作簿");
                return;
            }
            else if(ExcelApp.Sheets[ProjectName].Range["A1"].Value.ToString() != "[客户编号]")
            {
                MessageBox.Show("请先加工往来款");
                return;
            }
            //定义第二个工作表
            Excel.Worksheet WST2 = ExcelApp.ActiveWorkbook.Sheets[ProjectName];

            DialogResult dr = MessageBox.Show(" 请在上年账龄表中使用该功能" + Environment.NewLine + "是否继续？", "请选择", MessageBoxButtons.YesNo);
            if (dr == DialogResult.No) { return;}

            //规范原始数据
            if (FunC.RangeIsStandard() == false)
            {
                MessageBox.Show("请规范数据格式，保证数据内容不超出首行和首列");
                return;
            }

            AllRows = FunC.AllRows();
            AllColumns = FunC.AllColumns();
            if (AllRows < 2) { return; }

            //将表格读入数组ORG
            ORG = WST.Range["A1:" + FunC.CName(AllColumns) + AllRows.ToString()].Value2;
            //创建目标新数组NRG
            NRG = new object[AllRows, 8];

            //将列名读入List
            List<string> OName = new List<string> { };
            for (int i = 1; i <= AllColumns; i++)
            {
                OName.Add(ORG[1, i].ToString());
            }

            //选择[客户名称]或[客户编号]列，做为引用的依据
            ColumnName = new List<string> { "[客户名称]或[客户编号]" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 1);
            NRG[0, 0] = ColumnName[0];
            ColumnName.Clear();

            //选择审定[期末余额]列
            ColumnName = new List<string> { "[期末余额]", "审定期末数", "期末审定数", "审定期末余额" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 2);
            if (!FunC.IsNumColumn(NRG, 1, 1, AllRows)) { return; }
            NRG[0, 1] = ColumnName[0];
            ColumnName.Clear();

            //选择账龄[1年以内]列
            ColumnName = new List<string> { "[1年以内]", "1年以内", "一年以内" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 3);
            if (!FunC.IsNumColumn(NRG, 2, 1, AllRows)) { return; }
            NRG[0, 2] = ColumnName[0];
            ColumnName.Clear();

            //选择账龄[1-2年]列
            ColumnName = new List<string> { "[1-2年]", "1-2年", "1~2年" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 4);
            if (!FunC.IsNumColumn(NRG, 3, 1, AllRows)) { return; }
            NRG[0, 3] = ColumnName[0];
            ColumnName.Clear();

            //选择账龄[2-3年]列
            ColumnName = new List<string> { "[2-3年]", "2-3年", "2~3年" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 5);
            if (!FunC.IsNumColumn(NRG, 4, 1, AllRows)) { return; }
            NRG[0, 4] = ColumnName[0];
            ColumnName.Clear();

            //选择账龄[3-4年]列
            ColumnName = new List<string> { "[3-4年]", "3-4年", "3~4年" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 6);
            if (!FunC.IsNumColumn(NRG, 5, 1, AllRows)) { return; }
            NRG[0, 5] = ColumnName[0];
            ColumnName.Clear();

            //选择账龄[4-5年]列
            ColumnName = new List<string> { "[4-5年]", "4-5年", "4~5年" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 7);
            if (!FunC.IsNumColumn(NRG, 6, 1, AllRows)) { return; }
            NRG[0, 6] = ColumnName[0];
            ColumnName.Clear();

            //选择账龄[5年以上]列
            ColumnName = new List<string> { "[5年以上]", "5年以上", "五年以上" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            if (ColumnNumber == 0) { return; }
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 8);
            if (!FunC.IsNumColumn(NRG, 7, 1, AllRows)) { return; }
            NRG[0, 7] = ColumnName[0];
            ColumnName.Clear();

            //借用数组ORG删除NRG中期末余额为0或空值的行
            ORG = new object[AllRows, 8];
            for (int i = 1; i < 8; i++)
            {
                ORG[0, i] = NRG[0, i];
            }
            int i3 = 1;
            for (int i = 1; i < AllRows; i++)
            {
                if(Math.Abs(FunC.TD(NRG[i, 1])) > PRECISION)
                {
                    for (int i1 = 0; i1 < 8; i1++)
                    {
                        ORG[i3, i1] = NRG[i, i1];
                    }
                    i3 += 1;
                }
            }

            NRG = null;
            NRG = new object[i3, 9];
            for (int i = 0; i < i3; i++)
            {
                for (int i1 = 0; i1 < 8; i1++)
                {
                    NRG[i, i1] = ORG[i, i1];
                }
            }


            //检查余额行
            if (i3 == 1)
            {
                MessageBox.Show("余额表中未发现余额大于零的行，请检查并重新开始");
                return;
            }

            //选中第二个工作表
            WST2.Select();

            //读取行列数
            AllRows = FunC.AllRows();
            AllColumns = FunC.AllColumns();

            //将表格读入数组ORG
            ORG = WST2.Range["A1:" + FunC.CName(AllColumns + 7) + AllRows].Value2;

            //找对应的列号
            int ColumnNumber1 = 0;//客户编号列或客户名称列
            int ColumnNumber2 = 0;//期初审定数
            int ColumnNumber3 = 0;//期末审定数
            int ColumnNumber4 = 0;//本期借方
            int ColumnNumber5 = 0;//本期贷方
            ColumnName = new List<string> { "[客户编号]", "[客户名称]", "[期初审定数]", "[期末审定数]", "[本期借方]", "[本期贷方]" };
            if (FunC.IsNumber(NRG[1, 0].ToString()))
            {
                ColumnNumber = 0;
            }
            else
            {
                ColumnNumber = 1;
            }
            for (int i = 1; i <= AllColumns; i++)
            {
                if(ORG[1,i].ToString() == ColumnName[ColumnNumber])
                {
                    ColumnNumber1 = i;
                }
                else if(ORG[1, i].ToString() == ColumnName[2])
                {
                    ColumnNumber2 = i;
                }
                else if (ORG[1, i].ToString() == ColumnName[3])
                {
                    ColumnNumber3 = i;
                }
                else if (ORG[1, i].ToString() == ColumnName[4])
                {
                    ColumnNumber4 = i;
                }
                else if (ORG[1, i].ToString() == ColumnName[5])
                {
                    ColumnNumber5 = i;
                }
            }
            //检查是否匹配成功
            if(ColumnNumber1 == 0 || ColumnNumber2 == 0 || ColumnNumber3 == 0 || ColumnNumber4 == 0 || ColumnNumber5 == 0)
            {
                MessageBox.Show("未匹配到[期初审定数]、[期末审定数]、[本期借方]或[本期贷方]列，请检查并重新开始");
                return;
            }

            ////检查指定列是否有非数字内容
            //ExcelApp.ScreenUpdating = false;//关闭Excel视图刷新
            //FunC.ColorNotNum(FunC.CName(ColumnNumber2) + "2:" + FunC.CName(ColumnNumber2) + AllRows);
            //FunC.ColorNotNum(FunC.CName(ColumnNumber3) + "2:" + FunC.CName(ColumnNumber3) + AllRows);
            //FunC.ColorNotNum(FunC.CName(ColumnNumber4) + "2:" + FunC.CName(ColumnNumber4) + AllRows);
            //FunC.ColorNotNum(FunC.CName(ColumnNumber5) + "2:" + FunC.CName(ColumnNumber5) + AllRows);
            //ExcelApp.ScreenUpdating = true;//打开Excel视图刷新
            //改表头
            for (int i = 1;i <= 6; i++)
            {
                ORG[1, AllColumns + i] = NRG[0,i + 1];
            }
            //匹配余额
            for (int i = 2;i <= AllRows; i++)
            {
                //如果期末余额为0，则账龄为0
                if(Math.Abs(FunC.TD(ORG[i,ColumnNumber3])) <= PRECISION)
                {
                    for(int i1 = 1; i1 <= 6; i1++)
                    {
                        ORG[i, AllColumns + i1] = 0;
                    }
                }
                //如果期初余额为0，则账龄为[1年以内]
                else if(Math.Abs(FunC.TD(ORG[i, ColumnNumber2])) <= PRECISION)
                {
                    ORG[i, AllColumns + 1] = ORG[i, ColumnNumber3];
                    for (int i1 = 2; i1 <= 6; i1++)
                    {
                        ORG[i, AllColumns + i1] = 0;
                    }
                }
                //如果借方科目借方金额大于期末余额
                else if ((ProjectName.Contains("应收") || ProjectName.Contains("预付")) && FunC.TD(ORG[i, ColumnNumber4]) >= FunC.TD(ORG[i, ColumnNumber3]))
                {
                    ORG[i, AllColumns + 1] = FunC.TD(ORG[i, ColumnNumber3]);
                    for (int i2 = 2; i2 <= 6; i2++)
                    {
                        ORG[i, AllColumns + i2] = 0;
                    }
                }
                //如果贷方科目贷方金额大于期末余额
                else if ((ProjectName.Contains("应付") || ProjectName.Contains("预收")) && FunC.TD(ORG[i, ColumnNumber5]) >= FunC.TD(ORG[i, ColumnNumber3]))
                {
                    ORG[i, AllColumns + 1] = FunC.TD(ORG[i, ColumnNumber3]);
                    for (int i2 = 2; i2 <= 6; i2++)
                    {
                        ORG[i, AllColumns + i2] = 0;
                    }
                }
                //否则从账龄表中匹配
                else
                {
                    for (int i1 = 1; i1 < i3; i1++)
                    {
                        //如果不匹配，直接下一个
                        if (Math.Abs(FunC.TD(ORG[i, ColumnNumber2]) - FunC.TD(NRG[i1,1])) > PRECISION || ORG[i, ColumnNumber1].ToString() != NRG[i1, 0].ToString())
                        {
                            continue;
                        }
                        //检查是否被匹配过，如果被匹配过直接下一个，防止出现同一供应商多个余额相等的情况
                        if (NRG[i1, 8] != null)
                        {
                            ORG[i, AllColumns + 1] = "请手动计算该供应商的账龄";
                            continue;
                        }
                        
                        NRG[i1, 8] = 1;
                        //检查本期发生额，如果为0则平移账龄
                        if (Math.Abs(FunC.TD(ORG[i, ColumnNumber4])) > PRECISION && Math.Abs(FunC.TD(ORG[i, ColumnNumber5])) > PRECISION)
                        {
                            ORG[i, AllColumns + 1] = 0;
                            for (int i2 = 2; i2 <= 5; i2++)
                            {
                                ORG[i, AllColumns + i2] = FunC.TD(NRG[i1, i2]);
                            }
                            ORG[i, AllColumns + 6] = FunC.TD(NRG[i1, 6]) + FunC.TD(NRG[i1, 7]);
                        }
                        //如果本期有发生额，则计算账龄
                        else
                        {
                            if (ProjectName.Contains("应收") || ProjectName.Contains("预付"))
                            {
                                //一年以内
                                ORG[i, AllColumns + 1] = FunC.TD(ORG[i, ColumnNumber4]);
                                //1-2年
                                ORG[i, AllColumns + 2] = Math.Min(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber4]), FunC.TD(NRG[i1, 2]));
                                //2-3年
                                ORG[i, AllColumns + 3] = Math.Min(Math.Max(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber4]) - FunC.TD(NRG[i1, 2]), 0), FunC.TD(NRG[i1, 3]));
                                //3-4年
                                ORG[i, AllColumns + 4] = Math.Min(Math.Max(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber4]) - FunC.TD(NRG[i1, 2]) - FunC.TD(NRG[i1, 3]), 0), FunC.TD(NRG[i1, 4]));
                                //4-5年
                                ORG[i, AllColumns + 5] = Math.Min(Math.Max(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber4]) - FunC.TD(NRG[i1, 2]) - FunC.TD(NRG[i1, 3]) - FunC.TD(NRG[i1, 4]), 0), FunC.TD(NRG[i1, 5]));
                                //5年以上
                                ORG[i, AllColumns + 6] = FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, AllColumns + 1]) - FunC.TD(ORG[i, AllColumns + 2]) - FunC.TD(ORG[i, AllColumns + 3]) - FunC.TD(ORG[i, AllColumns + 4]) - FunC.TD(ORG[i, AllColumns + 5]);

                            }
                            else
                            {
                                //一年以内
                                ORG[i, AllColumns + 1] = FunC.TD(ORG[i, ColumnNumber5]);
                                //1-2年
                                ORG[i, AllColumns + 2] = Math.Min(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber5]), FunC.TD(NRG[i1, 2]));
                                //2-3年
                                ORG[i, AllColumns + 3] = Math.Min(Math.Max(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber5]) - FunC.TD(NRG[i1, 2]), 0), FunC.TD(NRG[i1, 3]));
                                //3-4年
                                ORG[i, AllColumns + 4] = Math.Min(Math.Max(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber5]) - FunC.TD(NRG[i1, 2]) - FunC.TD(NRG[i1, 3]), 0), FunC.TD(NRG[i1, 4]));
                                //4-5年
                                ORG[i, AllColumns + 5] = Math.Min(Math.Max(FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, ColumnNumber5]) - FunC.TD(NRG[i1, 2]) - FunC.TD(NRG[i1, 3]) - FunC.TD(NRG[i1, 4]), 0), FunC.TD(NRG[i1, 5]));
                                //5年以上
                                ORG[i, AllColumns + 6] = FunC.TD(ORG[i, ColumnNumber3]) - FunC.TD(ORG[i, AllColumns + 1]) - FunC.TD(ORG[i, AllColumns + 2]) - FunC.TD(ORG[i, AllColumns + 3]) - FunC.TD(ORG[i, AllColumns + 4]) - FunC.TD(ORG[i, AllColumns + 5]);
                            }
                        }
                    }
                }

                //最后增加一列验证列
                ORG[i, AllColumns + 7] = "=ABS(" + FunC.CName(ColumnNumber3) + i + "-sum(" + FunC.CName(AllColumns + 1) + i + ":" + FunC.CName(AllColumns + 6) + i + "))<0.01";
            }

            //赋值
            WST2.Range["A1:" + FunC.CName(AllColumns + 7) + AllRows].Value2 = ORG;
            //定义rg为有效区域
            Excel.Range rg = WST2.Range["A1:" + FunC.CName(AllColumns + 7) + AllRows];
            //加框线
            rg.Borders.LineStyle = 1;
            //设置首行颜色为灰色
            rg = WST2.Range["A1:" + FunC.CName(AllColumns + 7) + "1"];
            rg.Interior.ColorIndex = 15;

        }

        //往来款加工设置
        private void CurrentAccountSetting_Click(object sender, RibbonControlEventArgs e)
        {
            Form CASetting = new CASetting
            {
                StartPosition = FormStartPosition.CenterScreen
            };
            CASetting.Show();
        }

        //对比两列数
        private void CompareTwoColumns_Click(object sender, RibbonControlEventArgs e)
        {
            int AllRows;
            string SelectColomn;
            string SelectColomn2;
            object[,] ORG;//原始数组ORG
            object[,] NRG;//新数组NRG
            string[,] ARG;//计算用数组

            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            //WBK = ExcelApp.ActiveWorkbook;

            //选择第一列，并捕获用户 直接点击取消 的情况
            try
            {
                SelectColomn = FunC.CName(ExcelApp.InputBox(Prompt: "请选择第一列", Type: 8).Column);
            }
            catch
            {
                return;
            }

            //获取所选列的行数
            AllRows = FunC.AllRows(SelectColomn);
            //如果行数为1，则终止程序
            if (AllRows < 2)
            {
                MessageBox.Show("所选列行数小于2，请重新开始");
                return;
            }

            //将所选列的赋值至数组
            ORG = WST.Range[SelectColomn + "1:" + SelectColomn + AllRows].Value2;

            //选择第二列，并捕获用户 直接点击取消 的情况
            try
            {
                SelectColomn2 = FunC.CName(ExcelApp.InputBox(Prompt: "请选择第二列", Type: 8).Column);
            }
            catch
            {
                return;
            }

            //获取所选列的行数
            AllRows = FunC.AllRows(SelectColomn2);
            //如果行数为1，则终止程序
            if (AllRows < 2)
            {
                MessageBox.Show("所选列行数小于2，请重新开始");
                return;
            }

            //将所选列的赋值至数组
            NRG = WST.Range[SelectColomn2 + "1:" + SelectColomn2 + AllRows].Value2;

            ExcelApp.ScreenUpdating = false;//关闭Excel视图刷新

            AllRows = Math.Max(ORG.GetLength(0), NRG.GetLength(0));
            ARG = new string[AllRows, 2];
            //将数组org存入
            for (int i = 1; i <= ORG.GetLength(0); i++)
            {
                if (ORG[i, 1] != null)
                {
                    ARG[i - 1, 0] = ORG[i, 1].ToString();
                }
                else
                {
                    ARG[i - 1, 0] = "0";
                }
            }
            ORG = null;//释放数组
            //将数组nrg存入
            for (int i = 1; i <= NRG.GetLength(0); i++)
            {
                if (NRG[i, 1] != null)
                {
                    ARG[i - 1, 1] = NRG[i, 1].ToString();
                }
                else
                {
                    ARG[i - 1, 1] = "0";
                }
            }
            NRG = null;//释放数组

            //重新定义NRG存放计算过程
            int[,] Arr = new int[AllRows, 8];

            //计算是否重复出现
            for (int i = 0; i < AllRows; i++)
            {
                for (int i1 = 0; i1 < AllRows; i1++)
                {
                    //Arr第三列表示第一列中重复的次数
                    if (ARG[i, 0] != null && ARG[i, 0] == ARG[i1, 0])
                    {
                        Arr[i, 2] = Arr[i, 2] + 1;
                    }

                    //Arr第四列表示第二列中重复的次数
                    if (ARG[i, 1] != null && ARG[i, 1] == ARG[i1, 1])
                    {
                        Arr[i, 3] = Arr[i, 3] + 1;
                    }

                    //第五列表示第一列数在第二列中出现的次数
                    if (ARG[i, 0] != null && ARG[i, 0] == ARG[i1, 1])
                    {
                        Arr[i, 4] = Arr[i, 4] + 1;
                    }

                    //第六列表示第二列数在第一列中出现的次数
                    if (ARG[i, 1] != null && ARG[i, 1] == ARG[i1, 0])
                    {
                        Arr[i, 5] = Arr[i, 5] + 1;
                    }
                }
            }

            for (int i = 0; i < AllRows; i++)
            {
                for (int i1 = 0; i1 < AllRows; i1++)
                {
                    //第七列为第二列中是否存在与第一列数相同的值
                    if (ARG[i, 0] == ARG[i1, 1] && Arr[i, 2] == Arr[i, 4])
                    {
                        Arr[i, 6] = 1;
                    }

                    //第八列为第一列中是否存在与第二列数相同的值
                    if (ARG[i, 1] == ARG[i1, 0] && Arr[i, 3] == Arr[i, 5])
                    {
                        Arr[i, 7] = 1;
                    }
                }
            }

            //调整单元格颜色
            //调整第二列的格式
            WST.Columns[SelectColomn2 + ":" + SelectColomn2].Interior.ColorIndex = 0;
            string CellsRange = "0";
            for (int i = 0; i < AllRows; i++)
            {
                if (Arr[i, 7] != 1 && ARG[i, 1] != null && ARG[i, 1] != "0")
                {
                    CellsRange = CellsRange + "," + SelectColomn2 + (i+1);
                }
            }
            if (CellsRange != "0")
            {
                WST.Range[CellsRange.Remove(0, 2)].Interior.Color = Color.Yellow;
            }
            //调整第一列的格式
            WST.Columns[SelectColomn + ":" + SelectColomn].Interior.ColorIndex = 0;
            CellsRange = "0";
            for (int i = 0; i < AllRows; i++)
            {
                if (Arr[i, 6] != 1 && ARG[i, 0] != null && ARG[i, 0] != "0")
                {
                    CellsRange = CellsRange + "," + SelectColomn + (i + 1);
                }
            }
            if (CellsRange != "0")
            {
                WST.Range[CellsRange.Remove(0, 2)].Interior.Color = Color.Yellow;
            }

            ExcelApp.ScreenUpdating = true;//打开Excel视图刷新
        }

        //检查非数字单元格
        private void CheckNum_Click(object sender, RibbonControlEventArgs e)
        {

        }

        //版本信息
        private void VersionInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Form InfoForm = new VerInfo
            {
                StartPosition = FormStartPosition.CenterScreen
            };
            InfoForm.Show();
        }

    }
}
