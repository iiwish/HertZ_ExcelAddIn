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
        private readonly FunCtion FunC = new FunCtion();
        
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

            ExcelApp.Visible = false;//关闭Excel视图刷新
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

            ExcelApp.Visible = true;//打开Excel视图刷新
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
                OName.Add(ORG[1, i].ToString());
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
            ColumnName = new List<string> { "[本期借方]", "本期借方", "本年累计借方" };
            ColumnNumber = FunC.SelectColumn(ColumnName, OName, true);
            FunC.TrColumn(ORG, NRG, AllRows, ColumnNumber, 6);
            if (ColumnNumber == 0) { return; }
            if (!FunC.IsNumColumn(NRG, 5, 1, AllRows)) { return; }
            NRG[0, 5] = ColumnName[0];
            ColumnName.Clear();

            //选择[本期贷方]列
            ColumnName = new List<string> { "[本期贷方]", "本期贷方", "本年累计贷方" };
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
                NRG[0, 8] = "[" + NRG[0, 8].ToString() + "]";
            }
            else
            {
                NRG[0, 8] = ColumnName[0];
            }
            ColumnName.Clear();

            //删除sheet中的原始数据
            WST.Range["A:" + FunC.CName(AllColumns)].Delete();

            //写入数据
            WST.Range["A1:I" + AllRows.ToString()].Value2 = NRG;

            //释放数组
            ORG = null;
            ORG = NRG;
            NRG = null;

            //对一级科目去重
            List<string> SheetsName0 = new List<string> { "[一级科目]" };
            for (int i = 1; i < AllRows; i++)
            {
                try
                {
                    if (ORG[i, 2].ToString() != SheetsName0[SheetsName0.Count - 1])
                    {
                        SheetsName0.Add(ORG[i, 2].ToString());
                    }
                }
                catch(NullReferenceException)
                {
                    MessageBox.Show("一级科目列第" + (i + 1).ToString() + "行存在空值，请检查");
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
            DialogResult dr = MessageBox.Show("是否修改贷方科目以及贷方发生额列的正负号？"+ Environment.NewLine + "如果是SAP导出的往来款，请选择“是”", "请选择", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                for (int i = 1; i < AllRows; i++)
                {
                    if (SheetsName[ORG[i,2].ToString()] == "贷")
                    {
                        ORG[i, 4] = -double.Parse(ORG[i, 4].ToString());
                        ORG[i, 7] = -double.Parse(ORG[i, 7].ToString());
                    }
                }
            }

            ExcelApp.Visible = false;//关闭Excel视图刷新

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

            ExcelApp.Visible = true;//打开Excel视图刷新

        }

        //拆分账龄
        private void AgeOfAccount_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;

            //定义第二个表
            Excel.Worksheet WST2;

            string PromptText = "请选择上一年度账龄表";
            try
            {
                WST2 = ExcelApp.ActiveWorkbook.Worksheets[ExcelApp.InputBox(Prompt: PromptText, Type: 2).Replace("!", "").Replace("=", "")];
                WST2.Select();
            }
            catch
            {
                return;
            }
            //WST2.Select();
        }

        //往来款加工设置
        private void CurrentAccountSetting_Click(object sender, RibbonControlEventArgs e)
        {
            Form CASetting = new CASetting();
            CASetting.StartPosition = FormStartPosition.CenterScreen;
            CASetting.Show();
        }


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

            ExcelApp.Visible = false;//关闭Excel视图刷新

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
            WST.Select();
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

            ExcelApp.Visible = true;//打开Excel视图刷新
        }


        //版本信息
        private void VersionInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Form InfoForm = new VerInfo();
            InfoForm.StartPosition = FormStartPosition.CenterScreen;
            InfoForm.Show();
        }
    }
}
