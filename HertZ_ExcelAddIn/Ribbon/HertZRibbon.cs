using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

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

            //选择期初期末余额列示方式
            DialogResult dr = MessageBox.Show("期初期末余额是否按[借贷方向][金额]列示？" + Environment.NewLine + "若以[借方余额][贷方余额]方式列示，请选否", "请选择", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                
            }




        }

        private void JournalSheet_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void BalanceAndJournalSetting_Click(object sender, RibbonControlEventArgs e)
        {
            Form BAJSetting = new BAJSetting();
            BAJSetting.StartPosition = FormStartPosition.CenterScreen;
            BAJSetting.Show();
        }

        private void CurrentAccount_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

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

        private void AgeOfAccount_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void CurrentAccountSetting_Click(object sender, RibbonControlEventArgs e)
        {
            Form CASetting = new CASetting();
            CASetting.StartPosition = FormStartPosition.CenterScreen;
            CASetting.Show();
        }

        private void VersionInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Form InfoForm = new VerInfo();
            InfoForm.StartPosition = FormStartPosition.CenterScreen;
            InfoForm.Show();
        }
    }
}
