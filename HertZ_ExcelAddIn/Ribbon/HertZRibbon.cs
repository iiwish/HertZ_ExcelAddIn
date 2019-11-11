using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace HertZ_ExcelAddIn
{
    public partial class HertZRibbon
    {
        private Excel.Application ExcelApp;
        private FunCtion FunC = new FunCtion();
        private void HertZRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BalanceSheet_Click(object sender, RibbonControlEventArgs e)
        {
            new B_TableProcessing().testbox();
        }

        private void JournalSheet_Click(object sender, RibbonControlEventArgs e)
        {

        }


        private void BalanceAndJournalSetting_Click(object sender, RibbonControlEventArgs e)
        {
            Form BAJSettingForm = new BAJSettingForm();
            BAJSettingForm.Show();
        }

        private void EditCurrentAccount_Click(object sender, RibbonControlEventArgs e)
        {
            //ExcelApp.Visible = false;//关闭Excel视图刷新

            //选中往来款明细表并继续
            if (FunC.SelectSheet("往来款明细") == false) { return; };


            MessageBox.Show(FunC.RangeIsStandard().ToString());

        }
    }
}
