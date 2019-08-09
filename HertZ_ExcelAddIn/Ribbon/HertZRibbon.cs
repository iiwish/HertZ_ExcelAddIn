using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace HertZ_ExcelAddIn
{
    public partial class HertZRibbon
    {
        private void HertZRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BalanceSheet_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("你说啥");
        }

        private void JournalSheet_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void 加工余额表_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void BalanceAndJournalSetting_Click(object sender, RibbonControlEventArgs e)
        {
            Form BAJSettingForm = new BAJSettingForm();
            BAJSettingForm.Show();
        }
    }
}
