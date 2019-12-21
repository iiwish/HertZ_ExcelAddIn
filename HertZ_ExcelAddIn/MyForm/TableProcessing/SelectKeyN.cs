using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HertZ_ExcelAddIn
{
    public partial class SelectKeyN : Form
    {
        public string ReturnValue { get; set; }
        public SelectKeyN()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "客户编号";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "客户名称";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
