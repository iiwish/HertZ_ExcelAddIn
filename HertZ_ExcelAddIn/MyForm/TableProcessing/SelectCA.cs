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
    public partial class SelectCA : Form
    {
        //返回值
        public string ReturnValue { get; set; }
        public SelectCA()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "应收账款";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "预付账款";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "其他应收款";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "应付账款";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "预收账款";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.ReturnValue = "其他应付款";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
