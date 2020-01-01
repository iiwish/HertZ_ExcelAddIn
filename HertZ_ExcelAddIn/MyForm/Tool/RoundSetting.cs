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
    public partial class RoundSetting : Form
    {
        public RoundSetting()
        {
            InitializeComponent();
        }

        private void RoundSetting_Load(object sender, EventArgs e)
        {
            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //从父节点Round中读取配置名为Num的值，默认为2
            numericUpDown1.Value = clsConfig.ReadConfig<int>("Round", "Num", 2);
        }

        private void ConfirmBtn_Click(object sender, EventArgs e)
        {
            //向我的文档写入配置文件
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //写入父节点Round中配置名Num(自动识别编码）的配置项
            clsConfig.WriteConfig("Round", "Num", numericUpDown1.Value.ToString());

            //关闭窗体
            this.Close();
        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            //关闭窗体
            this.Close();
        }
    }
}
