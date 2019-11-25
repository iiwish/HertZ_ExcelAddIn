using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HertZ_ExcelAddIn
{
    public partial class VerInfo : Form
    {
        public VerInfo()
        {
            InitializeComponent();
        }

        private void VerInfo_Load(object sender, EventArgs e)
        {
            //从我的文档读取配置
            //string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            //ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //从父节点Info中读取配置名为Vertion的值，该值为字符串
            //string VerInfo = clsConfig.ReadConfig<string>("Info", "Vertion", "0.0.0.01");
            try
            {
                label1.Text = "当前版本：" + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch
            {

            }
        }

        private void Manual_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("");
        }

        private void OnlineVideo_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("");
        }

    }
}
