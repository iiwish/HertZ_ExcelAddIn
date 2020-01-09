using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Deployment.Application;
using System.Windows.Forms;

namespace HertZ_ExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //获取当前版本信息
            string Nverinfo;
            string Overinfo;
            try
            {
                Nverinfo = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch
            {
                Nverinfo = "版本号获取异常";
            }

            //获取保存的版本信息

            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //从父节点VerInfo中读取配置名为VerNum的值，该值为字符串。默认为"0.0.0.0"
            Overinfo = clsConfig.ReadConfig<string>("VerInfo", "VerNum", "0.0.0.0");

            //判断版本是否相等
            if(Nverinfo == Overinfo){ return; }

            //写入父节点VerInfo中配置名VerNum的配置项
            clsConfig.WriteConfig("VerInfo", "VerNum", Nverinfo);

            string msg = "HertZ_ExcelAddin已更新" + Environment.NewLine + "当前版本为" + Nverinfo + Environment.NewLine;
            MessageBox.Show(msg + "修复了 加工往来款表时 数据区域有空单元格时会出错的bug");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
