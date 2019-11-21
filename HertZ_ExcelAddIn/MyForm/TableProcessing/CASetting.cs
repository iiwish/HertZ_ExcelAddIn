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
    public partial class CASetting : Form
    {
        public CASetting()
        {
            InitializeComponent();
        }

        private void CASetting_Load(object sender, EventArgs e)
        {
            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //从父节点CurrentAccount中读取配置名为AccountingFirmName的值，作为事务所名称，默认为致同
            AccountingFirmName.Text = clsConfig.ReadConfig<string>("CurrentAccount", "AccountingFirmName", "致同会计师事务所（特殊普通合伙）");
            //从父节点CurrentAccount中读取配置名为Auditee的值，作为被审计单位名称，默认为空
            Auditee.Text = clsConfig.ReadConfig<string>("CurrentAccount", "Auditee", "请修改");
            //从父节点CurrentAccount中读取配置名为ReplyAddress的值，作为回函地址，默认为致同
            ReplyAddress.Text = clsConfig.ReadConfig<string>("CurrentAccount", "ReplyAddress", "北京建外大街22号赛特大厦十五层");
            //从父节点CurrentAccount中读取配置名为PostalCode的值，作为回函邮编，默认为致同
            PostalCode.Text = clsConfig.ReadConfig<string>("CurrentAccount", "PostalCode", "100004");
            //从父节点CurrentAccount中读取配置名为AuditDeadline的值，作为审计截止日，默认为2019年12月31日
            AuditDeadline.Text = clsConfig.ReadConfig<string>("CurrentAccount", "AuditDeadline", "2019年12月31日");
            //从父节点CurrentAccount中读取配置名为Contact的值，作为联系人名称，默认为空
            Contact.Text = clsConfig.ReadConfig<string>("CurrentAccount", "Contact", "请修改");
            //从父节点CurrentAccount中读取配置名为Telephone的值，作为联系电话，默认为空
            Telephone.Text = clsConfig.ReadConfig<string>("CurrentAccount", "Telephone", "请修改");
            //从父节点CurrentAccount中读取配置名为Department的值，作为部门，默认为空
            Department.Text = clsConfig.ReadConfig<string>("CurrentAccount", "Department", "请修改");
            //从父节点CurrentAccount中读取配置名为Leading的值，作为部门负责人，默认为空
            Leading.Text = clsConfig.ReadConfig<string>("CurrentAccount", "Leading", "请修改");

        }

        private void ConfirmBtn_Click(object sender, EventArgs e)
        {
            //向我的文档写入配置文件
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //写入父节点CurrentAccount中配置名为AccountingFirmName的值，作为事务所名称
            clsConfig.WriteConfig("CurrentAccount", "AccountingFirmName", AccountingFirmName.Text.ToString());
            //写入父节点CurrentAccount中配置名为Auditee的值，作为被审计单位名称
            clsConfig.WriteConfig("CurrentAccount", "Auditee", Auditee.Text.ToString());
            //写入父节点CurrentAccount中读取配置名为ReplyAddress的值，作为回函地址
            clsConfig.WriteConfig("CurrentAccount", "ReplyAddress", ReplyAddress.Text.ToString());
            //从父节点CurrentAccount中读取配置名为PostalCode的值，作为回函邮编
            clsConfig.WriteConfig("CurrentAccount", "PostalCode", PostalCode.Text.ToString());
            //从父节点CurrentAccount中读取配置名为AuditDeadline的值，作为审计截止日，默认为2019年12月31日
            clsConfig.WriteConfig("CurrentAccount", "AuditDeadline", AuditDeadline.Text.ToString());
            //从父节点CurrentAccount中读取配置名为Contact的值，作为联系人名称，默认为空
            clsConfig.WriteConfig("CurrentAccount", "Contact", Contact.Text.ToString());
            //从父节点CurrentAccount中读取配置名为Telephone的值，作为联系电话，默认为空
            clsConfig.WriteConfig("CurrentAccount", "Telephone", Telephone.Text.ToString());
            //从父节点CurrentAccount中读取配置名为Department的值，作为部门，默认为空
            clsConfig.WriteConfig("CurrentAccount", "Department", Department.Text.ToString());
            //从父节点CurrentAccount中读取配置名为Leading的值，作为部门负责人，默认为空
            clsConfig.WriteConfig("CurrentAccount", "Leading", Leading.Text.ToString());

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
