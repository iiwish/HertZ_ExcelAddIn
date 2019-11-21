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
    public partial class BAJSettingForm : Form
    {
        
        public BAJSettingForm()
        {
            InitializeComponent();
        }

        private void BAJSettingForm_Load(object sender, EventArgs e)
        {
            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //从父节点BalanceAndJournal中读取配置名SubjectCodeButton1为的值，该值为布尔值。默认为true
            SubjectCodeButton1.Checked = clsConfig.ReadConfig<bool>("BalanceAndJournal", "SubjectCodeButton1", true);
            
            //从父节点BalanceAndJournal中读取配置名SubjectCodeButton2为的值，该值为布尔值。默认为false
            SubjectCodeButton2.Checked = clsConfig.ReadConfig<bool>("BalanceAndJournal", "SubjectCodeButton2", false);
            //从父节点BalanceAndJournal中读取配置名SubjectCodeSign为的值，该值为字符串。默认为"."
            SubjectCodeSign.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeSign", ".");

            //从父节点BalanceAndJournal中读取配置名SubjectCodeButton3为的值，该值为布尔值。默认为false
            SubjectCodeButton3.Checked = clsConfig.ReadConfig<bool>("BalanceAndJournal", "SubjectCodeButton3", false);
            //从父节点BalanceAndJournal中读取配置名SubjectCodeLength1为的值，该值为字符串。默认为"4"
            SubjectCodeLength1.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeLength1", "4");
            //从父节点BalanceAndJournal中读取配置名SubjectCodeLength2为的值，该值为字符串。默认为"2"
            SubjectCodeLength2.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeLength2", "2");
            //从父节点BalanceAndJournal中读取配置名SubjectCodeLength3为的值，该值为字符串。默认为"2"
            SubjectCodeLength3.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeLength3", "2");
            //从父节点BalanceAndJournal中读取配置名SubjectCodeLength4为的值，该值为字符串。默认为"2"
            SubjectCodeLength4.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeLength4", "2");
            //从父节点BalanceAndJournal中读取配置名SubjectCodeLength5为的值，该值为字符串。默认为"2"
            SubjectCodeLength5.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeLength5", "2");
            //从父节点BalanceAndJournal中读取配置名SubjectCodeLength6为的值，该值为字符串。默认为"2"
            SubjectCodeLength6.Text = clsConfig.ReadConfig<string>("BalanceAndJournal", "SubjectCodeLength6", "2");

            //从父节点BalanceAndJournal中读取配置名为OrderCheckBox的值，该值为布尔值。默认为true
            OrderCheckBox.Checked = clsConfig.ReadConfig<bool>("BalanceAndJournal", "OrderCheckBox", true);
        }

        private void ChangeState(object sender, EventArgs e)
        {
            //实现单选效果
            foreach (var btn in SubjectCodeGroupBox.Controls.OfType<RadioButton>().ToList())
            {
                if (btn.Name != (sender as Control).Name)
                {
                    btn.Checked = false;
                }
            }
            //选项2只读
            if (SubjectCodeButton2.Checked == true)
            {
                SubjectCodeSign.ReadOnly = false;
            }
            else
            {
                SubjectCodeSign.ReadOnly = true;
            }
            //选项3只读
            if (SubjectCodeButton3.Checked == true)
            {
                SubjectCodeLength1.ReadOnly = false;
                SubjectCodeLength2.ReadOnly = false;
                SubjectCodeLength3.ReadOnly = false;
                SubjectCodeLength4.ReadOnly = false;
                SubjectCodeLength5.ReadOnly = false;
                SubjectCodeLength6.ReadOnly = false;
            }
            else
            {
                SubjectCodeLength1.ReadOnly = true;
                SubjectCodeLength2.ReadOnly = true;
                SubjectCodeLength3.ReadOnly = true;
                SubjectCodeLength4.ReadOnly = true;
                SubjectCodeLength5.ReadOnly = true;
                SubjectCodeLength6.ReadOnly = true;
            }
        }

        private void ConfirmBtn_Click(object sender, EventArgs e)
        {
            bool IsValid = true;
            //检查输入字节的长度
            if (SubjectCodeSign.Text.Length != 1 || SubjectCodeLength1.Text.Length != 1 || SubjectCodeLength2.Text.Length != 1 || SubjectCodeLength3.Text.Length != 1 || SubjectCodeLength4.Text.Length != 1 || SubjectCodeLength5.Text.Length != 1 || SubjectCodeLength6.Text.Length != 1)
            {
                MessageBox.Show("编码长度输入有误，目前仅支持1-9，请检查并重新输入");
                IsValid = false;
            }

            //检查输入编码是否为数字
            try
            {
                long n = long.Parse(SubjectCodeLength1.Text);
                n = long.Parse(SubjectCodeLength2.Text);
                n = long.Parse(SubjectCodeLength3.Text);
                n = long.Parse(SubjectCodeLength4.Text);
                n = long.Parse(SubjectCodeLength5.Text);
                n = long.Parse(SubjectCodeLength6.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("科目编码长度应为数字，请检查并重新输入");
                IsValid = false;
            }

            //如果输入信息无误，保存并关闭窗体
            if (IsValid)
            {
                //向我的文档写入配置文件
                string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
                ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

                //写入父节点BalanceAndJournal中配置名SubjectCodeButton1(自动识别编码）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeButton1", SubjectCodeButton1.Checked.ToString());

                //写入父节点BalanceAndJournal中配置名SubjectCodeButton2（使用分隔符拆分）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeButton2", SubjectCodeButton2.Checked.ToString());
                //写入父节点BalanceAndJournal中配置名SubjectCodeSign（分隔符号）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeSign", SubjectCodeSign.Text);

                //写入父节点BalanceAndJournal中配置名SubjectCodeButton3（按编码长度拆分）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeButton3", SubjectCodeButton3.Checked.ToString());
                //写入父节点BalanceAndJournal中配置名SubjectCodeLength1（一级科目长度）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeLength1", SubjectCodeLength1.Text);
                //写入父节点BalanceAndJournal中配置名SubjectCodeLength2（二级科目长度）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeLength2", SubjectCodeLength2.Text);
                //写入父节点BalanceAndJournal中配置名SubjectCodeLength3（三级科目长度）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeLength3", SubjectCodeLength3.Text);
                //写入父节点BalanceAndJournal中配置名SubjectCodeLength4（四级科目长度）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeLength4", SubjectCodeLength4.Text);
                //写入父节点BalanceAndJournal中配置名SubjectCodeLength5（五级科目长度）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeLength5", SubjectCodeLength5.Text);
                //写入父节点BalanceAndJournal中配置名SubjectCodeLength6（六级科目长度）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "SubjectCodeLength6", SubjectCodeLength6.Text);

                //写入父节点BalanceAndJournal中配置名OrderCheckBox(按科目编码排序）的配置项
                clsConfig.WriteConfig("BalanceAndJournal", "OrderCheckBox", OrderCheckBox.Checked.ToString());
            
                //关闭窗体
                this.Close();
            }
        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            //关闭窗体
            this.Close();
        }
    }
}
