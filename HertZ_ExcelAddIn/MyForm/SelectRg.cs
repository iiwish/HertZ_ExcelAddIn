using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace HertZ_ExcelAddIn
{

    public partial class SelectRg : Form
    {
        private Excel.Application ExcelApp;
        private Excel.Worksheet WST;
        private FunCtion FunC = new FunCtion();
        public string returnValue { get; set; }
        public SelectRg()//string FormCaption,int FormType = 2)
        {
            //FormType 1-row;2-column;3-range,4-sheet
            //label1.Tag = FormType.ToString();
            //label1.Text = FormCaption;
            InitializeComponent();
        }

        private void SelectRg_Load(object sender, EventArgs e)
        {

        }

        private void ConfirmBtn_Click(object sender, EventArgs e)
        {
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            WST = (Excel.Worksheet)ExcelApp.ActiveSheet;
            string Rg;
            returnValue = "false";
            bool RgDim = false;
            //规范输入数据
            try
            {
                Rg = Regex.Split(textBox1.Text.ToString(), "!")[1];
                WST.Range[Rg.Replace("$", "")].Select();
            }
            catch
            {
                MessageBox.Show("输入的数据格式有误,请检查并重新选择");
                return;
            }

            //判断输入的选区长度
            if ((Rg.Length - Rg.Replace("$", "").Length) == 4)
            {
                RgDim = true;
            }

            //输入的格式有4种：Sheet!$A:$A; Sheet!$A$1; Sheet!$1:$1; Sheet!$A$1:$B$1

            switch (int.Parse(label1.Tag.ToString()))
            {
                //返回行
                case 1:
                    if (RgDim)
                    {
                        //如果选中的是一个区域，直接返回行
                        returnValue = Regex.Split(Rg, "$")[2] + Regex.Split(Rg, "$")[4];
                    }
                    else
                    {
                        if(FunC.IsNumber(Regex.Split(Rg, "$")[1].Replace(":", "")))
                        {
                            //如果选中的是正确的行格式，返回行
                            returnValue = Regex.Split(Rg, "$")[1].Replace("$", "");
                        }
                        else
                        {
                            if (FunC.IsNumber(Regex.Split(Rg, "$")[2]))
                            {
                                //如果选中的是单个单元格，返回该单元格的行数
                                returnValue = Regex.Split(Rg, "$")[2] + ":" + Regex.Split(Rg, "$")[2];
                            }
                            else
                            {
                                //如果选中的是列区域，报错
                                MessageBox.Show("需要选择Excel行,请检查并重新选择");
                                return;
                            }
                        }
                    }
                    break;
                //返回列
                case 2:
                    if (RgDim)
                    {
                        //如果选中的是一个区域，直接返回列
                        returnValue = Regex.Split(Rg, "$")[1] + ":" + Regex.Split(Rg, "$")[3];
                    }
                    else
                    {
                        if (FunC.IsLetter(Regex.Split(Rg, "$")[2]))
                        {
                            //如果选中的是正确的列格式，返回列
                            returnValue = Regex.Split(Rg, "$")[1].Replace("$", "");
                        }
                        else
                        {
                            if (FunC.IsLetter(Regex.Split(Rg, "$")[1]))
                            {
                                //如果选中的是单个单元格，返回该单元格的行数
                                returnValue = Regex.Split(Rg, "$")[1] + ":" + Regex.Split(Rg, "$")[1];
                            }
                            else
                            {
                                //如果选中的是行区域，报错
                                MessageBox.Show("需要选择Excel列,请检查并重新选择");
                                return;
                            }
                        }
                    }
                    break;
                case 3:
                    if (RgDim)
                    {
                        //如果选中的是一个区域，直接返回区域
                        returnValue = Regex.Split(Rg, "$")[1].Replace("$", "");
                    }
                    else
                    {
                        //如果选中的是不是区域，报错
                        MessageBox.Show("需要选择Excel区域,请检查并重新选择");
                        return;
                    }
                    break;
                case 4:
                    returnValue = Regex.Split(textBox1.Text.ToString(), "!")[0];
                    break;
            }
            
            DialogResult = DialogResult.OK;
            Close();
        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            returnValue = "false";
            //关闭窗体
            this.Close();
        }
    }
}
