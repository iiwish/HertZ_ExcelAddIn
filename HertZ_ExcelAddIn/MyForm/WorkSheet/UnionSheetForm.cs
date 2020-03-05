using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace HertZ_ExcelAddIn.MyForm.WorkSheet
{
    public partial class UnionSheetForm : Form
    {
        public UnionSheetForm()
        {
            InitializeComponent();
        }

        //引用函数模块
        private readonly FunCtion FunC = new FunCtion();

        private void UnionSheetForm_Load(object sender, EventArgs e)
        {
            Excel.Application ExcelApp = Globals.ThisAddIn.Application;
            Excel.Workbook WBK = ExcelApp.ActiveWorkbook;

            //写入父节点VerInfo中配置名VerNum的配置项
            //clsConfig.WriteConfig("VerInfo", "VerNum", Nverinfo);
            
            //将表名添加为列表项
            foreach(Excel.Worksheet wst in WBK.Worksheets)
            {
                ListBox.Items.Add(wst.Name);
            }

            //从我的文档读取配置
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);

            //从父节点WorkSheet中读取配置名为HeadRows的值，该值为int。默认为"1"
            int HeadRows = clsConfig.ReadConfig<int>("WorkSheet", "HeadRows", 1);

            //设置行数值
            NumUpDown.Value = HeadRows;

        }

        private void Confirm_Click(object sender, EventArgs e)
        {
            
            Excel.Application ExcelApp = Globals.ThisAddIn.Application;
            Excel.Workbook WBK = ExcelApp.ActiveWorkbook;
            Excel.Worksheet WST = ExcelApp.ActiveSheet;
            Excel.Worksheet wst;
            //存储行数到我的文档
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            ClsThisAddinConfig clsConfig = new ClsThisAddinConfig(strPath);
            clsConfig.WriteConfig("WorkSheet", "HeadRows", NumUpDown.Value.ToString());


            //合并表名
            string SheetNameStr = "合并表";
            if (FunC.SheetExist(SheetNameStr))
            {
                for (int i = 1; i < 11; i++)
                {
                    if (!FunC.SheetExist(SheetNameStr + i))
                    {
                        SheetNameStr = SheetNameStr + i;
                        break;
                    }
                }
                if (SheetNameStr == "合并表")
                {
                    MessageBox.Show("已存在多个“合并表”表，请删除或重命名后再试");
                    return;
                }
            }

            ExcelApp.ScreenUpdating = false;

            //添加新表
            WST = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets.Add(ExcelApp.ActiveWorkbook.Worksheets[1]);
            WST.Name = SheetNameStr;

            //被合并表开始行数
            int StartRows = FunC.TI(NumUpDown.Value) + 1;
            int AllColumns ;
            int AllRows ;
            int NewColumns;
            int MaxRows = WST.Rows.Count;

            //合并表行数
            int UnionRows = StartRows;

            //创建数组
            object[,] ORG;

            //读取表头
            if(StartRows != 1) 
            {
                for (int i = 0; i < ListBox.Items.Count; i++)
                {
                    if (ListBox.GetItemChecked(i))
                    {
                        AllColumns = 2;
                        AllRows = 0;
                        wst = WBK.Worksheets[ListBox.GetItemText(ListBox.Items[i])];
                        //获取列数
                        for (int i1 = 1; i1 <= StartRows; i1++)
                        {
                            NewColumns = ((Excel.Range)(wst.Cells[i1, "IV"])).End[Excel.XlDirection.xlToLeft].Column;
                            AllColumns = Math.Max(AllColumns, NewColumns);
                        }

                        //获取行数
                        for (int i1 = 1; i1 <= 10; i1++)
                        {
                            NewColumns = ((Excel.Range)(wst.Cells[wst.Rows.Count, i1])).End[Excel.XlDirection.xlUp].Row;
                            AllRows = Math.Max(AllRows, NewColumns);
                        }

                        if(AllRows < 2) { continue; }

                        //读到数组并存储到合并表
                        ORG = wst.Range["A1:" + FunC.CName(AllColumns) + (StartRows - 1)].Value2;
                        WST.Range["B1:" + FunC.CName(AllColumns+1) + (StartRows - 1)].Value2 = ORG;
                        WST.Range["A1"].Value2 = "工作表名";
                        ORG = null;
                        break;
                    }
                }
            }

            //读取表内容
            for (int i = 0; i < ListBox.Items.Count; i++)
            {
                if (ListBox.GetItemChecked(i))
                {
                    AllColumns = 2;
                    AllRows = 0;
                    wst = WBK.Worksheets[ListBox.GetItemText(ListBox.Items[i])];
                    //获取列数
                    for (int i1 = 1; i1 <= StartRows; i1++)
                    {
                        NewColumns = ((Excel.Range)(wst.Cells[i1, "IV"])).End[Excel.XlDirection.xlToLeft].Column;
                        AllColumns = Math.Max(AllColumns, NewColumns);
                    }

                    //获取行数
                    for (int i1 = 1; i1 <= 10; i1++)
                    {
                        NewColumns = ((Excel.Range)(wst.Cells[wst.Rows.Count, i1])).End[Excel.XlDirection.xlUp].Row;
                        AllRows = Math.Max(AllRows, NewColumns);
                    }

                    if (AllRows + UnionRows >= MaxRows) 
                    {
                        ExcelApp.ScreenUpdating = true;
                        //关闭窗体
                        this.Close();
                        MessageBox.Show("表格行数不够，请检查！");
                        return;
                    }

                    //读到内容并存储到合并表
                    ORG = wst.Range["A" + StartRows + ":" + FunC.CName(AllColumns) + AllRows].Value2;
                    WST.Range[ "B" + UnionRows + ":" + FunC.CName(AllColumns + 1) + (UnionRows+AllRows-StartRows)].Value2 = ORG;
                    WST.Range["A" + UnionRows + ":A" + (UnionRows + AllRows - StartRows)].Value2 = wst.Name;
                    UnionRows = UnionRows + AllRows - StartRows + 1;
                    ORG = null;
                }
            }

            ExcelApp.ScreenUpdating = true ;
            this.Close();
            MessageBox.Show("合并完成！");
        }

        private void SelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ListBox.Items.Count; i++)
            {
                ListBox.SetItemChecked(i, true);
            }
        }

        private void SelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ListBox.Items.Count; i++)
            {
                ListBox.SetItemChecked(i, false);
            }
        }
    }
}
