namespace HertZ_ExcelAddIn
{
    partial class HertZRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HertZRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.HertZTab = this.Factory.CreateRibbonTab();
            this.TableProcessing = this.Factory.CreateRibbonGroup();
            this.BalanceAndJournal = this.Factory.CreateRibbonMenu();
            this.BalanceSheet = this.Factory.CreateRibbonButton();
            this.JournalSheet = this.Factory.CreateRibbonButton();
            this.VoucherCheckList = this.Factory.CreateRibbonButton();
            this.BalanceAndJournalSetting = this.Factory.CreateRibbonButton();
            this.CurrentAccount = this.Factory.CreateRibbonMenu();
            this.EditCurrentAccount = this.Factory.CreateRibbonButton();
            this.AgeOfAccount = this.Factory.CreateRibbonButton();
            this.Confirmation = this.Factory.CreateRibbonButton();
            this.ConfirmationWord = this.Factory.CreateRibbonButton();
            this.CurrentAccountSetting = this.Factory.CreateRibbonButton();
            this.CheckBAJ = this.Factory.CreateRibbonCheckBox();
            this.Tool = this.Factory.CreateRibbonGroup();
            this.AutoFillInTheBlanks = this.Factory.CreateRibbonButton();
            this.CompareTwoColumns = this.Factory.CreateRibbonButton();
            this.Exportxlsx = this.Factory.CreateRibbonButton();
            this.ChangeSign = this.Factory.CreateRibbonButton();
            this.RoundButton = this.Factory.CreateRibbonSplitButton();
            this.RoundSetting = this.Factory.CreateRibbonButton();
            this.NoRound = this.Factory.CreateRibbonButton();
            this.TenThousand = this.Factory.CreateRibbonSplitButton();
            this.NoTenThousand = this.Factory.CreateRibbonButton();
            this.CheckNum = this.Factory.CreateRibbonButton();
            this.DateFormate = this.Factory.CreateRibbonButton();
            this.Protect = this.Factory.CreateRibbonGroup();
            this.ProtectBook = this.Factory.CreateRibbonButton();
            this.ProtectSheet = this.Factory.CreateRibbonButton();
            this.ProtectRange = this.Factory.CreateRibbonButton();
            this.UnlockBook = this.Factory.CreateRibbonButton();
            this.UnlockSheet = this.Factory.CreateRibbonButton();
            this.ProtectSetting = this.Factory.CreateRibbonButton();
            this.VersionGroup = this.Factory.CreateRibbonGroup();
            this.VersionInfo = this.Factory.CreateRibbonButton();
            this.TotalBalance = this.Factory.CreateRibbonButton();
            this.HertZTab.SuspendLayout();
            this.TableProcessing.SuspendLayout();
            this.Tool.SuspendLayout();
            this.Protect.SuspendLayout();
            this.VersionGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // HertZTab
            // 
            this.HertZTab.Groups.Add(this.TableProcessing);
            this.HertZTab.Groups.Add(this.Tool);
            this.HertZTab.Groups.Add(this.Protect);
            this.HertZTab.Groups.Add(this.VersionGroup);
            this.HertZTab.Label = "HertZ";
            this.HertZTab.Name = "HertZTab";
            this.HertZTab.Position = this.Factory.RibbonPosition.AfterOfficeId("TabDeveloper");
            // 
            // TableProcessing
            // 
            this.TableProcessing.Items.Add(this.BalanceAndJournal);
            this.TableProcessing.Items.Add(this.CurrentAccount);
            this.TableProcessing.Items.Add(this.CheckBAJ);
            this.TableProcessing.Label = "加工";
            this.TableProcessing.Name = "TableProcessing";
            // 
            // BalanceAndJournal
            // 
            this.BalanceAndJournal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BalanceAndJournal.Items.Add(this.BalanceSheet);
            this.BalanceAndJournal.Items.Add(this.JournalSheet);
            this.BalanceAndJournal.Items.Add(this.VoucherCheckList);
            this.BalanceAndJournal.Items.Add(this.TotalBalance);
            this.BalanceAndJournal.Items.Add(this.BalanceAndJournalSetting);
            this.BalanceAndJournal.Label = "账表加工";
            this.BalanceAndJournal.Name = "BalanceAndJournal";
            this.BalanceAndJournal.OfficeImageId = "AnimationTransitionGallery";
            this.BalanceAndJournal.ScreenTip = "加工余额表序时账";
            this.BalanceAndJournal.ShowImage = true;
            this.BalanceAndJournal.SuperTip = "单击展开下级标签，鼠标悬停在相关标签显示简介";
            // 
            // BalanceSheet
            // 
            this.BalanceSheet.Description = "large";
            this.BalanceSheet.Label = "加工余额表";
            this.BalanceSheet.Name = "BalanceSheet";
            this.BalanceSheet.OfficeImageId = "OutlineSubtotals";
            this.BalanceSheet.ScreenTip = "单击开始加工余额表，规范余额表格式，便于后续操作";
            this.BalanceSheet.ShowImage = true;
            this.BalanceSheet.SuperTip = "请在余额表中使用该功能，在加工前检查余额表科目层级是否正确";
            this.BalanceSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BalanceSheet_Click);
            // 
            // JournalSheet
            // 
            this.JournalSheet.Description = "large";
            this.JournalSheet.Label = "加工序时账";
            this.JournalSheet.Name = "JournalSheet";
            this.JournalSheet.OfficeImageId = "QueryUpdate";
            this.JournalSheet.ScreenTip = "点击开始加工序时账";
            this.JournalSheet.ShowImage = true;
            this.JournalSheet.SuperTip = "加工序时账之前需先加工余额表";
            this.JournalSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.JournalSheet_Click);
            // 
            // VoucherCheckList
            // 
            this.VoucherCheckList.Description = "large";
            this.VoucherCheckList.Label = "生成抽凭表";
            this.VoucherCheckList.Name = "VoucherCheckList";
            this.VoucherCheckList.OfficeImageId = "CreateQueryFromWizard";
            this.VoucherCheckList.ScreenTip = "从序时账生成抽凭表";
            this.VoucherCheckList.ShowImage = true;
            this.VoucherCheckList.SuperTip = "补充同一凭证的其余发生额";
            this.VoucherCheckList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.VoucherCheckList_Click);
            // 
            // BalanceAndJournalSetting
            // 
            this.BalanceAndJournalSetting.Description = "large";
            this.BalanceAndJournalSetting.Label = "加工设置";
            this.BalanceAndJournalSetting.Name = "BalanceAndJournalSetting";
            this.BalanceAndJournalSetting.OfficeImageId = "AddInManager";
            this.BalanceAndJournalSetting.ScreenTip = "设置加工账相关信息";
            this.BalanceAndJournalSetting.ShowImage = true;
            this.BalanceAndJournalSetting.SuperTip = "如需更改，需在加工前设置好，如科目级次、科目排序等";
            this.BalanceAndJournalSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BalanceAndJournalSetting_Click);
            // 
            // CurrentAccount
            // 
            this.CurrentAccount.Items.Add(this.EditCurrentAccount);
            this.CurrentAccount.Items.Add(this.AgeOfAccount);
            this.CurrentAccount.Items.Add(this.Confirmation);
            this.CurrentAccount.Items.Add(this.ConfirmationWord);
            this.CurrentAccount.Items.Add(this.CurrentAccountSetting);
            this.CurrentAccount.Label = "往来款";
            this.CurrentAccount.Name = "CurrentAccount";
            this.CurrentAccount.OfficeImageId = "OrganizationChartSelectAllConnectors";
            this.CurrentAccount.ScreenTip = "加工往来款及生成函证";
            this.CurrentAccount.ShowImage = true;
            this.CurrentAccount.SuperTip = "点击展开下级标签，点击相关功能即可";
            // 
            // EditCurrentAccount
            // 
            this.EditCurrentAccount.Label = "加工往来款";
            this.EditCurrentAccount.Name = "EditCurrentAccount";
            this.EditCurrentAccount.OfficeImageId = "OrganizationChartSelectAllConnectors";
            this.EditCurrentAccount.ScreenTip = "加工往来款明细表，自动重分类";
            this.EditCurrentAccount.ShowImage = true;
            this.EditCurrentAccount.SuperTip = "要求往来款一级科目需规范，如“应收账款”、“应付账款”等";
            this.EditCurrentAccount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditCurrentAccount_Click);
            // 
            // AgeOfAccount
            // 
            this.AgeOfAccount.Label = "拆分账龄";
            this.AgeOfAccount.Name = "AgeOfAccount";
            this.AgeOfAccount.OfficeImageId = "BusinessFormWizard";
            this.AgeOfAccount.ScreenTip = "按上年账龄拆分本年账龄";
            this.AgeOfAccount.ShowImage = true;
            this.AgeOfAccount.SuperTip = "需要先将上一年度往来款及账龄粘到一张表里，并将这张表放到加工完的本年度往来款表中";
            this.AgeOfAccount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AgeOfAccount_Click);
            // 
            // Confirmation
            // 
            this.Confirmation.Label = "生成发函清单";
            this.Confirmation.Name = "Confirmation";
            this.Confirmation.OfficeImageId = "FieldChooser";
            this.Confirmation.ScreenTip = "点击生成函证列表";
            this.Confirmation.ShowImage = true;
            this.Confirmation.SuperTip = "根据各往来款的抽函情况补充同一公司未抽中的款项，生成发函清单";
            this.Confirmation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Confirmation_Click);
            // 
            // ConfirmationWord
            // 
            this.ConfirmationWord.Label = "生成Word函证";
            this.ConfirmationWord.Name = "ConfirmationWord";
            this.ConfirmationWord.OfficeImageId = "SignatureInsertMenu";
            this.ConfirmationWord.ScreenTip = "点击生成Word函证";
            this.ConfirmationWord.ShowImage = true;
            this.ConfirmationWord.SuperTip = "从模板生成word函证，并存放到指定文件夹";
            this.ConfirmationWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ConfirmationWord_Click);
            // 
            // CurrentAccountSetting
            // 
            this.CurrentAccountSetting.Label = "加工设置";
            this.CurrentAccountSetting.Name = "CurrentAccountSetting";
            this.CurrentAccountSetting.OfficeImageId = "AddInManager";
            this.CurrentAccountSetting.ScreenTip = "设置函证相关信息";
            this.CurrentAccountSetting.ShowImage = true;
            this.CurrentAccountSetting.SuperTip = "如被审计单位名称、回函单位等。";
            this.CurrentAccountSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CurrentAccountSetting_Click);
            // 
            // CheckBAJ
            // 
            this.CheckBAJ.Label = "看 账";
            this.CheckBAJ.Name = "CheckBAJ";
            this.CheckBAJ.ScreenTip = "勾选即可双击看账";
            this.CheckBAJ.SuperTip = "在加工账中勾选可双击看明细及凭证";
            this.CheckBAJ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBAJ_Click);
            // 
            // Tool
            // 
            this.Tool.Items.Add(this.AutoFillInTheBlanks);
            this.Tool.Items.Add(this.CompareTwoColumns);
            this.Tool.Items.Add(this.Exportxlsx);
            this.Tool.Items.Add(this.ChangeSign);
            this.Tool.Items.Add(this.RoundButton);
            this.Tool.Items.Add(this.TenThousand);
            this.Tool.Items.Add(this.CheckNum);
            this.Tool.Items.Add(this.DateFormate);
            this.Tool.Label = "实用工具";
            this.Tool.Name = "Tool";
            // 
            // AutoFillInTheBlanks
            // 
            this.AutoFillInTheBlanks.Label = "填充空行";
            this.AutoFillInTheBlanks.Name = "AutoFillInTheBlanks";
            this.AutoFillInTheBlanks.OfficeImageId = "MergeCellsAcross";
            this.AutoFillInTheBlanks.ScreenTip = "填充所选列的空单元格";
            this.AutoFillInTheBlanks.ShowImage = true;
            this.AutoFillInTheBlanks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AutoFillInTheBlanks_Click);
            // 
            // CompareTwoColumns
            // 
            this.CompareTwoColumns.Label = "对比两列";
            this.CompareTwoColumns.Name = "CompareTwoColumns";
            this.CompareTwoColumns.OfficeImageId = "TableStyleBandedColumns";
            this.CompareTwoColumns.ScreenTip = "对比两列数据";
            this.CompareTwoColumns.ShowImage = true;
            this.CompareTwoColumns.SuperTip = "选择两列进行对比，对两列中不同的数据用黄色标注，要求两列需在同一sheet中";
            this.CompareTwoColumns.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareTwoColumns_Click);
            // 
            // Exportxlsx
            // 
            this.Exportxlsx.Label = "存为xlsx";
            this.Exportxlsx.Name = "Exportxlsx";
            this.Exportxlsx.OfficeImageId = "ExportExcel";
            this.Exportxlsx.ScreenTip = "将xls文件另存为xlsx格式并删除原文件";
            this.Exportxlsx.ShowImage = true;
            this.Exportxlsx.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Exportxlsx_Click);
            // 
            // ChangeSign
            // 
            this.ChangeSign.Label = "正负转换";
            this.ChangeSign.Name = "ChangeSign";
            this.ChangeSign.OfficeImageId = "PivotPlusMinusButtonsShowHide";
            this.ChangeSign.ScreenTip = "改变所选单元格内容的正负号";
            this.ChangeSign.ShowImage = true;
            this.ChangeSign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeSign_Click);
            // 
            // RoundButton
            // 
            this.RoundButton.Items.Add(this.RoundSetting);
            this.RoundButton.Items.Add(this.NoRound);
            this.RoundButton.Label = "小数";
            this.RoundButton.Name = "RoundButton";
            this.RoundButton.OfficeImageId = "R";
            this.RoundButton.ScreenTip = "为所选内容加Round";
            this.RoundButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RoundButton_Click);
            // 
            // RoundSetting
            // 
            this.RoundSetting.Label = "设置";
            this.RoundSetting.Name = "RoundSetting";
            this.RoundSetting.OfficeImageId = "AddInManager";
            this.RoundSetting.ScreenTip = "设置保留的小数位数";
            this.RoundSetting.ShowImage = true;
            this.RoundSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RoundSetting_Click);
            // 
            // NoRound
            // 
            this.NoRound.Label = "去Round";
            this.NoRound.Name = "NoRound";
            this.NoRound.OfficeImageId = "Delete";
            this.NoRound.ScreenTip = "去除所选单元格的Round函数";
            this.NoRound.ShowImage = true;
            this.NoRound.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NoRound_Click);
            // 
            // TenThousand
            // 
            this.TenThousand.Items.Add(this.NoTenThousand);
            this.TenThousand.Label = "万元";
            this.TenThousand.Name = "TenThousand";
            this.TenThousand.OfficeImageId = "T";
            this.TenThousand.ScreenTip = "将所选区域单元格内容除以一万";
            this.TenThousand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TenThousand_Click);
            // 
            // NoTenThousand
            // 
            this.NoTenThousand.Label = "乘一万";
            this.NoTenThousand.Name = "NoTenThousand";
            this.NoTenThousand.OfficeImageId = "Delete";
            this.NoTenThousand.ScreenTip = "去除万元格式的公式";
            this.NoTenThousand.ShowImage = true;
            this.NoTenThousand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NoTenThousand_Click);
            // 
            // CheckNum
            // 
            this.CheckNum.Label = "检查数字";
            this.CheckNum.Name = "CheckNum";
            this.CheckNum.OfficeImageId = "ConditionalFormattingBottomNItems";
            this.CheckNum.ScreenTip = "检查所选单元格是否都是数字，用黄色标注非数字单元格";
            this.CheckNum.ShowImage = true;
            this.CheckNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckNum_Click);
            // 
            // DateFormate
            // 
            this.DateFormate.Label = "日期格式";
            this.DateFormate.Name = "DateFormate";
            this.DateFormate.OfficeImageId = "ProposeNewTime";
            this.DateFormate.ScreenTip = "将所选单元格规范为短日期格式";
            this.DateFormate.ShowImage = true;
            this.DateFormate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DateFormate_Click);
            // 
            // Protect
            // 
            this.Protect.Items.Add(this.ProtectBook);
            this.Protect.Items.Add(this.ProtectSheet);
            this.Protect.Items.Add(this.ProtectRange);
            this.Protect.Items.Add(this.UnlockBook);
            this.Protect.Items.Add(this.UnlockSheet);
            this.Protect.Items.Add(this.ProtectSetting);
            this.Protect.Label = "保护";
            this.Protect.Name = "Protect";
            // 
            // ProtectBook
            // 
            this.ProtectBook.Label = "锁定工作簿";
            this.ProtectBook.Name = "ProtectBook";
            this.ProtectBook.OfficeImageId = "ReviewProtectWorkbook";
            this.ProtectBook.ScreenTip = "锁定当前工作簿中的全部工作表";
            this.ProtectBook.ShowImage = true;
            this.ProtectBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtectBook_Click);
            // 
            // ProtectSheet
            // 
            this.ProtectSheet.Label = "锁定工作表";
            this.ProtectSheet.Name = "ProtectSheet";
            this.ProtectSheet.OfficeImageId = "SheetProtect";
            this.ProtectSheet.ScreenTip = "锁定当前工作表";
            this.ProtectSheet.ShowImage = true;
            this.ProtectSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtectSheet_Click);
            // 
            // ProtectRange
            // 
            this.ProtectRange.Label = "锁定单元格";
            this.ProtectRange.Name = "ProtectRange";
            this.ProtectRange.OfficeImageId = "DatabaseMakeMdeFile";
            this.ProtectRange.ScreenTip = "锁定选中单元格";
            this.ProtectRange.ShowImage = true;
            this.ProtectRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtectRange_Click);
            // 
            // UnlockBook
            // 
            this.UnlockBook.Label = "解锁工作簿";
            this.UnlockBook.Name = "UnlockBook";
            this.UnlockBook.OfficeImageId = "RecordsDeleteRecord";
            this.UnlockBook.ScreenTip = "解除当前工作簿中所有工作表的锁定";
            this.UnlockBook.ShowImage = true;
            this.UnlockBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnlockBook_Click);
            // 
            // UnlockSheet
            // 
            this.UnlockSheet.Label = "解锁工作表";
            this.UnlockSheet.Name = "UnlockSheet";
            this.UnlockSheet.OfficeImageId = "FrameDelete";
            this.UnlockSheet.ScreenTip = "解除当前工作表的锁定";
            this.UnlockSheet.ShowImage = true;
            this.UnlockSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnlockSheet_Click);
            // 
            // ProtectSetting
            // 
            this.ProtectSetting.Label = "密码设置";
            this.ProtectSetting.Name = "ProtectSetting";
            this.ProtectSetting.OfficeImageId = "AddInManager";
            this.ProtectSetting.ScreenTip = "设置默认密码";
            this.ProtectSetting.ShowImage = true;
            this.ProtectSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtectSetting_Click);
            // 
            // VersionGroup
            // 
            this.VersionGroup.Items.Add(this.VersionInfo);
            this.VersionGroup.Label = "更多";
            this.VersionGroup.Name = "VersionGroup";
            // 
            // VersionInfo
            // 
            this.VersionInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.VersionInfo.Image = global::HertZ_ExcelAddIn.Properties.Resources.HertZ_Logo;
            this.VersionInfo.Label = "版本信息";
            this.VersionInfo.Name = "VersionInfo";
            this.VersionInfo.ScreenTip = "点击查看版本信息及设置更新";
            this.VersionInfo.ShowImage = true;
            this.VersionInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.VersionInfo_Click);
            // 
            // TotalBalance
            // 
            this.TotalBalance.Label = "汇总余额表";
            this.TotalBalance.Name = "TotalBalance";
            this.TotalBalance.OfficeImageId = "DesignXml";
            this.TotalBalance.ScreenTip = "从末级科目汇总至一级科目";
            this.TotalBalance.ShowImage = true;
            this.TotalBalance.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TotalBalance_Click);
            // 
            // HertZRibbon
            // 
            this.Name = "HertZRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.HertZTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HertZRibbon_Load);
            this.HertZTab.ResumeLayout(false);
            this.HertZTab.PerformLayout();
            this.TableProcessing.ResumeLayout(false);
            this.TableProcessing.PerformLayout();
            this.Tool.ResumeLayout(false);
            this.Tool.PerformLayout();
            this.Protect.ResumeLayout(false);
            this.Protect.PerformLayout();
            this.VersionGroup.ResumeLayout(false);
            this.VersionGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab HertZTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TableProcessing;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu BalanceAndJournal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton JournalSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VoucherCheckList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BalanceAndJournalSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BalanceSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EditCurrentAccount;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CurrentAccountSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Confirmation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AgeOfAccount;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu CurrentAccount;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup VersionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VersionInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Tool;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CompareTwoColumns;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CheckNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ConfirmationWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox CheckBAJ;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AutoFillInTheBlanks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Exportxlsx;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeSign;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton RoundButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RoundSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NoRound;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DateFormate;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton TenThousand;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NoTenThousand;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Protect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProtectSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProtectBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProtectRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnlockBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnlockSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProtectSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TotalBalance;
    }

    partial class ThisRibbonCollection
    {
        internal HertZRibbon HertZRibbon
        {
            get { return this.GetRibbon<HertZRibbon>(); }
        }
    }
}
