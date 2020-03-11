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
            this.TotalBalance = this.Factory.CreateRibbonButton();
            this.BalanceAndJournalSetting = this.Factory.CreateRibbonButton();
            this.CurrentAccount = this.Factory.CreateRibbonMenu();
            this.EditCurrentAccount = this.Factory.CreateRibbonButton();
            this.AgeOfAccount = this.Factory.CreateRibbonButton();
            this.Confirmation = this.Factory.CreateRibbonButton();
            this.ConfirmationWord = this.Factory.CreateRibbonButton();
            this.CurrentAccountSetting = this.Factory.CreateRibbonButton();
            this.CheckBAJ = this.Factory.CreateRibbonCheckBox();
            this.WorkBook = this.Factory.CreateRibbonGroup();
            this.UnionBook = this.Factory.CreateRibbonButton();
            this.SplitBook = this.Factory.CreateRibbonButton();
            this.Exportxlsx = this.Factory.CreateRibbonButton();
            this.WorkSheet = this.Factory.CreateRibbonGroup();
            this.MakeIndex = this.Factory.CreateRibbonSplitButton();
            this.DeleteFirstRow = this.Factory.CreateRibbonButton();
            this.ChangeName = this.Factory.CreateRibbonButton();
            this.SplitSheet = this.Factory.CreateRibbonButton();
            this.UnionSheet = this.Factory.CreateRibbonButton();
            this.Tool = this.Factory.CreateRibbonGroup();
            this.RangeFormat = this.Factory.CreateRibbonMenu();
            this.DateFormate = this.Factory.CreateRibbonButton();
            this.TextFormat = this.Factory.CreateRibbonButton();
            this.NumFormat = this.Factory.CreateRibbonButton();
            this.ToUpper = this.Factory.CreateRibbonButton();
            this.ToLower = this.Factory.CreateRibbonButton();
            this.AutoFillInTheBlanks = this.Factory.CreateRibbonButton();
            this.CompareTwoColumns = this.Factory.CreateRibbonButton();
            this.CheckNum = this.Factory.CreateRibbonButton();
            this.ChangeSign = this.Factory.CreateRibbonButton();
            this.RoundButton = this.Factory.CreateRibbonSplitButton();
            this.RoundSetting = this.Factory.CreateRibbonButton();
            this.NoRound = this.Factory.CreateRibbonButton();
            this.TenThousand = this.Factory.CreateRibbonSplitButton();
            this.NoTenThousand = this.Factory.CreateRibbonButton();
            this.RegText = this.Factory.CreateRibbonButton();
            this.Protect = this.Factory.CreateRibbonGroup();
            this.ProtectMenu = this.Factory.CreateRibbonMenu();
            this.ProtectBook = this.Factory.CreateRibbonButton();
            this.ProtectSheet = this.Factory.CreateRibbonButton();
            this.ProtectRange = this.Factory.CreateRibbonButton();
            this.Unlock = this.Factory.CreateRibbonMenu();
            this.UnlockBook = this.Factory.CreateRibbonButton();
            this.UnlockSheet = this.Factory.CreateRibbonButton();
            this.ProtectSetting = this.Factory.CreateRibbonButton();
            this.JiuQi = this.Factory.CreateRibbonGroup();
            this.EditJiuQi = this.Factory.CreateRibbonButton();
            this.ExportNotes = this.Factory.CreateRibbonButton();
            this.OpenNoteTemplate = this.Factory.CreateRibbonSplitButton();
            this.OpenFloder = this.Factory.CreateRibbonButton();
            this.VersionGroup = this.Factory.CreateRibbonGroup();
            this.VersionInfo = this.Factory.CreateRibbonButton();
            this.GlobalSetting = this.Factory.CreateRibbonMenu();
            this.TableProcessingCheck = this.Factory.CreateRibbonCheckBox();
            this.WorkBookCheck = this.Factory.CreateRibbonCheckBox();
            this.WorkSheetCheck = this.Factory.CreateRibbonCheckBox();
            this.ToolCheck = this.Factory.CreateRibbonCheckBox();
            this.ProtectCheck = this.Factory.CreateRibbonCheckBox();
            this.JiuQiCheck = this.Factory.CreateRibbonCheckBox();
            this.HertZTab.SuspendLayout();
            this.TableProcessing.SuspendLayout();
            this.WorkBook.SuspendLayout();
            this.WorkSheet.SuspendLayout();
            this.Tool.SuspendLayout();
            this.Protect.SuspendLayout();
            this.JiuQi.SuspendLayout();
            this.VersionGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // HertZTab
            // 
            this.HertZTab.Groups.Add(this.TableProcessing);
            this.HertZTab.Groups.Add(this.WorkBook);
            this.HertZTab.Groups.Add(this.WorkSheet);
            this.HertZTab.Groups.Add(this.Tool);
            this.HertZTab.Groups.Add(this.Protect);
            this.HertZTab.Groups.Add(this.JiuQi);
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
            // TotalBalance
            // 
            this.TotalBalance.Label = "汇总余额表";
            this.TotalBalance.Name = "TotalBalance";
            this.TotalBalance.OfficeImageId = "DesignXml";
            this.TotalBalance.ScreenTip = "从末级科目汇总至一级科目";
            this.TotalBalance.ShowImage = true;
            this.TotalBalance.SuperTip = "同时规范格式";
            this.TotalBalance.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TotalBalance_Click);
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
            this.BalanceAndJournalSetting.Visible = false;
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
            // WorkBook
            // 
            this.WorkBook.Items.Add(this.UnionBook);
            this.WorkBook.Items.Add(this.SplitBook);
            this.WorkBook.Items.Add(this.Exportxlsx);
            this.WorkBook.Label = "工作簿";
            this.WorkBook.Name = "WorkBook";
            // 
            // UnionBook
            // 
            this.UnionBook.Label = "汇总工作簿";
            this.UnionBook.Name = "UnionBook";
            this.UnionBook.OfficeImageId = "ImportExcel";
            this.UnionBook.ScreenTip = "汇总一个文件夹中所有的Excel工作簿";
            this.UnionBook.ShowImage = true;
            this.UnionBook.SuperTip = "要求表头一致，仅汇总活动表一张表格";
            this.UnionBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnionBook_Click);
            // 
            // SplitBook
            // 
            this.SplitBook.Label = "拆分工作簿";
            this.SplitBook.Name = "SplitBook";
            this.SplitBook.OfficeImageId = "CopyToFolder";
            this.SplitBook.ScreenTip = "将当前工作簿中的每一个工作表都拆分为单独的工作簿并保存";
            this.SplitBook.ShowImage = true;
            this.SplitBook.SuperTip = "不拆分隐藏工作表";
            this.SplitBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SplitBook_Click);
            // 
            // Exportxlsx
            // 
            this.Exportxlsx.Label = "另存为xlsx";
            this.Exportxlsx.Name = "Exportxlsx";
            this.Exportxlsx.OfficeImageId = "ExportExcel";
            this.Exportxlsx.ScreenTip = "将xls文件另存为xlsx格式并删除原文件";
            this.Exportxlsx.ShowImage = true;
            this.Exportxlsx.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Exportxlsx_Click);
            // 
            // WorkSheet
            // 
            this.WorkSheet.Items.Add(this.MakeIndex);
            this.WorkSheet.Items.Add(this.ChangeName);
            this.WorkSheet.Items.Add(this.SplitSheet);
            this.WorkSheet.Items.Add(this.UnionSheet);
            this.WorkSheet.Label = "工作表";
            this.WorkSheet.Name = "WorkSheet";
            // 
            // MakeIndex
            // 
            this.MakeIndex.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MakeIndex.Items.Add(this.DeleteFirstRow);
            this.MakeIndex.Label = "生成索引";
            this.MakeIndex.Name = "MakeIndex";
            this.MakeIndex.OfficeImageId = "DatabaseMoveToSharePoint";
            this.MakeIndex.ScreenTip = "生成索引表，默认生成相应链接和返回链接";
            this.MakeIndex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MakeIndex_Click);
            // 
            // DeleteFirstRow
            // 
            this.DeleteFirstRow.Label = "删除首行";
            this.DeleteFirstRow.Name = "DeleteFirstRow";
            this.DeleteFirstRow.OfficeImageId = "FrameDelete";
            this.DeleteFirstRow.ScreenTip = "点击删除生成索引时默认添加的首行返回链接";
            this.DeleteFirstRow.ShowImage = true;
            this.DeleteFirstRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteFirstRow_Click);
            // 
            // ChangeName
            // 
            this.ChangeName.Label = "修改表名";
            this.ChangeName.Name = "ChangeName";
            this.ChangeName.OfficeImageId = "TablePropertiesDialog";
            this.ChangeName.ScreenTip = "点击批量修改工作表名称";
            this.ChangeName.ShowImage = true;
            this.ChangeName.SuperTip = "需要先生成索引表，并在索引列后新建一列存放新表名";
            this.ChangeName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeName_Click);
            // 
            // SplitSheet
            // 
            this.SplitSheet.Label = "按列拆表";
            this.SplitSheet.Name = "SplitSheet";
            this.SplitSheet.OfficeImageId = "TableColumnsInsertRight";
            this.SplitSheet.ScreenTip = "按照所选列将当前表格拆分成多张表格";
            this.SplitSheet.ShowImage = true;
            this.SplitSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SplitSheet_Click);
            // 
            // UnionSheet
            // 
            this.UnionSheet.Label = "多表合并";
            this.UnionSheet.Name = "UnionSheet";
            this.UnionSheet.OfficeImageId = "ReplicationOptionsMenu";
            this.UnionSheet.ScreenTip = "合并当前工作簿中的多张工作表";
            this.UnionSheet.ShowImage = true;
            this.UnionSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnionSheet_Click);
            // 
            // Tool
            // 
            this.Tool.Items.Add(this.RangeFormat);
            this.Tool.Items.Add(this.AutoFillInTheBlanks);
            this.Tool.Items.Add(this.CompareTwoColumns);
            this.Tool.Items.Add(this.CheckNum);
            this.Tool.Items.Add(this.ChangeSign);
            this.Tool.Items.Add(this.RoundButton);
            this.Tool.Items.Add(this.TenThousand);
            this.Tool.Items.Add(this.RegText);
            this.Tool.Label = "实用工具";
            this.Tool.Name = "Tool";
            // 
            // RangeFormat
            // 
            this.RangeFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RangeFormat.Items.Add(this.DateFormate);
            this.RangeFormat.Items.Add(this.TextFormat);
            this.RangeFormat.Items.Add(this.NumFormat);
            this.RangeFormat.Items.Add(this.ToUpper);
            this.RangeFormat.Items.Add(this.ToLower);
            this.RangeFormat.Label = "格式";
            this.RangeFormat.Name = "RangeFormat";
            this.RangeFormat.OfficeImageId = "ControlsGallery";
            this.RangeFormat.ScreenTip = "批量修改选区格式";
            this.RangeFormat.ShowImage = true;
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
            // TextFormat
            // 
            this.TextFormat.Label = "文本格式";
            this.TextFormat.Name = "TextFormat";
            this.TextFormat.OfficeImageId = "FormControlEditBox";
            this.TextFormat.ScreenTip = "将所选单元格强制存储为文本格式";
            this.TextFormat.ShowImage = true;
            this.TextFormat.SuperTip = "加工后不保留公式，仅保留值";
            this.TextFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TextFormat_Click);
            // 
            // NumFormat
            // 
            this.NumFormat.Label = "数字格式";
            this.NumFormat.Name = "NumFormat";
            this.NumFormat.OfficeImageId = "FormattingUnique";
            this.NumFormat.ScreenTip = "将所选单元格转换为数字格式";
            this.NumFormat.ShowImage = true;
            this.NumFormat.SuperTip = "加工后不保留公式，仅保留值";
            this.NumFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NumFormat_Click);
            // 
            // ToUpper
            // 
            this.ToUpper.Label = "字母大写";
            this.ToUpper.Name = "ToUpper";
            this.ToUpper.OfficeImageId = "QuickStylesSets";
            this.ToUpper.ScreenTip = "点击将选区的字母全部转换为大写格式";
            this.ToUpper.ShowImage = true;
            this.ToUpper.SuperTip = "类似upper公式，仅保留值";
            this.ToUpper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToUpper_Click);
            // 
            // ToLower
            // 
            this.ToLower.Label = "字母小写";
            this.ToLower.Name = "ToLower";
            this.ToLower.OfficeImageId = "TextEffectTransformGallery";
            this.ToLower.ScreenTip = "点击将选区的字母全部转换为大写格式";
            this.ToLower.ShowImage = true;
            this.ToLower.SuperTip = "类似Lower公式，仅保留值";
            this.ToLower.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToLower_Click);
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
            // CheckNum
            // 
            this.CheckNum.Label = "检查数字";
            this.CheckNum.Name = "CheckNum";
            this.CheckNum.OfficeImageId = "ConditionalFormattingBottomNItems";
            this.CheckNum.ScreenTip = "检查所选单元格是否都是数字，用黄色标注非数字单元格";
            this.CheckNum.ShowImage = true;
            this.CheckNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckNum_Click);
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
            // RegText
            // 
            this.RegText.Label = "正则匹配";
            this.RegText.Name = "RegText";
            this.RegText.OfficeImageId = "FunctionWizard";
            this.RegText.ScreenTip = "使用正则表达式处理所选单元格值，不保留公式";
            this.RegText.ShowImage = true;
            this.RegText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RegText_Click);
            // 
            // Protect
            // 
            this.Protect.Items.Add(this.ProtectMenu);
            this.Protect.Items.Add(this.Unlock);
            this.Protect.Items.Add(this.ProtectSetting);
            this.Protect.Label = "保护";
            this.Protect.Name = "Protect";
            // 
            // ProtectMenu
            // 
            this.ProtectMenu.Items.Add(this.ProtectBook);
            this.ProtectMenu.Items.Add(this.ProtectSheet);
            this.ProtectMenu.Items.Add(this.ProtectRange);
            this.ProtectMenu.Label = "锁定";
            this.ProtectMenu.Name = "ProtectMenu";
            this.ProtectMenu.OfficeImageId = "Lock";
            this.ProtectMenu.ScreenTip = "锁定当前工作簿、工作表或者所选单元格";
            this.ProtectMenu.ShowImage = true;
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
            // Unlock
            // 
            this.Unlock.Items.Add(this.UnlockBook);
            this.Unlock.Items.Add(this.UnlockSheet);
            this.Unlock.Label = "解锁";
            this.Unlock.Name = "Unlock";
            this.Unlock.OfficeImageId = "AdpPrimaryKey";
            this.Unlock.ScreenTip = "解锁当前工作簿或者工作表";
            this.Unlock.ShowImage = true;
            // 
            // UnlockBook
            // 
            this.UnlockBook.Label = "解锁工作簿";
            this.UnlockBook.Name = "UnlockBook";
            this.UnlockBook.OfficeImageId = "RecordsDeleteRecord";
            this.UnlockBook.ScreenTip = "解除当前工作簿中所有工作表的锁定";
            this.UnlockBook.ShowImage = true;
            this.UnlockBook.SuperTip = "当工作簿中工作表较多时建议先解锁一张表进行测试";
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
            // JiuQi
            // 
            this.JiuQi.Items.Add(this.EditJiuQi);
            this.JiuQi.Items.Add(this.ExportNotes);
            this.JiuQi.Items.Add(this.OpenNoteTemplate);
            this.JiuQi.Label = "久其";
            this.JiuQi.Name = "JiuQi";
            // 
            // EditJiuQi
            // 
            this.EditJiuQi.Label = "加工久其";
            this.EditJiuQi.Name = "EditJiuQi";
            this.EditJiuQi.OfficeImageId = "SharePointListsWorkOffline";
            this.EditJiuQi.ScreenTip = "点击加工久其导出的表格";
            this.EditJiuQi.ShowImage = true;
            this.EditJiuQi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditJiuQi_Click);
            // 
            // ExportNotes
            // 
            this.ExportNotes.Label = "生成附注";
            this.ExportNotes.Name = "ExportNotes";
            this.ExportNotes.OfficeImageId = "ExportWord";
            this.ExportNotes.ScreenTip = "从久其表生成word附注";
            this.ExportNotes.ShowImage = true;
            this.ExportNotes.SuperTip = "生成的附注存放在久其表同一目录下";
            this.ExportNotes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportNotes_Click);
            // 
            // OpenNoteTemplate
            // 
            this.OpenNoteTemplate.Items.Add(this.OpenFloder);
            this.OpenNoteTemplate.Label = "附注模板";
            this.OpenNoteTemplate.Name = "OpenNoteTemplate";
            this.OpenNoteTemplate.OfficeImageId = "FileSaveAsWordDocx";
            this.OpenNoteTemplate.ScreenTip = "点击打开附注模板";
            this.OpenNoteTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenNoteTemplate_Click);
            // 
            // OpenFloder
            // 
            this.OpenFloder.Label = "打开文件夹";
            this.OpenFloder.Name = "OpenFloder";
            this.OpenFloder.OfficeImageId = "FileOpen";
            this.OpenFloder.ScreenTip = "打开模板文件夹";
            this.OpenFloder.ShowImage = true;
            this.OpenFloder.SuperTip = "如果想要恢复默认模板，可以打开文件夹删除模板文件";
            this.OpenFloder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenFloder_Click);
            // 
            // VersionGroup
            // 
            this.VersionGroup.Items.Add(this.VersionInfo);
            this.VersionGroup.Items.Add(this.GlobalSetting);
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
            // GlobalSetting
            // 
            this.GlobalSetting.Items.Add(this.TableProcessingCheck);
            this.GlobalSetting.Items.Add(this.WorkBookCheck);
            this.GlobalSetting.Items.Add(this.WorkSheetCheck);
            this.GlobalSetting.Items.Add(this.ToolCheck);
            this.GlobalSetting.Items.Add(this.ProtectCheck);
            this.GlobalSetting.Items.Add(this.JiuQiCheck);
            this.GlobalSetting.Label = "设置";
            this.GlobalSetting.Name = "GlobalSetting";
            this.GlobalSetting.ScreenTip = "设置选项卡显示情况";
            // 
            // TableProcessingCheck
            // 
            this.TableProcessingCheck.Label = "加工";
            this.TableProcessingCheck.Name = "TableProcessingCheck";
            this.TableProcessingCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableProcessingCheck_Click);
            // 
            // WorkBookCheck
            // 
            this.WorkBookCheck.Label = "工作簿";
            this.WorkBookCheck.Name = "WorkBookCheck";
            this.WorkBookCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WorkBookCheck_Click);
            // 
            // WorkSheetCheck
            // 
            this.WorkSheetCheck.Label = "工作表";
            this.WorkSheetCheck.Name = "WorkSheetCheck";
            this.WorkSheetCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WorkSheetCheck_Click);
            // 
            // ToolCheck
            // 
            this.ToolCheck.Label = "实用工具";
            this.ToolCheck.Name = "ToolCheck";
            this.ToolCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToolCheck_Click);
            // 
            // ProtectCheck
            // 
            this.ProtectCheck.Label = "保护";
            this.ProtectCheck.Name = "ProtectCheck";
            this.ProtectCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProtectCheck_Click);
            // 
            // JiuQiCheck
            // 
            this.JiuQiCheck.Label = "久其";
            this.JiuQiCheck.Name = "JiuQiCheck";
            this.JiuQiCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.JiuQiCheck_Click);
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
            this.WorkBook.ResumeLayout(false);
            this.WorkBook.PerformLayout();
            this.WorkSheet.ResumeLayout(false);
            this.WorkSheet.PerformLayout();
            this.Tool.ResumeLayout(false);
            this.Tool.PerformLayout();
            this.Protect.ResumeLayout(false);
            this.Protect.PerformLayout();
            this.JiuQi.ResumeLayout(false);
            this.JiuQi.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu GlobalSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox TableProcessingCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ToolCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProtectCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup JiuQi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EditJiuQi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportNotes;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox JiuQiCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton OpenNoteTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenFloder;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WorkSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton MakeIndex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteFirstRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SplitSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnionSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WorkBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnionBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SplitBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TextFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NumFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu ProtectMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu Unlock;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu RangeFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ToUpper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ToLower;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RegText;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox WorkBookCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox WorkSheetCheck;
    }

    partial class ThisRibbonCollection
    {
        internal HertZRibbon HertZRibbon
        {
            get { return this.GetRibbon<HertZRibbon>(); }
        }
    }
}
