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
            this.HertZTab.SuspendLayout();
            this.TableProcessing.SuspendLayout();
            this.SuspendLayout();
            // 
            // HertZTab
            // 
            this.HertZTab.Groups.Add(this.TableProcessing);
            this.HertZTab.Label = "赫兹";
            this.HertZTab.Name = "HertZTab";
            this.HertZTab.Position = this.Factory.RibbonPosition.AfterOfficeId("TabDeveloper");
            // 
            // TableProcessing
            // 
            this.TableProcessing.Items.Add(this.BalanceAndJournal);
            this.TableProcessing.Label = "加工";
            this.TableProcessing.Name = "TableProcessing";
            // 
            // BalanceAndJournal
            // 
            this.BalanceAndJournal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BalanceAndJournal.Items.Add(this.BalanceSheet);
            this.BalanceAndJournal.Items.Add(this.JournalSheet);
            this.BalanceAndJournal.Items.Add(this.VoucherCheckList);
            this.BalanceAndJournal.Items.Add(this.BalanceAndJournalSetting);
            this.BalanceAndJournal.Label = "账表加工";
            this.BalanceAndJournal.Name = "BalanceAndJournal";
            this.BalanceAndJournal.OfficeImageId = "AnimationTransitionGallery";
            this.BalanceAndJournal.ScreenTip = "加工余额表序时账";
            this.BalanceAndJournal.ShowImage = true;
            this.BalanceAndJournal.SuperTip = "单击展开，选择加工余额表、序时账或更改相关设置";
            // 
            // BalanceSheet
            // 
            this.BalanceSheet.Label = "加工余额表";
            this.BalanceSheet.Name = "BalanceSheet";
            this.BalanceSheet.OfficeImageId = "OutlineSubtotals";
            this.BalanceSheet.ScreenTip = "单击开始加工余额表";
            this.BalanceSheet.ShowImage = true;
            this.BalanceSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BalanceSheet_Click);
            // 
            // JournalSheet
            // 
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
            this.VoucherCheckList.Label = "生成抽凭表";
            this.VoucherCheckList.Name = "VoucherCheckList";
            this.VoucherCheckList.OfficeImageId = "CreateQueryFromWizard";
            this.VoucherCheckList.ScreenTip = "点击生成抽凭表";
            this.VoucherCheckList.ShowImage = true;
            this.VoucherCheckList.SuperTip = "根据抽凭比例自动补全抽凭，可在“设置”中修改默认配置";
            // 
            // BalanceAndJournalSetting
            // 
            this.BalanceAndJournalSetting.Label = "加工设置";
            this.BalanceAndJournalSetting.Name = "BalanceAndJournalSetting";
            this.BalanceAndJournalSetting.OfficeImageId = "AddInManager";
            this.BalanceAndJournalSetting.ShowImage = true;
            this.BalanceAndJournalSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BalanceAndJournalSetting_Click);
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
    }

    partial class ThisRibbonCollection
    {
        internal HertZRibbon HertZRibbon
        {
            get { return this.GetRibbon<HertZRibbon>(); }
        }
    }
}
