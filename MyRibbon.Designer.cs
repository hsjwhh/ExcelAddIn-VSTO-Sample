namespace ExcelAddIn_VSTO_Sample
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
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
            this.tabToolBox = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btClose = this.Factory.CreateRibbonButton();
            this.spotlightCheckBox = this.Factory.CreateRibbonCheckBox();
            this.tabToolBox.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabToolBox
            // 
            this.tabToolBox.Groups.Add(this.group1);
            this.tabToolBox.Label = "工具箱";
            this.tabToolBox.Name = "tabToolBox";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btClose);
            this.group1.Items.Add(this.spotlightCheckBox);
            this.group1.Label = "自用";
            this.group1.Name = "group1";
            // 
            // btClose
            // 
            this.btClose.Label = "自定义按钮";
            this.btClose.Name = "btClose";
            this.btClose.Visible = false;
            // 
            // spotlightCheckBox
            // 
            this.spotlightCheckBox.Label = "开启聚光灯";
            this.spotlightCheckBox.Name = "spotlightCheckBox";
            this.spotlightCheckBox.Visible = false;
            this.spotlightCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spotlightCheckBox_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabToolBox);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tabToolBox.ResumeLayout(false);
            this.tabToolBox.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabToolBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btClose;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox spotlightCheckBox;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
