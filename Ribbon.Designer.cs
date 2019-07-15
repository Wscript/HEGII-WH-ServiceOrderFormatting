namespace HEGII_WH_ServiceOrderFormatting
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupForemanFormatting = this.Factory.CreateRibbonGroup();
            this.buttonForemanFormatting = this.Factory.CreateRibbonButton();
            this.groupCallerFormatting = this.Factory.CreateRibbonGroup();
            this.buttonCallerFormatting = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupForemanFormatting.SuspendLayout();
            this.groupCallerFormatting.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupForemanFormatting);
            this.tab1.Groups.Add(this.groupCallerFormatting);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupForemanFormatting
            // 
            this.groupForemanFormatting.Items.Add(this.buttonForemanFormatting);
            this.groupForemanFormatting.Label = "片区主管";
            this.groupForemanFormatting.Name = "groupForemanFormatting";
            // 
            // buttonForemanFormatting
            // 
            this.buttonForemanFormatting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonForemanFormatting.Image = global::HEGII_WH_ServiceOrderFormatting.Properties.Resources.整理_横;
            this.buttonForemanFormatting.Label = "服务单格式整理";
            this.buttonForemanFormatting.Name = "buttonForemanFormatting";
            this.buttonForemanFormatting.ShowImage = true;
            this.buttonForemanFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonForemanFormatting_Click);
            // 
            // groupCallerFormatting
            // 
            this.groupCallerFormatting.Items.Add(this.buttonCallerFormatting);
            this.groupCallerFormatting.Label = "回访专员";
            this.groupCallerFormatting.Name = "groupCallerFormatting";
            // 
            // buttonCallerFormatting
            // 
            this.buttonCallerFormatting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonCallerFormatting.Image = global::HEGII_WH_ServiceOrderFormatting.Properties.Resources.整理_竖;
            this.buttonCallerFormatting.Label = "服务单格式整理";
            this.buttonCallerFormatting.Name = "buttonCallerFormatting";
            this.buttonCallerFormatting.ShowImage = true;
            this.buttonCallerFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonCallerFormatting_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupForemanFormatting.ResumeLayout(false);
            this.groupForemanFormatting.PerformLayout();
            this.groupCallerFormatting.ResumeLayout(false);
            this.groupCallerFormatting.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupForemanFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonForemanFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCallerFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCallerFormatting;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
