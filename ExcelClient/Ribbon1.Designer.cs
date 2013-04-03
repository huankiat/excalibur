namespace Excalibur.ExcelClient
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.processClick = this.Factory.CreateRibbonTab();
            this.dataFeed = this.Factory.CreateRibbonGroup();
            this.subButton = this.Factory.CreateRibbonButton();
            this.pubButton = this.Factory.CreateRibbonButton();
            this.repubButton = this.Factory.CreateRibbonButton();
            this.processClick.SuspendLayout();
            this.dataFeed.SuspendLayout();
            // 
            // processClick
            // 
            this.processClick.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.processClick.Groups.Add(this.dataFeed);
            this.processClick.Label = "PROCESSCLICK";
            this.processClick.Name = "processClick";
            // 
            // dataFeed
            // 
            this.dataFeed.Items.Add(this.subButton);
            this.dataFeed.Items.Add(this.pubButton);
            this.dataFeed.Items.Add(this.repubButton);
            this.dataFeed.Label = "Data Feed";
            this.dataFeed.Name = "dataFeed";
            // 
            // subButton
            // 
            this.subButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.subButton.Label = "Subscribe";
            this.subButton.Name = "subButton";
            this.subButton.OfficeImageId = "GetExternalDataFromWeb";
            this.subButton.ShowImage = true;
            // 
            // pubButton
            // 
            this.pubButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.pubButton.Label = "Publish";
            this.pubButton.Name = "pubButton";
            this.pubButton.OfficeImageId = "GoToNewRecord";
            this.pubButton.ShowImage = true;
            // 
            // repubButton
            // 
            this.repubButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.repubButton.Label = "Refresh";
            this.repubButton.Name = "repubButton";
            this.repubButton.OfficeImageId = "DatabaseLinedTableManager";
            this.repubButton.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.processClick);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.processClick.ResumeLayout(false);
            this.processClick.PerformLayout();
            this.dataFeed.ResumeLayout(false);
            this.dataFeed.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab processClick;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup dataFeed;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton subButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pubButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton repubButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
