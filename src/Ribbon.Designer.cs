namespace ExcellMoodleAddin
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            if (disposing && ( components != null ))
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.sabnaaTab = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buildSheetButton = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.sabnaaTab.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            //
            // tab1
            //
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            //
            // sabnaaTab
            //
            this.sabnaaTab.Groups.Add(this.group2);
            this.sabnaaTab.Groups.Add(this.group1);
            this.sabnaaTab.Label = "سابنا";
            this.sabnaaTab.Name = "sabnaaTab";
            //
            // group2
            //
            this.group2.Items.Add(this.buildSheetButton);
            this.group2.Items.Add(this.button2);
            this.group2.Label = "Questions";
            this.group2.Name = "group2";
            //
            // buildSheetButton
            //
            this.buildSheetButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buildSheetButton.Description = "This option creates a sheet for you to enter your questions";
            this.buildSheetButton.Image = global::ExcellMoodleAddin.Properties.Resources.DataSheet;
            this.buildSheetButton.Label = "Create Question Sheet";
            this.buildSheetButton.Name = "buildSheetButton";
            this.buildSheetButton.ShowImage = true;
            this.buildSheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BuildSheetButton_Click);
            //
            // button2
            //
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Description = "This option saves the questions with the gift format";
            this.button2.Image = global::ExcellMoodleAddin.Properties.Resources.SaveAs;
            this.button2.Label = "Save Questions";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button2_Click);
            //
            // group1
            //
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Label = "Guide";
            this.group1.Name = "group1";
            //
            // button3
            //
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Description = "How to use Moodle question import";
            this.button3.Image = global::ExcellMoodleAddin.Properties.Resources.Help;
            this.button3.Label = "Help";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HelpButton_Click);
            //
            // button4
            //
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Description = "About Addin";
            this.button4.Image = global::ExcellMoodleAddin.Properties.Resources.Info;
            this.button4.Label = "About";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);
            //
            // Ribbon
            //
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.sabnaaTab);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.sabnaaTab.ResumeLayout(false);
            this.sabnaaTab.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab sabnaaTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buildSheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
