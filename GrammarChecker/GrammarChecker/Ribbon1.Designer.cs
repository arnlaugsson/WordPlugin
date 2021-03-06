﻿namespace GrammarChecker
{
    partial class GrammarRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public GrammarRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1CheckText = this.Factory.CreateRibbonButton();
            this.button1ResetErrors = this.Factory.CreateRibbonButton();
            this.button1ShowErrorList = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1CheckText);
            this.group1.Items.Add(this.button1ShowErrorList);
            this.group1.Items.Add(this.button1ResetErrors);
            this.group1.Label = "Icelandic grammar checker";
            this.group1.Name = "group1";
            // 
            // button1CheckText
            // 
            this.button1CheckText.Label = "Check grammar";
            this.button1CheckText.Name = "button1CheckText";
            this.button1CheckText.ShowImage = true;
            this.button1CheckText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_ClickCheckSelectedText);
            // 
            // button1ResetErrors
            // 
            this.button1ResetErrors.Label = "Clear error list";
            this.button1ResetErrors.Name = "button1ResetErrors";
            this.button1ResetErrors.ShowImage = true;
            this.button1ResetErrors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1ResetErrors_Click);
            // 
            // button1ShowErrorList
            // 
            this.button1ShowErrorList.Label = "Show last error list";
            this.button1ShowErrorList.Name = "button1ShowErrorList";
            this.button1ShowErrorList.ShowImage = true;
            this.button1ShowErrorList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1ShowErrorList_Click);
            // 
            // GrammarRibbon
            // 
            this.Name = "GrammarRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1CheckText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1ResetErrors;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1ShowErrorList;
    }

    partial class ThisRibbonCollection
    {
        internal GrammarRibbon Ribbon1
        {
            get { return this.GetRibbon<GrammarRibbon>(); }
        }
    }
}
