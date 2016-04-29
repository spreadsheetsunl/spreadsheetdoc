namespace FirstExcelAddin
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
            this.lab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Spreadsheet_Comment = this.Factory.CreateRibbonButton();
            this.Worksheet_Comment = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.cellRibbon = this.Factory.CreateRibbonButton();
            this.rowRibbon = this.Factory.CreateRibbonButton();
            this.columnRibbon = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.rangeRibbon = this.Factory.CreateRibbonButton();
            this.Input_Output_Doc = this.Factory.CreateRibbonGroup();
            this.InputRibbon = this.Factory.CreateRibbonButton();
            this.OutputRibbon = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.SSdocRibbon = this.Factory.CreateRibbonButton();
            this.worksheetDocRibbon = this.Factory.CreateRibbonButton();
            this.cellDocRibbon = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.rowDocRibbon = this.Factory.CreateRibbonButton();
            this.columnDocRibbon = this.Factory.CreateRibbonButton();
            this.rangeDocRibbon = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.showInput = this.Factory.CreateRibbonButton();
            this.showOutput = this.Factory.CreateRibbonButton();
            this.webPageDocRibbon = this.Factory.CreateRibbonButton();
            this.Xml_documentation = this.Factory.CreateRibbonGroup();
            this.Import = this.Factory.CreateRibbonButton();
            this.Export = this.Factory.CreateRibbonButton();
            this.teste = this.Factory.CreateRibbonLabel();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.lab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.Input_Output_Doc.SuspendLayout();
            this.group4.SuspendLayout();
            this.Xml_documentation.SuspendLayout();
            // 
            // lab1
            // 
            this.lab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.lab1.Groups.Add(this.group1);
            this.lab1.Groups.Add(this.group2);
            this.lab1.Groups.Add(this.Input_Output_Doc);
            this.lab1.Groups.Add(this.group4);
            this.lab1.Groups.Add(this.Xml_documentation);
            this.lab1.Label = "Spreadsheet Doc";
            this.lab1.Name = "lab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Spreadsheet_Comment);
            this.group1.Items.Add(this.Worksheet_Comment);
            this.group1.Label = "General Documentation";
            this.group1.Name = "group1";
            // 
            // Spreadsheet_Comment
            // 
            this.Spreadsheet_Comment.Label = "Spreadsheet";
            this.Spreadsheet_Comment.Name = "Spreadsheet_Comment";
            this.Spreadsheet_Comment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Spreadsheet_Comment_Click);
            // 
            // Worksheet_Comment
            // 
            this.Worksheet_Comment.Label = "Worksheet";
            this.Worksheet_Comment.Name = "Worksheet_Comment";
            this.Worksheet_Comment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Worksheet_Comment_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.cellRibbon);
            this.group2.Items.Add(this.rowRibbon);
            this.group2.Items.Add(this.columnRibbon);
            this.group2.Items.Add(this.separator3);
            this.group2.Items.Add(this.rangeRibbon);
            this.group2.Label = "Content Documentation";
            this.group2.Name = "group2";
            // 
            // cellRibbon
            // 
            this.cellRibbon.Label = "Cell";
            this.cellRibbon.Name = "cellRibbon";
            this.cellRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Cell_Comment_Click);
            // 
            // rowRibbon
            // 
            this.rowRibbon.Label = "Row ";
            this.rowRibbon.Name = "rowRibbon";
            this.rowRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Row_Comment_Click);
            // 
            // columnRibbon
            // 
            this.columnRibbon.Label = "Column";
            this.columnRibbon.Name = "columnRibbon";
            this.columnRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Column_Comment_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // rangeRibbon
            // 
            this.rangeRibbon.Enabled = false;
            this.rangeRibbon.Label = "Range";
            this.rangeRibbon.Name = "rangeRibbon";
            this.rangeRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Range_Comment_Click);
            // 
            // Input_Output_Doc
            // 
            this.Input_Output_Doc.Items.Add(this.InputRibbon);
            this.Input_Output_Doc.Items.Add(this.OutputRibbon);
            this.Input_Output_Doc.Label = "Input/Output Documentation";
            this.Input_Output_Doc.Name = "Input_Output_Doc";
            // 
            // InputRibbon
            // 
            this.InputRibbon.Label = "Input";
            this.InputRibbon.Name = "InputRibbon";
            this.InputRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Input_Click);
            // 
            // OutputRibbon
            // 
            this.OutputRibbon.Enabled = false;
            this.OutputRibbon.Label = "Output";
            this.OutputRibbon.Name = "OutputRibbon";
            this.OutputRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Output_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.SSdocRibbon);
            this.group4.Items.Add(this.worksheetDocRibbon);
            this.group4.Items.Add(this.cellDocRibbon);
            this.group4.Items.Add(this.separator1);
            this.group4.Items.Add(this.rowDocRibbon);
            this.group4.Items.Add(this.columnDocRibbon);
            this.group4.Items.Add(this.rangeDocRibbon);
            this.group4.Items.Add(this.separator2);
            this.group4.Items.Add(this.showInput);
            this.group4.Items.Add(this.showOutput);
            this.group4.Items.Add(this.webPageDocRibbon);
            this.group4.Label = "Read Documentation";
            this.group4.Name = "group4";
            // 
            // SSdocRibbon
            // 
            this.SSdocRibbon.Enabled = false;
            this.SSdocRibbon.Label = "Spreadsheet";
            this.SSdocRibbon.Name = "SSdocRibbon";
            this.SSdocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SSdocRibbon_Click);
            // 
            // worksheetDocRibbon
            // 
            this.worksheetDocRibbon.Enabled = false;
            this.worksheetDocRibbon.Label = "Worksheet";
            this.worksheetDocRibbon.Name = "worksheetDocRibbon";
            this.worksheetDocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.worksheetDocRibbon_Click);
            // 
            // cellDocRibbon
            // 
            this.cellDocRibbon.Enabled = false;
            this.cellDocRibbon.Label = "Cell";
            this.cellDocRibbon.Name = "cellDocRibbon";
            this.cellDocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cellDocRibbon_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // rowDocRibbon
            // 
            this.rowDocRibbon.Enabled = false;
            this.rowDocRibbon.Label = "Row";
            this.rowDocRibbon.Name = "rowDocRibbon";
            this.rowDocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rowDocRibbon_Click);
            // 
            // columnDocRibbon
            // 
            this.columnDocRibbon.Enabled = false;
            this.columnDocRibbon.Label = "Column";
            this.columnDocRibbon.Name = "columnDocRibbon";
            this.columnDocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.columnDocRibbon_Click);
            // 
            // rangeDocRibbon
            // 
            this.rangeDocRibbon.Enabled = false;
            this.rangeDocRibbon.Label = "Range";
            this.rangeDocRibbon.Name = "rangeDocRibbon";
            this.rangeDocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rangeDocRibbon_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // showInput
            // 
            this.showInput.Enabled = false;
            this.showInput.Label = "Input";
            this.showInput.Name = "showInput";
            this.showInput.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showInput_Click);
            // 
            // showOutput
            // 
            this.showOutput.Enabled = false;
            this.showOutput.Label = "Ouput";
            this.showOutput.Name = "showOutput";
            this.showOutput.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showOutput_Click);
            // 
            // webPageDocRibbon
            // 
            this.webPageDocRibbon.Label = "Web page";
            this.webPageDocRibbon.Name = "webPageDocRibbon";
            this.webPageDocRibbon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.webPageDocRibbon_Click);
            // 
            // Xml_documentation
            // 
            this.Xml_documentation.Items.Add(this.Import);
            this.Xml_documentation.Items.Add(this.Export);
            this.Xml_documentation.Label = "XML Documentation";
            this.Xml_documentation.Name = "Xml_documentation";
            // 
            // Import
            // 
            this.Import.Label = "Import";
            this.Import.Name = "Import";
            this.Import.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Import_Click);
            // 
            // Export
            // 
            this.Export.Label = "Export";
            this.Export.Name = "Export";
            this.Export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Export_Click);
            // 
            // teste
            // 
            this.teste.Label = "";
            this.teste.Name = "teste";
            // 
            // toggleButton1
            // 
            this.toggleButton1.Label = "";
            this.toggleButton1.Name = "toggleButton1";
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "";
            this.checkBox1.Name = "checkBox1";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.lab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.lab1.ResumeLayout(false);
            this.lab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.Input_Output_Doc.ResumeLayout(false);
            this.Input_Output_Doc.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.Xml_documentation.ResumeLayout(false);
            this.Xml_documentation.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cellRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rowRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton columnRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Spreadsheet_Comment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Worksheet_Comment;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab lab1;
        public Microsoft.Office.Tools.Ribbon.RibbonLabel teste;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Input_Output_Doc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OutputRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SSdocRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rangeRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton webPageDocRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Xml_documentation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Import;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Export;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InputRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showInput;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showOutput;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton worksheetDocRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rowDocRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton columnDocRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cellDocRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rangeDocRibbon;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
