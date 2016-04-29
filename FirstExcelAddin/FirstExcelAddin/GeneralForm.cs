using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FirstExcelAddin
{
    public partial class GeneralForm : Form
    {
        private static String worksheetType = "WORKSHEET";
        private static String columnType = "COLUMN";
        private static String rowType = "ROW";
        private static String cellValue = "CELL_VALUE";

        private String textBoxDescription;
        private bool clearButtonClicked;
        public String cellType;

        public GeneralForm()
        {
            InitializeComponent();
            this.FormClosing += GeneralForm_FormClosing;
        }

        private void GeneralForm_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            textBoxDescription = General_DescriptionGFBox.Text;
            cellType = getRibbon().cellType();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (getRibbon().cell_Comment_IsClicked())
            {
                getRibbon().saveCellValue(getRibbon().getCellFormName(), this);
                getRibbon().InputRibbon.Enabled = false;
                getRibbon().OutputRibbon.Enabled = false;
                if (this.General_DescriptionGFBox.Text.Equals(String.Empty))
                    getRibbon().cellDocRibbon.Enabled = false;
                else
                    getRibbon().cellDocRibbon.Enabled = true;
            }
            else
            {
                if (getRibbon().getGeneralType(this).Equals(worksheetType))
                {
                    getRibbon().saveWorksheetGeneralForm();
                    getRibbon().saveGeneralForm(getRibbon().getGeneralFormName(worksheetType), this, worksheetType);
                    if (this.General_DescriptionGFBox.Text.Equals(String.Empty))
                        getRibbon().worksheetDocRibbon.Enabled = false;
                    else
                        getRibbon().worksheetDocRibbon.Enabled = true;
                }
                else if (getRibbon().getGeneralType(this).Equals(columnType))
                {
                    getRibbon().saveGeneralForm(getRibbon().getGeneralFormName(columnType), this, columnType);
                    if (this.General_DescriptionGFBox.Text.Equals(String.Empty))
                        getRibbon().columnDocRibbon.Enabled = false;
                    else
                        getRibbon().columnDocRibbon.Enabled = true;

                }
                else if (getRibbon().getGeneralType(this).Equals(rowType))
                {
                    getRibbon().saveGeneralForm(getRibbon().getGeneralFormName(rowType), this, rowType);
                    if (this.General_DescriptionGFBox.Text.Equals(String.Empty))
                        getRibbon().rowDocRibbon.Enabled = false;
                    else
                        getRibbon().rowDocRibbon.Enabled = true;
                }
                else
                {
                    getRibbon().setSpreadsheetDocumentationForm(this);

                    if (this.General_DescriptionGFBox.Text.Equals(String.Empty))
                        getRibbon().SSdocRibbon.Enabled = false;
                    else
                        getRibbon().SSdocRibbon.Enabled = true;

                }
            }

            clearButtonClicked = false;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.General_DescriptionGFBox.Undo();
            if (clearButtonClicked)
                this.General_DescriptionGFBox.Text = textBoxDescription;
            clearButtonClicked = false;
        }

        private void remove_button_Click(object sender, EventArgs e)
        {
            getRibbon().removeCell(getRibbon().getCellFormName(), cellValue);
            getRibbon().InputRibbon.Enabled = true;
            getRibbon().OutputRibbon.Enabled = true;
            getRibbon().cellDocRibbon.Enabled = false;
            clearButtonClicked = false;
        }

        private void clear_button_Click(object sender, EventArgs e)
        {
            clearButtonClicked = true;
            General_DescriptionGFBox.Clear();
        }

        private Ribbon1 getRibbon()
        {
            return Globals.Ribbons.Ribbon1;
        }

        private void GeneralForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                this.General_DescriptionGFBox.Undo();

                if (clearButtonClicked)
                    this.General_DescriptionGFBox.Text = textBoxDescription;
            }
        }
    }
}
