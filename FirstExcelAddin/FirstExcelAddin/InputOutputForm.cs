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
    public partial class InputOutputForm : Form
    {
        private static String inputType = "INPUT";
        private static String outputType = "OUTPUT";

        public CheckBox rangeCheckBox;
        private bool rangeCheckBoxSavedState;
        private String textBoxDescription;
        private bool clearButtonClicked;
        public String cellType;
        public String formula;

        public InputOutputForm()
        {
            InitializeComponent();
            this.FormClosing += InputOutputForm_FormClosing;
            this.remove_button.Click += new EventHandler(removeClicked);
        }

        private void InputOutputForm_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            textBoxDescription = inputOutputDescriptionBox.Text;
            cellType = getRibbon().cellType();
            if (getRibbon().Range_Comment_IsClicked())
                rangeCheckBoxSavedState = rangeCheckBox.Checked;
        }


        private void okButton_Click(object sender, EventArgs e)
        {
            if (getRibbon().Range_Comment_IsClicked())
            {
                if (rangeCheckBox.Checked)
                    getRibbon().getRange().BorderAround2(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);
                else
                    getRibbon().getRange().Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;

                if (getRibbon().countRange() == 0)
                    getRibbon().rangeDocRibbon.Enabled = true;
                getRibbon().saveRangeForm(getRibbon().getRangeName(), this);
            }
            else
            {
                ColorConverter cc = new ColorConverter();
                if (getRibbon().getInputOutputType(this).Equals(inputType))
                {
                    if (getRibbon().getInputOutputList(inputType).Equals(String.Empty))
                        getRibbon().showInput.Enabled = true;
                    getRibbon().saveInput(formName(), this);
                    getRibbon().OutputRibbon.Enabled = false;
                    getRibbon().cellRibbon.Enabled = false;
                    getRibbon().changeCellBackgroundColor(formName(), inputType, "Show input");
                }
                else
                {
                    if (getRibbon().getInputOutputList(outputType).Equals(String.Empty))
                        getRibbon().showOutput.Enabled = true;
                    getRibbon().saveOutput(formName(), this);
                    getRibbon().InputRibbon.Enabled = false;
                    getRibbon().cellRibbon.Enabled = false;
                    getRibbon().changeCellBackgroundColor(formName(), outputType, "Show output");
                }
            }
            clearButtonClicked = false;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.inputOutputDescriptionBox.Undo();
            if (getRibbon().Range_Comment_IsClicked())
                rangeCheckBox.Checked = rangeCheckBoxSavedState;
            if (clearButtonClicked)
                this.inputOutputDescriptionBox.Text = textBoxDescription;
            clearButtonClicked = false;
        }

        private void clear_button_Click(object sender, EventArgs e)
        {
            clearButtonClicked = true;
            inputOutputDescriptionBox.Clear();
        }

        private Ribbon1 getRibbon()
        {
            return Globals.Ribbons.Ribbon1;
        }

        private String formName()
        {
            return getRibbon().getCellFormName();
        }

        private void InputOutputForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                this.inputOutputDescriptionBox.Undo();
                if (getRibbon().Range_Comment_IsClicked())
                    rangeCheckBox.Checked = rangeCheckBoxSavedState;
                if (clearButtonClicked)
                    this.inputOutputDescriptionBox.Text = textBoxDescription;
            }
            clearButtonClicked = false;

        }

        private void remove_button_Click(object sender, System.EventArgs e)
        {
            if (getRibbon().Range_Comment_IsClicked())
            {
                getRibbon().removeRange(getRibbon().getRangeName());
                if (getRibbon().countRange() == 0)
                    getRibbon().rangeDocRibbon.Enabled = false;
            }
            else
            {
                getRibbon().removeInputOutput(formName());
                getRibbon().InputRibbon.Enabled = true;
                getRibbon().OutputRibbon.Enabled = true;
            }
        }

        public void createCheckbox()
        {
            rangeCheckBox = new CheckBox();
            rangeCheckBox.Text = "Mark range";
            rangeCheckBox.Location = new Point(196, 232);
            this.Controls.Add(rangeCheckBox);
        }

        public void removeClicked(object sender, System.EventArgs e)
        {
            if (getRibbon().getInputOutputList(inputType).Equals(String.Empty))
                getRibbon().showInput.Enabled = false;

            if (getRibbon().getInputOutputList(outputType).Equals(String.Empty))
                getRibbon().showOutput.Enabled = false;
        }
    }
}
