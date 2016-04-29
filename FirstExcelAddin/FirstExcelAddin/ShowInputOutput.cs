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
    public partial class ShowInputOutput : Form
    {

        private static String inputType = "INPUT";
        private static String outputType = "OUTPUT";
        private bool checkboxSavedState;

        public ShowInputOutput()
        {
            InitializeComponent();
            this.FormClosing += ShowInputOutput_FormClosing;
        }

        private void ShowInputOutput_Load(object sender, EventArgs e)
        {
            checkboxSavedState = backgroundInOut.Checked;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }

        private void ok_showInOut_Click(object sender, EventArgs e)
        {
            String key = "";
            if (getRibbon().getShowInputOutputType(this).Equals(inputType))
            {
                key = "Show input";
                getRibbon().callSetCellBackgroundColor(this, inputType);
                getRibbon().saveShowInputOutputForm(key, this);
            }
            else
            {
                key = "Show output";
                getRibbon().callSetCellBackgroundColor(this, outputType);
                getRibbon().saveShowInputOutputForm(key, this);
            }
        }

        private void cancelInOut_Click(object sender, EventArgs e)
        {
            backgroundInOut.Checked = checkboxSavedState;
        }

        private Ribbon1 getRibbon()
        {
            return Globals.Ribbons.Ribbon1;
        }

        private void ShowInputOutput_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
                this.backgroundInOut.Checked = checkboxSavedState;
        }
    }
}
