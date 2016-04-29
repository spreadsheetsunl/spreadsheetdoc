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
    public partial class ShowRangeForm : Form
    {
        public ShowRangeForm()
        {
            InitializeComponent();
        }

        public void createFormSizes()
        {
            this.Size = new System.Drawing.Size(368, 219);
            this.Text = "Show Range";
            this.GD_Label.Visible = false;
            this.GD_TextBox.Visible = false;
            this.rangeCellListBox.Visible = false;
            this.rangeCellListLabel.Visible = false;
            this.rangeListShowRange.Text = getRibbon().dictionaryRanges();
            this.rangeListShowRange.Select(0, 0);
            this.ShowDialog();
        }

        private Ribbon1 getRibbon()
        {
            return Globals.Ribbons.Ribbon1;
        }
    }
}
