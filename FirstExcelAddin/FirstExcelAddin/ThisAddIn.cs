using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing;
using System.ComponentModel;
using System.Data;
using System.IO;


namespace FirstExcelAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SheetChange += new Excel.AppEvents_SheetChangeEventHandler(Application_SheetChange);
            this.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Worksheet_Change);
        }

        void Application_SheetSelectionChange(object Sh, Excel.Range target)
        {
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;
            ribbon.buttonStates();
        }

        void Application_SheetChange(object Sh, Excel.Range target)
        {
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;
            int row = target.Row;
            int column = target.Column;
            String newValue = ribbon.getCellType(row, column);
            String address = target.Address;
            string[] cellArray = address.Split('$');
            String key = cellArray[1] + cellArray[2];
            Boolean hasFormula = target.HasFormula;
            ribbon.cellChanged(newValue, key, row, column, hasFormula);
        }
        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;
            if (getWorksheet("WorkQWERTY10745TY") == null)
                createNewWorksheet();
            getWorksheet("WorkQWERTY10745TY").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            //ver se o ficheiro temp para se ler existe
            //se existir ler (fazer o import)     
            ribbon.initializeAllWorksheets();
            ribbon.importTemporaryXMLFile();
        }

        private void createNewWorksheet()
        {
            try { 
            Excel.Worksheet newWorksheet;
            newWorksheet = newWorksheet = (Excel.Worksheet)this.Application.Worksheets.Add();
            newWorksheet.Name = "WorkQWERTY10745TY";         
                }
            catch (Exception ex)
                {
                } 
        }

        private Excel.Worksheet getWorksheet(String name)
        {
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                if (sheet.Name.Equals(name))
                    return sheet;
            }
            return null;
        }

        void Worksheet_Change(object sh)
        {
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;
            ribbon.buttonStates();
            ribbon.removeInexistentWorksheet();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

}
