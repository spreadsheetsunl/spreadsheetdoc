using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel.Extensions;
using System.Collections;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Xml;
using XLParser;
using Irony.Parsing;
using System.IO;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Reflection;



namespace FirstExcelAddin
{
    public partial class Ribbon1
    {
        private static String inputType = "INPUT";
        private static String outputType = "OUTPUT";
        private static String columnType = "COLUMN";
        private static String rowType = "ROW";
        private static String worksheetType = "WORKSHEET";
        private static String cellFormula = "CELL_FORMULA";
        private static String cellValue = "CELL_VALUE";
        private static String spreadsheetType = "SPREADSHEET";
        private static int numberOfDictionaries = 5;
        private static String xsltPath = AppDomain.CurrentDomain.BaseDirectory + "XSLTFile.xslt";

        private Object spreadsheetForm = null;
        private bool rangeIsClicked = false;
        private bool cellIsClicked = false;

        Dictionary<string, List<String>> rangeCellList = new Dictionary<string, List<String>>();
        Dictionary<string, List<String>> cellArgumentsList = new Dictionary<string, List<string>>();
        Dictionary<Excel.Worksheet, String> worksheetList = new Dictionary<Excel.Worksheet, String>();

        //posiçao 0 -> inputs
        //posiçao 1 -> outputs
        //posição 2 ->  worksheet, row, column
        //posição 3 -> ranges
        //posição 4 -> cells
        Dictionary<Excel.Worksheet, List<Dictionary<string, object>>> worksheetDictionary = new Dictionary<Excel.Worksheet, List<Dictionary<string, object>>>();

        //Classe responsavel por criar e gravar o xml (Exportar ficheiro xml)
        SaveXMLClass xml;
        //Classe responsavel por fazer load do xml, ou seja, ler de um ficheiro e colocar nas estruturas. (Importar ficheiro XML)
        LoadXMLClass loadxml;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            xml = new SaveXMLClass();
            loadxml = new LoadXMLClass();
            xml.initializeXML();
        }

        private void Spreadsheet_Comment_Click(object sender, RibbonControlEventArgs e)
        {
            String sheetName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            cellIsClicked = false;
            openSpreadsheetDocumentationForm(sheetName);
        }

        private void Worksheet_Comment_Click(object sender, RibbonControlEventArgs e)
        {
            cellIsClicked = false;
            openWorksheetGeneralForm(getWorksheet());
        }

        private void Cell_Comment_Click(object sender, RibbonControlEventArgs e)
        {
            cellIsClicked = true;

            Excel.Worksheet worksheet = getWorksheet();
            Excel.Range activeCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;

            int row = activeCell.Row;
            int column = activeCell.Column;

            Boolean hasFormula = activeCell.HasFormula;
            String formula = activeCell.Formula;

            //quando existe formula
            if (hasFormula)
            {
                if (!cellArgumentsList.ContainsKey(getCellFormName()))
                    cellArgumentsList.Add(getCellFormName(), getFormulaArguments(formula));

                openCellForm(getCellFormName());
            }
            //quando não tem fórmula
            else
                openCellValue(getCellFormName());
        }

        private void Row_Comment_Click(object sender, RibbonControlEventArgs e)
        {
            cellIsClicked = false;
            openGeneralForm(getRow().ToString(), rowType);
        }

        private void Column_Comment_Click(object sender, RibbonControlEventArgs e)
        {
            cellIsClicked = false;
            openGeneralForm(getExcelColumnName(getColumn()), columnType);
        }

        private void Range_Comment_Click(object sender, RibbonControlEventArgs e)
        {
            rangeIsClicked = true;
            openRangeForm(firstLastRangeElement(rangeList()));
        }

        private void Input_Click(object sender, RibbonControlEventArgs e)
        {
            rangeIsClicked = false;
            openInputForm(getCellFormName());
        }

        private void Output_Click(object sender, RibbonControlEventArgs e)
        {
            rangeIsClicked = false;
            openOutputForm(getCellFormName());
        }

        private void SSdocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            GeneralForm form = (GeneralForm)getSpreadsheetDocumentationForm();
            generalFormChanges(form, spreadsheetType);
        }

        private void worksheetDocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet worksheet = getWorksheet();
            Dictionary<string, object> generalDictionary = worksheetDictionary[worksheet].ElementAt(2);

            GeneralForm form = (GeneralForm)generalDictionary[worksheet.Name];
            generalFormChanges(form, worksheetType);
        }

        private void cellDocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range activeCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            Boolean hasFormula = activeCell.HasFormula;
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);
            String cell = getCellFormName();

            if (hasFormula)
            {
                CellForm form = (CellForm)cellDictionary[cell];
                form.readOnlyTextBox();
            }
            else
            {
                GeneralForm form = (GeneralForm)cellDictionary[cell];
                generalFormChanges(form, cellValue);
            }
        }

        private void rowDocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            String row = splitCell(getCellFormName(), rowType);
            GeneralForm form = (GeneralForm)generalDictionary[row];
            generalFormChanges(form, rowType);
        }

        private void columnDocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            String column = splitCell(getCellFormName(), columnType);
            GeneralForm form = (GeneralForm)generalDictionary[column];
            generalFormChanges(form, columnType);
        }

        private void rangeDocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            if (getRange().Cells.Count == 1)
                createShowRangeForm(getCellFormName());
            else
                createShowRangeForm(firstLastRangeElement(rangeList()));
        }

        private void showInput_Click(object sender, RibbonControlEventArgs e)
        {
            openShowInputOutputForm("Show input", inputType);
        }

        private void showOutput_Click(object sender, RibbonControlEventArgs e)
        {
            openShowInputOutputForm("Show output", outputType);
        }

        private void webPageDocRibbon_Click(object sender, RibbonControlEventArgs e)
        {
            transformXML();
            System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "output.html");
        }

        private void Import_Click(object sender, RibbonControlEventArgs e)
        {
            XmlNodeList worksheets;
            Dictionary<string, object> generalDictionary;
            Dictionary<string, object> inputDictionary;
            Dictionary<string, object> outputDictionary;
            Dictionary<string, object> rangeDictionary;
            Dictionary<string, object> cellDictionary;

            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = ".";
            openFileDialog.Filter = "Xml files (*.xml)|*.xml";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                worksheetDictionary = new Dictionary<Excel.Worksheet, List<Dictionary<string, object>>>();
                initializeAllWorksheets();

                try
                {
                    loadxml.readXMLToInvisibleWorksheet(invisibleWorksheet, openFileDialog.FileName);
                    loadxml.initializeXML(openFileDialog.FileName);
                    xml.load(openFileDialog.FileName);

                    spreadsheetForm = loadxml.loadRootNode();
                    worksheets = loadxml.getWorksheets();

                    foreach (XmlNode chNode in worksheets)
                    {
                        String worksheetName = chNode.Attributes[0].Value;
                        Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[worksheetName];
                        generalDictionary = worksheetDictionary[worksheet].ElementAt(2);
                        inputDictionary = worksheetDictionary[worksheet].ElementAt(0);
                        outputDictionary = worksheetDictionary[worksheet].ElementAt(1);
                        rangeDictionary = worksheetDictionary[worksheet].ElementAt(3);
                        cellDictionary = worksheetDictionary[worksheet].ElementAt(4);
                        loadxml.loadWorksheet(generalDictionary);
                        loadxml.loadWorksheetElements(generalDictionary, inputDictionary, outputDictionary, rangeDictionary,
                            cellDictionary, worksheetName);
                    }
                    buttonStates();
                    MessageBox.Show("Import sucessfull.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read xml file structure. Original error: " + ex.Message);
                }
            }
        }
        private void Export_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Title = "Save XML File";
            saveFile.Filter = "XML Files (*.xml) | *.xml";
            saveFile.DefaultExt = "xml";
            
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                xml.saveXML(saveFile.FileName);
                MessageBox.Show("Export sucessfull.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not export xml file. Original error: " + ex.Message);
                }
            }

        }

        //--------------------------------------------- Metodos auxiliares -----------------------------------------------------//

        private void initiliazeNewWorksheet(Excel.Worksheet worksheet)
        {
            List<Dictionary<string, object>> dictList = new List<Dictionary<string, object>>();
            for (int i = 0; i < numberOfDictionaries; i++)
            {
                dictList.Add(new Dictionary<string, object>());
            }
            worksheetDictionary.Add(worksheet, dictList);
            worksheetList.Add(worksheet, worksheet.Name);
        }

        public void initializeAllWorksheets()
        {
            foreach (Excel.Worksheet displayWorksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                List<Dictionary<string, object>> dictList = new List<Dictionary<string, object>>();
                for (int i = 0; i < numberOfDictionaries; i++)
                {
                    dictList.Add(new Dictionary<string, object>());
                }
                if (!worksheetDictionary.ContainsKey(displayWorksheet))
                {
                    worksheetDictionary.Add(displayWorksheet, dictList);
                    if (!worksheetList.ContainsKey(displayWorksheet))
                        worksheetList.Add(displayWorksheet, displayWorksheet.Name);
                    else
                        worksheetList[displayWorksheet] = displayWorksheet.Name;
                }

                else
                {
                    worksheetDictionary[displayWorksheet] = dictList;
                    worksheetList[displayWorksheet] = displayWorksheet.Name;
                }

            }
        }

        private void generalFormChanges(GeneralForm form, String type)
        {
            form.General_DescriptionGFBox.Enabled = false;
            form.clear_buttonGF.Visible = false;
            form.okButtonGF.Visible = false;
            form.cancelButtonGF.Text = "CLOSE";
            if (type.Equals(cellValue))
                form.remove_button_GF.Visible = false;
            form.ShowDialog();
        }

        //devolve a worksheet selecionada
        private Excel.Worksheet getWorksheet()
        {
            return Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
        }

        //devolve a célula selecionada
        public Excel.Range getRange()
        {
            getWorksheet();
            return (Excel.Range)Globals.ThisAddIn.Application.Selection;
        }

        //devolve a linha selecionada
        private int getRow()
        {
            return getRange().Row;
        }

        //devolve a coluna selecionada
        private int getColumn()
        {
            return getRange().Column;
        }

        //cria a caixa para documentar as linhas, colunas, worksheets e spreadsheets
        private void createGeneralForm(String formName, String type)
        {
            var myForm = new GeneralForm();
            myForm.Text = formName;
            if (type.Equals(cellValue))
                myForm.remove_button_GF.Visible = true;
            else
                myForm.remove_button_GF.Visible = false;
            myForm.ShowDialog();
        }

        private void createCellForm(String formName)
        {
            Excel.Range activeCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            String formula = activeCell.Formula;

            int row = activeCell.Row;
            int column = activeCell.Column;
            var cell = new CellForm();

            cell.Text = "Cell " + formName;
            cell.createFormSizes(cellArgumentsList[formName]);
            cell.Output.Text = "Output \n" + getCellType(row, column) + "\n" + getHeader(row, column);
            cell.ShowDialog();
        }

        //cria a caixa onde é possivel visualizar uma lista dos inputs/ouputs
        //e comentar os mesmos (falta passar o input e output list)
        private void createInputOutputForm(String formName, String type)
        {
            var myForm = new InputOutputForm();
            myForm.Text = formName;
            if (type.Equals(inputType))
            {
                myForm.InputListBox.Text = getInputOutputList(inputType);
            }
            else
            {
                myForm.InputOutputList.Text = "Output List";
                myForm.InputListBox.Text = getInputOutputList(outputType);
            }
            myForm.ShowDialog();
        }

        private void openCellValue(String formName)
        {
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);

            if (cellDictionary.ContainsKey(formName))
                loadCellValue(formName);
            else
                createGeneralForm("Cell " + formName, cellValue);
        }

        private void openCellForm(String formName)
        {
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);

            if (cellDictionary.ContainsKey(formName))
                loadCellFormula(formName);
            else
                createCellForm(formName);
        }

        private void loadCellValue(String key)
        {
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);
            GeneralForm form = (GeneralForm)cellDictionary[key];
            form.remove_button_GF.Enabled = true;
            form.remove_button_GF.Visible = true;
            form.General_DescriptionGFBox.Enabled = true;
            form.clear_buttonGF.Visible = true;
            form.okButtonGF.Visible = true;
            form.cancelButtonGF.Text = "CANCEL";
            form.ShowDialog();
        }

        private void loadCellFormula(String key)
        {
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);
            CellForm form = (CellForm)cellDictionary[key];
            List<string> cellList = cellArgumentsList[key];
            int cellListSize = cellList.Count;
            if (cellListSize == 1)
            {
                String cell = cellList[0];
                if (cell.Contains(":"))
                {
                    string[] cellArguments = cell.Split(':');
                    String inference = "";
                    for (int i = 0; i < cellArguments.Length; i++)
                    {
                        String columnLetter = splitCell(cellArguments[i], columnType);
                        int row = Int32.Parse(splitCell(cellArguments[i], rowType));
                        int column = getExcelColumnNumber(columnLetter);
                        if (i == 0)
                            inference = getHeader2(row, column) + " ...\n";
                        else
                            inference += getHeader2(row, column);
                    }

                    if (inference.Equals(" ...\n"))
                        form.inputCell.Text = cell + " (" + getCellType(cell) + ")";
                    else
                        form.inputCell.Text = cell + " (" + getCellType(cell) + ")\n(" + inference + ")";
                }
                else
                {
                    String columnLetter = splitCell(cell, columnType);
                    int row = Int32.Parse(splitCell(cell, rowType));
                    int column = getExcelColumnNumber(columnLetter);
                    form.inputCell.Text = cell + " (" + getCellType(cell) + ")\n" + getHeader(row, column);
                }
            }
            else
            {
                for (int i = 0; i < cellList.Count; i++)
                {
                    String cell = cellList[i];
                    String columnLetter = splitCell(cell, columnType);
                    int row = Int32.Parse(splitCell(cell, rowType));
                    int column = getExcelColumnNumber(columnLetter);
                    form.loadInputLabels(cellList[i], i, row, column);
                }
            }
            form.Output.Text = "Output \n" + getCellType(key) + "\n" + getHeader(getRow(), getColumn());
            form.ControlBox = true;
            form.cancelButtonCell.Visible = true;
            form.writeReadTextBox();
        }

        public void saveCellFormula(String formName, CellForm form)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);
            if (cellDictionary.ContainsKey(formName))
                cellDictionary[formName] = form;
            else
                cellDictionary.Add(formName, form);

            xml.addCellWithFormulaToWorksheet(formName, getWorksheet().Name, form);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);

        }

        public void saveCellValue(String formName, GeneralForm form)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);
            if (cellDictionary.ContainsKey(formName))
                cellDictionary[formName] = form;
            else
                cellDictionary.Add(formName, form);

            xml.addCellWithValueToWorksheet(formName, getWorksheet().Name, form.General_DescriptionGFBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        private void openWorksheetGeneralForm(Excel.Worksheet worksheet)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[worksheet].ElementAt(2);

            if (worksheetList.ContainsKey(worksheet))
            {
                if (generalDictionary.ContainsKey(worksheet.Name))
                    changeAndLoadWorksheetGeneralForm(worksheet);
                else
                    createGeneralForm("Worksheet " + worksheet.Name, "");
            }
            else
            {
                createGeneralForm("Worksheet " + worksheet.Name, "");
            }
        }

        //Abre o form. Se existir faz load senão cria um novo form.
        private void openGeneralForm(String formName, String identifier)
        {
            String name = "";
            //contem comentarios da folha de calculo, da worksheet, da linha e da coluna
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            if (generalDictionary.ContainsKey(formName))
                loadGeneralForm(getGeneralFormName(identifier));
            else
            {
                if (identifier.Equals(columnType))
                    name = "Column " + formName;
                else
                    name = "Row " + formName;
                createGeneralForm(name, "");
            }
        }

        //Guarda no dicionario o form com a documentação
        public void saveGeneralForm(String formName, GeneralForm form, String type)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            if (generalDictionary.ContainsKey(formName))
                generalDictionary[formName] = form;
            else
                generalDictionary.Add(formName, form);

            if (type.Equals(worksheetType))
                xml.addWorksheetNodeToRoot(worksheetList[getWorksheet()], form.General_DescriptionGFBox.Text);
            if (type.Equals(columnType))
                xml.addColumnToWorksheet(getGeneralFormName(columnType), getGeneralFormName(worksheetType), form.General_DescriptionGFBox.Text);
            if (type.Equals(rowType))
                xml.addRowToWorksheet(getGeneralFormName(rowType), getGeneralFormName(worksheetType), form.General_DescriptionGFBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        public void saveWorksheetGeneralForm()
        {
            if (!worksheetList.ContainsKey(getWorksheet()))
                worksheetList.Add(getWorksheet(), getWorksheet().Name);
            else
                worksheetList[getWorksheet()] = getWorksheet().Name;
        }

        private void changeAndLoadWorksheetGeneralForm(Excel.Worksheet worksheet)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[worksheet].ElementAt(2);
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            String elementName = worksheetList[worksheet];
            GeneralForm form = (GeneralForm)generalDictionary[elementName];
            form.Text = "Worksheet " + worksheet.Name;
            form.General_DescriptionGFBox.Enabled = true;
            form.clear_buttonGF.Visible = true;
            form.okButtonGF.Visible = true;
            form.cancelButtonGF.Text = "CANCEL";
            generalDictionary.Remove(elementName);
            worksheetList[worksheet] = worksheet.Name;
            generalDictionary.Add(worksheet.Name, form);
            xml.modifyNodeNameAndAttribute(worksheet, elementName, form.General_DescriptionGFBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
            form.ShowDialog();
        }

        //Carrega o form com a documentação
        private void loadGeneralForm(String key)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            GeneralForm form = (GeneralForm)generalDictionary[key];
            form.General_DescriptionGFBox.Enabled = true;
            form.clear_buttonGF.Visible = true;
            form.okButtonGF.Visible = true;
            form.cancelButtonGF.Text = "CANCEL";
            form.ShowDialog();
        }

        private void openShowInputOutputForm(String formName, String identifier)
        {
            //contem comentarios da folha de calculo, da worksheet, da linha e da coluna
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            //input list
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            //output list
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);
            String cellName = getCellFormName();
            String generalDescription = "";
            InputOutputForm form;

            if (inputDictionary.ContainsKey(cellName))
            {
                form = (InputOutputForm)inputDictionary[cellName];
                generalDescription = form.inputOutputDescriptionBox.Text;
            }
            if (outputDictionary.ContainsKey(cellName))
            {
                form = (InputOutputForm)outputDictionary[cellName];
                generalDescription = form.inputOutputDescriptionBox.Text;
            }

            if (generalDictionary.ContainsKey(formName))
                loadShowInputOutputForm(formName, identifier, generalDescription);
            else
                createShowInputOutputForm(formName, identifier, generalDescription);
        }

        private void createShowInputOutputForm(String formName, String type, String generalDescription)
        {
            var myForm = new ShowInputOutput();
            myForm.Text = formName;
            if (type.Equals(inputType))
            {
                myForm.showInOutLabel.Text = "Input List";
                myForm.backgroundInOut.Text = "Mark Input With Color";
                myForm.showInOutBox.Text = getInputOutputList(inputType);
            }
            else
            {
                myForm.showInOutLabel.Text = "Output List";
                myForm.backgroundInOut.Text = "Mark Output With Color";
                myForm.showInOutBox.Text = getInputOutputList(outputType);
            }

            myForm.textBoxShowInput.Text = generalDescription;
            myForm.showInOutBox.Select(0, 0);
            myForm.ShowDialog();
        }

        private void loadShowInputOutputForm(String key, String identifier, String generalDescription)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            ShowInputOutput form = (ShowInputOutput)generalDictionary[key];

            if (identifier.Equals(inputType))
                form.showInOutBox.Text = getInputOutputList(inputType);
            else
                form.showInOutBox.Text = getInputOutputList(outputType);

            form.textBoxShowInput.Text = generalDescription;
            form.ShowDialog();
        }

        public void saveShowInputOutputForm(String formName, ShowInputOutput form)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            if (generalDictionary.ContainsKey(formName))
                generalDictionary[formName] = form;
            else
                generalDictionary.Add(formName, form);
        }

        public String getGeneralFormName(String identifier)
        {
            if (identifier.Equals(worksheetType))
                return getWorksheet().Name;
            else if (identifier.Equals(columnType))
                return getExcelColumnName(getColumn());
            else
                return getRow().ToString();
        }

        //Devolve o nome do form da célula selecionada
        public String getCellFormName()
        {
            Excel.Range activeCell = getRange();

            int row = activeCell.Row;
            int column = activeCell.Column;

            string[] cells = (activeCell.get_AddressLocal(row, column)).Split('$');
            String cellName = cells[1] + cells[2];
            return cellName;
        }

        private void openSpreadsheetDocumentationForm(String formName)
        {
            //contem comentarios da folha de calculo, da worksheet, da linha e da coluna
            if (spreadsheetForm != null)
            {
                GeneralForm form = (GeneralForm)getSpreadsheetDocumentationForm();
                form.General_DescriptionGFBox.Enabled = true;
                form.clear_buttonGF.Visible = true;
                form.okButtonGF.Visible = true;
                form.cancelButtonGF.Text = "CANCEL";
                form.ShowDialog();
            }
            else
            {
                var myForm = new GeneralForm();
                myForm.Text = "Spreadsheet " + formName;
                myForm.ShowDialog();
            }
        }

        //metodo que é chamado na classe do showinputform de maneira a que quando se
        //clicar ok mude a cor do fundo das celulas de input/output
        public void callSetCellBackgroundColor(ShowInputOutput form, String type)
        {
            Dictionary<string, object> dictionary = null;
            if (type.Equals(inputType))
                dictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            else
                dictionary = worksheetDictionary[getWorksheet()].ElementAt(1);

            setCellBackgroundColor(form, dictionary, type);
        }

        //ciclos que mudam a cor de fundo da celula no consoante o estado da checkbox.
        private void setCellBackgroundColor(ShowInputOutput form, Dictionary<string, object> dictionary, String type)
        {
            Excel.Worksheet worksheet = getWorksheet();
            ColorConverter cc = new ColorConverter();
            double inputColor = ColorTranslator.ToOle((Color)cc.ConvertFromString("#FF3030"));
            double outputColor = ColorTranslator.ToOle((Color)cc.ConvertFromString("#00E5EE"));

            if (form.backgroundInOut.Checked)
            {
                String firstElement = dictionary.ElementAt(0).Key;
                if (!checkCellBackgroundColor(firstElement))
                {

                    System.Diagnostics.Debug.WriteLine("Vou pintar");
                    foreach (KeyValuePair<string, object> pair in dictionary)
                    {
                        String cell = pair.Key;
                        if (type.Equals(inputType))
                            worksheet.get_Range(cell).Interior.Color = inputColor;
                        else
                            worksheet.get_Range(cell).Interior.Color = outputColor;
                    }
                }
            }
            else
            {
                foreach (KeyValuePair<string, object> pair in dictionary)
                {
                    String cell = pair.Key;
                    var cellColor = worksheet.get_Range(cell).Interior.Color;
                    //Só mete a branco se a cor for a do input ou output
                    if (cellColor == inputColor || cellColor == outputColor)
                        worksheet.get_Range(cell).Interior.ColorIndex = 0;
                }
            }
        }

        private Object getSpreadsheetDocumentationForm()
        {
            return spreadsheetForm;
        }

        public void setSpreadsheetDocumentationForm(GeneralForm form)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            spreadsheetForm = form;
            String sheetName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            xml.addRootNode(sheetName, form.General_DescriptionGFBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        private void openRangeForm(String formName)
        {
            //contem comentarios do range
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);
            if (generalDictionary.ContainsKey(formName))
                loadRangeForm(formName);
            else
                createRangeForm(formName);
        }

        public void saveRangeForm(String formName, InputOutputForm form)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);
            if (generalDictionary.ContainsKey(formName))
                generalDictionary[formName] = form;
            else
                generalDictionary.Add(formName, form);

            xml.addRangeToWorksheet(formName, getWorksheet().Name, form.inputOutputDescriptionBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        private void loadRangeForm(String key)
        {
            Dictionary<string, object> rangeDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);
            InputOutputForm range = (InputOutputForm)rangeDictionary[key];

            if (rangeDictionary.ContainsKey(getRangeName()))
                range.remove_button.Enabled = true;
            else
                range.remove_button.Enabled = false;
            range.ShowDialog();
        }

        private void createRangeForm(String formName)
        {
            var myForm = new InputOutputForm();
            myForm.Text = "Range " + formName;
            myForm.InputOutputList.Text = "Range Cells";
            myForm.InputListBox.Text = getRangeList(formName);
            myForm.createCheckbox();
            myForm.ShowDialog();
        }

        private void createShowRangeForm(String formName)
        {
            Dictionary<string, object> rangeDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);
            var myForm = new ShowRangeForm();

            if (!rangeDictionary.ContainsKey(formName))
                myForm.createFormSizes();
            else
            {
                InputOutputForm rangeForm = (InputOutputForm)rangeDictionary[formName];
                myForm.Text = "Range " + formName;
                myForm.rangeListShowRange.Text = dictionaryRanges();
                myForm.rangeListShowRange.Select(0, 0);
                myForm.rangeCellListBox.Text = getRangeList(formName);
                myForm.rangeCellListBox.Select(0, 0);
                myForm.GD_TextBox.Text = rangeForm.inputOutputDescriptionBox.Text;
                myForm.ShowDialog();
            }
        }

        public bool Range_Comment_IsClicked()
        {
            return rangeIsClicked;
        }

        public bool cell_Comment_IsClicked()
        {
            return cellIsClicked;
        }

        //Ao selecionar uma celula e carregar no botao ouput
        //se já existir um form dessa celula faz-se o carregamento
        //dos dados, caso contrario, cria-se um novo form
        private void openOutputForm(String outputFormName)
        {
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);
            if (outputDictionary.ContainsKey(outputFormName))
                loadOutputForm(outputFormName);
            else
            {
                String formName = "Output " + outputFormName;
                createInputOutputForm(formName, outputType);
            }
        }

        //Ao selecionar uma celula e carregar no botao input
        //se já existir um form dessa celula faz-se o carregamento
        //dos dados, caso contrario, cria-se um novo form
        private void openInputForm(String inputFormName)
        {
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            if (inputDictionary.ContainsKey(inputFormName))
                loadInputForm(inputFormName);
            else
            {
                String formName = "Input " + inputFormName;
                createInputOutputForm(formName, inputType);
            }
        }


        //Grava dos dados do form do input
        public void saveInput(String formName, InputOutputForm form)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            if (inputDictionary.ContainsKey(formName))
                inputDictionary[formName] = form;
            else
                inputDictionary.Add(formName, form);

            xml.addInputToWorksheet(formName, getWorksheet().Name, form.inputOutputDescriptionBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        //Grava dos dados do form do ouput
        public void saveOutput(String formName, InputOutputForm form)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);
            if (outputDictionary.ContainsKey(formName))
                outputDictionary[formName] = form;
            else
                outputDictionary.Add(formName, form);

            xml.addOutputToWorksheet(formName, getWorksheet().Name, form.inputOutputDescriptionBox.Text);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        //Carrega os dados da celula de output selecionada
        private void loadOutputForm(String key)
        {
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);
            InputOutputForm output = (InputOutputForm)outputDictionary[key];
            output.InputListBox.Text = getInputOutputList(outputType);
            if (outputDictionary.ContainsKey(key))
                output.remove_button.Enabled = true;
            else
                output.remove_button.Enabled = false;
            output.ShowDialog();
        }

        //Carrega os dados da celula de input selecionada
        private void loadInputForm(String key)
        {
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            InputOutputForm input = (InputOutputForm)inputDictionary[key];
            input.InputListBox.Text = getInputOutputList(inputType);
            if (inputDictionary.ContainsKey(key))
                input.remove_button.Enabled = true;
            else
                input.remove_button.Enabled = false;
            input.ShowDialog();
        }

        private string getRangeList(String key)
        {
            string rangeList = "";
            foreach (String cell in rangeCellList[key])
            {
                String columnLetter = splitCell(cell, columnType);
                int row = Int32.Parse(splitCell(cell, rowType));
                int column = getExcelColumnNumber(columnLetter);
                rangeList += cell + "(" + getCellType(row, column) + "); ";
            }
            return rangeList;
        }

        //Retorna a lista de inputs ou outputs de cada dicionário
        public string getInputOutputList(String inputOrOuput)
        {
            string showList = "";
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);

            if (inputOrOuput.Equals(inputType))
            {
                foreach (KeyValuePair<string, object> pair in inputDictionary)
                    showList += getListString(pair);
                return showList;
            }
            else
            {
                foreach (KeyValuePair<string, object> pair in outputDictionary)
                    showList += getListString(pair);
                return showList;
            }
        }

        //remove a celula selecionada do respetivo dicionario
        //existem dois dicionarios, o do input e do output
        public void removeInputOutput(String key)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);

            Excel.Worksheet worksheet = getWorksheet();
            ColorConverter cc = new ColorConverter();
            double inputColor = ColorTranslator.ToOle((Color)cc.ConvertFromString("#FF3030"));
            double outputColor = ColorTranslator.ToOle((Color)cc.ConvertFromString("#00E5EE"));

            if (inputDictionary.ContainsKey(key))
                inputDictionary.Remove(key);
            else
                outputDictionary.Remove(key);

            xml.removeFromWorksheet(key, getWorksheet().Name);

            var cellColor = worksheet.get_Range(key).Interior.Color;
            //Só mete a branco se a cor for a do input ou output
            if (cellColor == inputColor || cellColor == outputColor)
                worksheet.get_Range(key).Interior.ColorIndex = 0;

            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        public void removeCell(String key, String type)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);

            cellDictionary.Remove(key);
            if (type.Equals(cellFormula))
                cellArgumentsList.Remove(key);
            xml.removeFromWorksheet(key, getWorksheet().Name);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);
        }

        public void removeRange(String key)
        {
            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            Dictionary<string, object> rangeDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);
            InputOutputForm range = (InputOutputForm)rangeDictionary[key];

            rangeDictionary.Remove(key);
            xml.removeFromWorksheet(key, getWorksheet().Name);
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            xml.saveXMLToWorksheet(invisibleWorksheet);

            if (range.rangeCheckBox.Checked)
                getRange().Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
        }

        public void buttonStates()
        {
            Excel.Worksheet worksheetName = getWorksheet();
            String cellName = getCellFormName();
            String row;
            String column;
            if (cellName.Contains(':'))
            {
                string[] result = cellName.Split(':');
                row = result[0];
                column = result[1];
            }
            else
            {
                row = splitCell(cellName, rowType);
                column = splitCell(cellName, columnType);
            }
            Excel.Range activeCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            GeneralForm form;

            if (worksheetDictionary.ContainsKey(getWorksheet()))
            {
                Dictionary<string, object> inputDictionary = worksheetDictionary[worksheetName].ElementAt(0);
                Dictionary<string, object> outputDictionary = worksheetDictionary[worksheetName].ElementAt(1);
                Dictionary<string, object> generalDictionary = worksheetDictionary[worksheetName].ElementAt(2);
                Dictionary<string, object> rangeDictionary = worksheetDictionary[worksheetName].ElementAt(3);
                Dictionary<string, object> cellDictionary = worksheetDictionary[worksheetName].ElementAt(4);

                if (getRange().Count > 1)
                {
                    this.rangeRibbon.Enabled = true;
                    this.InputRibbon.Enabled = false;
                    this.OutputRibbon.Enabled = false;
                    this.cellRibbon.Enabled = false;
                    this.rowRibbon.Enabled = false;
                    this.columnRibbon.Enabled = false;
                    this.cellDocRibbon.Enabled = false;
                    this.rowDocRibbon.Enabled = false;
                    this.columnDocRibbon.Enabled = false;
                }
                else
                {
                    Boolean hasFormula = activeCell.HasFormula;

                    this.rangeRibbon.Enabled = false;
                    this.rowRibbon.Enabled = true;
                    this.columnRibbon.Enabled = true;

                    if (inputDictionary.ContainsKey(cellName))
                    {
                        this.InputRibbon.Enabled = true;
                        this.OutputRibbon.Enabled = false;
                        this.cellRibbon.Enabled = false;
                    }

                    else if (outputDictionary.ContainsKey(cellName))
                    {
                        this.InputRibbon.Enabled = false;
                        this.OutputRibbon.Enabled = true;
                        this.cellRibbon.Enabled = false;
                    }
                    else if (cellDictionary.ContainsKey(cellName))
                    {
                        this.cellRibbon.Enabled = true;
                        this.InputRibbon.Enabled = false;
                        this.OutputRibbon.Enabled = false;

                        if (hasFormula)
                        {
                            var formCell = cellDictionary[cellName];
                            if (formCell.GetType().Name.Equals("CellForm"))
                            {
                                CellForm cellFormula = (CellForm)formCell;
                                if (cellFormula.checkIfTextBoxIsEmpty())
                                    this.cellDocRibbon.Enabled = false;
                                else
                                    this.cellDocRibbon.Enabled = true;
                            }
                            else
                            {
                                form = (GeneralForm)cellDictionary[getCellFormName()];
                                if (form.General_DescriptionGFBox.Text.Equals(String.Empty))
                                    this.cellDocRibbon.Enabled = false;
                                else
                                    this.cellDocRibbon.Enabled = true;
                            }
                        }
                        else
                        {
                            var formCell = cellDictionary[cellName];
                            if (formCell.GetType().Name.Equals("GeneralForm"))
                            {
                                form = (GeneralForm)formCell;
                                if (form.General_DescriptionGFBox.Text.Equals(String.Empty))
                                    this.cellDocRibbon.Enabled = false;
                                else
                                    this.cellDocRibbon.Enabled = true;
                            }
                            else
                            {
                                CellForm cellFormula = (CellForm)formCell;
                                if (cellFormula.checkIfTextBoxIsEmpty())
                                    this.cellDocRibbon.Enabled = false;
                                else
                                    this.cellDocRibbon.Enabled = true;
                            }
                        }
                    }
                    else
                    {
                        this.InputRibbon.Enabled = true;
                        this.OutputRibbon.Enabled = true;
                        this.cellRibbon.Enabled = true;
                        this.cellDocRibbon.Enabled = false;
                    }

                    form = (GeneralForm)getSpreadsheetDocumentationForm();
                    if (form == null || form.General_DescriptionGFBox.Text.Equals(String.Empty))
                        this.SSdocRibbon.Enabled = false;
                    else
                        this.SSdocRibbon.Enabled = true;

                    if (generalDictionary.ContainsKey(worksheetName.Name))
                    {
                        form = (GeneralForm)generalDictionary[worksheetName.Name];
                        if (form.General_DescriptionGFBox.Text.Equals(String.Empty))
                            this.worksheetDocRibbon.Enabled = false;
                        else
                            this.worksheetDocRibbon.Enabled = true;
                    }
                    else
                        this.worksheetDocRibbon.Enabled = false;

                    if (generalDictionary.ContainsKey(row))
                    {
                        form = (GeneralForm)generalDictionary[row];
                        if (form.General_DescriptionGFBox.Text.Equals(String.Empty))
                            this.rowDocRibbon.Enabled = false;
                        else
                            this.rowDocRibbon.Enabled = true;
                    }
                    else
                        this.rowDocRibbon.Enabled = false;

                    if (generalDictionary.ContainsKey(column))
                    {
                        form = (GeneralForm)generalDictionary[column];
                        if (form.General_DescriptionGFBox.Text.Equals(String.Empty))
                            this.columnDocRibbon.Enabled = false;
                        else
                            this.columnDocRibbon.Enabled = true;
                    }
                    else
                        this.columnDocRibbon.Enabled = false;

                    if (inputDictionary.Count == 0)
                        this.showInput.Enabled = false;
                    else
                        this.showInput.Enabled = true;

                    if (outputDictionary.Count == 0)
                        this.showOutput.Enabled = false;
                    else
                        this.showOutput.Enabled = true;

                    if (rangeDictionary.Count == 0)
                        this.rangeDocRibbon.Enabled = false;
                    else
                        this.rangeDocRibbon.Enabled = true;
                }
            }
            else
                initiliazeNewWorksheet(getWorksheet());
        }

        public void changeCellBackgroundColor(String cellName, String type, String key)
        {
            ColorConverter cc = new ColorConverter();
            if (getInputOuputCheckboxState(key, cellName))
            {
                if (type.Equals(inputType))
                    getRange().Interior.Color = ColorTranslator.ToOle((Color)cc.ConvertFromString("#FF3030"));
                else
                    getRange().Interior.Color = ColorTranslator.ToOle((Color)cc.ConvertFromString("#00E5EE"));
            }
        }

        public int countRange()
        {
            Dictionary<string, object> rangeDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);

            return rangeDictionary.Count;
        }

        private bool getInputOuputCheckboxState(String key, String cellName)
        {
            Dictionary<string, object> generalDictionary = worksheetDictionary[getWorksheet()].ElementAt(2);
            if (generalDictionary.ContainsKey(key))
            {
                ShowInputOutput form = (ShowInputOutput)generalDictionary[key];
                if (form.backgroundInOut.Checked)
                    return true;
            }
            return false;
        }

        //Constroi a string com o nome da coluna e o respetivo tipo
        private string getListString(KeyValuePair<string, object> pair)
        {
            string cell = pair.Key;
            int row = Int32.Parse(splitCell(cell, rowType));
            string auxColumn = splitCell(cell, columnType);
            int column = getExcelColumnNumber(auxColumn);
            string inputString = cell + " (" + getCellType(row, column) + "); ";

            return inputString;
        }

        public String cellType()
        {
            return getCellType(getRow(), getColumn());
        }

        //Retorna o tipo da celula numa string
        public String getCellType(int row, int column)
        {
            var cellType = getWorksheet().Cells[row, column].Value;
            if (cellType != null)
                return cellType.GetType().Name;
            else
                return null;
        }

        public String getCellType(String cell)
        {
            String rowString = splitCell(cell, rowType);
            String columnString = splitCell(cell, columnType);
            if(columnString.Contains('$'))
            {
                string[] columnAux = columnString.Split('$');
                columnString = columnAux[1];
            }
            int row = Convert.ToInt32(rowString);
            int column = getExcelColumnNumber(columnString);

            return getCellType(row, column);
        }

        //Converte numero da coluna na letra correspondente
        private string getExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
        //Converte letra da coluna no numero correspondente
        public int getExcelColumnNumber(string colAdress)
        {
            int[] digits = new int[colAdress.Length];
            for (int i = 0; i < colAdress.Length; ++i)
            {
                digits[i] = Convert.ToInt32(colAdress[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }

        //Separa a letra do numero da celula
        private string splitCell(String cell, String delimiter)
        {
            string pattern = "(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)";
            string[] result = Regex.Split(cell, pattern, RegexOptions.IgnoreCase);

            if (delimiter.Equals(columnType))
                return result[0];
            else
                return result[1];
        }

        //Identifica se o form escolhido é um form de uma celula de input ou  
        //uma celula de output. Este metodo é chamado noutra classe para permitir
        //identificar em que dicionario gravar quando se clica no botao ok. 
        //Retorna o tipo da celula em maiusculas
        public String getInputOutputType(InputOutputForm form)
        {
            string[] type = form.Text.Split(' ');
            return type[0].ToUpper();
        }

        public String getGeneralType(GeneralForm form)
        {
            string[] type = form.Text.Split(' ');
            System.Diagnostics.Debug.WriteLine("GENERAL TYPE: " + type[0]);
            return type[0].ToUpper();
        }

        public String getShowInputOutputType(ShowInputOutput form)
        {
            string[] type = form.showInOutLabel.Text.Split(' ');
            return type[0].ToUpper();
        }

        //Devolve a string da caixa de texto general description
        private String getGeneralFormDescription(String type)
        {
            GeneralForm form;
            if (type.Equals(spreadsheetType))
            {
                form = (GeneralForm)spreadsheetForm;
                return form.General_DescriptionGFBox.Text;
            }
            return "";
        }

        private List<String> rangeList()
        {
            List<String> rangeList = new List<String>();
            foreach (Excel.Range cellAddress in getRange())
            {
                //dá me os valores no conjuto das celulas selecionadas              
                String address = cellAddress.Address;
                string[] cellArray = address.Split('$');
                String cell = cellArray[1] + cellArray[2];
                rangeList.Add(cell);
            }
            return rangeList;
        }

        private String firstLastRangeElement(List<String> rangeList)
        {
            String firstElement = rangeList[0];
            String lastElement = rangeList[rangeList.Count - 1];
            String range = firstElement + ":" + lastElement;

            if (!rangeCellList.ContainsKey(range))
                rangeCellList.Add(range, rangeList);
            return range;
        }

        public String getRangeName()
        {
            return firstLastRangeElement(rangeList());
        }

        private bool checkCellBackgroundColor(String firstElement)
        {
            ColorConverter cc = new ColorConverter();
            int row = Int32.Parse(splitCell(firstElement, rowType));
            string auxColumn = splitCell(firstElement, columnType);
            int column = getExcelColumnNumber(auxColumn);
            double cellColor = getWorksheet().Cells[row, column].Interior.Color;
            double inputColor = ColorTranslator.ToOle((Color)cc.ConvertFromString("#FF3030"));
            double outputColor = ColorTranslator.ToOle((Color)cc.ConvertFromString("#00E5EE"));

            if (cellColor == inputColor || cellColor == outputColor)
                return true;
            else
                return false;
        }


        public void cellChanged(String newValue, String key, int row, int column, bool hasFormula)
        {
            Dictionary<string, object> inputDictionary = worksheetDictionary[getWorksheet()].ElementAt(0);
            Dictionary<string, object> outputDictionary = worksheetDictionary[getWorksheet()].ElementAt(1);
            Dictionary<string, object> cellDictionary = worksheetDictionary[getWorksheet()].ElementAt(4);
            InputOutputForm inputOutput = null;

            if (inputDictionary.ContainsKey(key))
            {
                inputOutput = (InputOutputForm)inputDictionary[key];
                if (!newValue.Equals(inputOutput.cellType))
                {
                    changeActiveCell(row, column);
                    openInputForm(key);
                }
            }
            else if (outputDictionary.ContainsKey(key))
            {
                inputOutput = (InputOutputForm)outputDictionary[key];
                if (!newValue.Equals(inputOutput.cellType))
                {
                    changeActiveCell(row, column);
                    openOutputForm(key);
                }
            }
            else if (cellDictionary.ContainsKey(key))
            {
                Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);

                if (hasFormula)
                {
                    changeActiveCell(row, column);
                    String formula = getRange().Formula;
                    List<String> newFormulaArguments = getFormulaArguments(formula);
                    var form = cellDictionary[getCellFormName()];


                    if (form.GetType().Name.Equals("CellForm"))
                    {
                        CellForm cellFormula = (CellForm)form;
                        if (!sameArgumentsDifferentOrders(cellFormula, newFormulaArguments))
                        {
                            xml.removeFromWorksheet(key, getWorksheet().Name);
                            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
                            xml.saveXMLToWorksheet(invisibleWorksheet);
                            createCellFormModified(cellFormula, getCellFormName(), formula);
                        }
                    }
                    else
                    {
                        cellDictionary.Remove(getCellFormName());
                        xml.removeFromWorksheet(key, getWorksheet().Name);
                        xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
                        xml.saveXMLToWorksheet(invisibleWorksheet);
                        CellForm cellFormula = new CellForm();
                        if (!cellArgumentsList.ContainsKey(getCellFormName()))
                            cellArgumentsList.Add(getCellFormName(), getFormulaArguments(formula));
                        createCellForm(getCellFormName());
                        if (!cellDictionary.ContainsKey(getCellFormName()))
                            cellDictionary.Add(getCellFormName(), cellFormula);
                    }
                }
                // Acaba aqui a formula

                //Começa o value aqui
                else
                {
                    var form = cellDictionary[key];
                    changeActiveCell(row, column);
                    if (form.GetType().Name.Equals("GeneralForm"))
                    {
                        GeneralForm cellValue = (GeneralForm)form;
                        if (!newValue.Equals(cellValue.cellType))
                        {
                            xml.removeFromWorksheet(key, getWorksheet().Name);
                            openCellValue(key);
                        }
                    }
                    else
                    {
                        cellDictionary.Remove(getCellFormName());
                        xml.removeFromWorksheet(key, getWorksheet().Name);
                        xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
                        xml.saveXMLToWorksheet(invisibleWorksheet);
                        GeneralForm cell = new GeneralForm();
                        createGeneralForm(getCellFormName(), cellValue);
                        if (!cellDictionary.ContainsKey(getCellFormName()))
                            cellDictionary.Add(getCellFormName(), cell);
                    }
                }

            }
            else
            { }
        }

        private void changeActiveCell(int row, int column)
        {
            Excel.Worksheet sht = getWorksheet();
            Excel.Range rng = (Excel.Range)sht.Cells[row, column];
            rng.Select();
        }

        private List<String> getFormulaArguments(String formula)
        {
            List<String> formulaArguments = new List<string>();
            ParseTreeNode treeRoot = XLParser.ExcelFormulaParser.Parse(formula);
            int ignoreNode = 0;

            foreach (ParseTreeNode node in XLParser.ExcelFormulaParser.AllNodesConditional(treeRoot))
            {
                if (XLParser.ExcelFormulaParser.Is(node, GrammarNames.ReferenceFunctionCall))
                {
                    ignoreNode = 0;
                    formulaArguments.Add(inputRange(node));
                    node.SkipToRelevant();
                }
                else
                {
                    if (XLParser.ExcelFormulaParser.Is(node, GrammarNames.TokenCell) && ((ignoreNode != 2 && ignoreNode != 6)))
                        formulaArguments.Add(node.Token.Text);
                    ignoreNode++;
                }

            }
            return formulaArguments;
        }

        private String inputRange(ParseTreeNode node)
        {
            String argument = "";
            foreach (ParseTreeNode cell in XLParser.ExcelFormulaParser.AllNodes(node))
            {
                if (cell.Token != null)
                    argument += cell.Token.Text;
            }
            return argument;
        }

        public String dictionaryRanges()
        {
            Dictionary<string, object> rangeDictionary = worksheetDictionary[getWorksheet()].ElementAt(3);
            String ranges = "";
            foreach (KeyValuePair<string, object> pair in rangeDictionary)
                ranges += pair.Key + "; ";

            return ranges;
        }


        private void createCellFormModified(CellForm cellForm, String formName, String formula)
        {
            Excel.Range activeCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            String generalDescription = "General_DescriptionBox";
            String outputDescription = "OutputBox";

            int row = activeCell.Row;
            int column = activeCell.Column;


            List<String> newFormulaArguments = getFormulaArguments(formula);
            cellArgumentsList[getCellFormName()] = newFormulaArguments;

            CellForm cell = new CellForm();
            cell.Text = "Cell " + formName;
            cell.createFormSizes(cellArgumentsList[formName]);
            for (int i = 0; i < newFormulaArguments.Count; i++)
            {
                if (cellForm.argumentsText.ContainsKey(newFormulaArguments[i]))
                {
                    cell.copyTextBox(newFormulaArguments[i], cellForm.argumentsText[newFormulaArguments[i]]);
                }
            }
            cell.copyTextBox(generalDescription, cellForm.argumentsText[generalDescription]);
            cell.General_DescriptionBox.Select(0, 0);
            cell.copyTextBox(outputDescription, cellForm.argumentsText[outputDescription]);
            cell.remove_buttonCell.Visible = false;
            cell.cancelButtonCell.Visible = false;
            cell.clear_buttonCell.Visible = false;
            cell.ControlBox = false;
            cell.Output.Text = "Output \n" + getCellType(row, column) + "\n" + getHeader(row, column);
            cell.ShowDialog();
        }

        private Boolean sameArgumentsDifferentOrders(CellForm cellForm, List<String> newFormulaArguments)
        {
            for (int i = 0; i < newFormulaArguments.Count; i++)
            {
                if (!cellForm.argumentsText.ContainsKey(newFormulaArguments[i]))
                    return false;
            }
            if (newFormulaArguments.Count + 2 < cellForm.argumentsText.Count)
                return false;

            return true;
        }

        public void importTemporaryXMLFile()
        {
            XmlNodeList worksheets;
            String path = AppDomain.CurrentDomain.BaseDirectory + "TemporaryXMLFileTemp.xml";
            Dictionary<string, object> generalDictionary;
            Dictionary<string, object> inputDictionary;
            Dictionary<string, object> outputDictionary;
            Dictionary<string, object> rangeDictionary;
            Dictionary<string, object> cellDictionary;


            Excel.Worksheet invisibleWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["WorkQWERTY10745TY"]);
            if (invisibleWorksheet.Cells[1, 1].Value != null)
            {
                loadxml.readXMLToSpreadsheet(invisibleWorksheet);
                loadxml.initializeXML(path);
                xml.loadDoc(path);
                spreadsheetForm = loadxml.loadRootNode();
                worksheets = loadxml.getWorksheets();

                foreach (XmlNode chNode in worksheets)
                {
                    String worksheetName = chNode.Attributes[0].Value;
                    if (getWorksheet(worksheetName) != null)
                    {
                        Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[worksheetName];
                        generalDictionary = worksheetDictionary[worksheet].ElementAt(2);
                        inputDictionary = worksheetDictionary[worksheet].ElementAt(0);
                        outputDictionary = worksheetDictionary[worksheet].ElementAt(1);
                        rangeDictionary = worksheetDictionary[worksheet].ElementAt(3);
                        cellDictionary = worksheetDictionary[worksheet].ElementAt(4);
                        loadxml.loadWorksheet(generalDictionary);
                        loadxml.loadWorksheetElements(generalDictionary, inputDictionary, outputDictionary, rangeDictionary,
                            cellDictionary, worksheetName);
                    }
                }
                buttonStates();
                File.Delete(path);
            }
        }

        public void removeInexistentWorksheet()
        {
            List<Excel.Worksheet> elementsToRemove = new List<Excel.Worksheet>();

            Excel.Worksheet invisibleWorksheet = null;
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                if (sheet.Name.Equals("WorkQWERTY10745TY"))
                {
                    invisibleWorksheet = sheet;
                    break;
                }
            }

            foreach (KeyValuePair<Excel.Worksheet, List<Dictionary<string, object>>> displayWorksheet in worksheetDictionary)
            {
                Excel.Worksheet worksheetName = displayWorksheet.Key;
                Excel.Worksheet worksheet = this.getWorksheet(worksheetName);
                if (worksheet == null)
                    elementsToRemove.Add(worksheetName);
            }
            for (int i = 0; i < elementsToRemove.Count; i++)
            {
                worksheetDictionary.Remove(elementsToRemove[i]);
            }
            if (xml.docExist())
            {
                xml.removeInexistentWorksheet(worksheetDictionary);
                xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
                xml.saveXMLToWorksheet(invisibleWorksheet);
            }

        }
        private Excel.Worksheet getWorksheet(Excel._Worksheet worksheetName)
        {
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                if (sheet.Equals(worksheetName))
                    return sheet;
            }
            return null;
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

        public void transformXML()
        {
            xml.saveHiddenFile("TemporaryXMLFileTemp.xml");
            string inputXML = AppDomain.CurrentDomain.BaseDirectory + "TemporaryXMLFileTemp.xml";
            string transformXSL = xsltPath;

            XmlTextReader reader = new XmlTextReader(transformXSL);
            reader.Read();
            reader.Read();

            XslCompiledTransform xslt = new XslCompiledTransform();
            xslt.Load(reader);
            XPathDocument doc = new XPathDocument(inputXML);
            XmlTextWriter writer = new XmlTextWriter(AppDomain.CurrentDomain.BaseDirectory + "output.html", null);

            //Transform the file and send the output to the console.
            xslt.Transform(doc, null, writer);
            writer.Close();
            File.Delete(inputXML);
        }

        private String getRowHeader(int rowNumber, int colNumber)
        {
            int counter;
            String rowHeader = string.Empty;

            for (counter = 1; counter <= colNumber; counter++)
            {
                if (getWorksheet().Cells[rowNumber, counter].Value != null)
                {
                    if ((getCellType(rowNumber, counter) != null && getCellType(rowNumber, counter).Equals("String")))
                    {
                        if (rowHeader.Equals(string.Empty))
                            rowHeader = getWorksheet().Cells[rowNumber, counter].Text;
                        else
                            rowHeader += " " + getWorksheet().Cells[rowNumber, counter].Text;
                    }
                }
                else
                    rowHeader = string.Empty;
            }
            return rowHeader;
        }

        private String getColumnHeader(int rowNumber, int colNumber)
        {
            int counter;
            String columnHeader = string.Empty;

            for (counter = 1; counter <= rowNumber; counter++)
            {
                if (getWorksheet().Cells[counter, colNumber].Value != null)
                {
                    if ((getCellType(counter, colNumber) != null && getCellType(counter, colNumber).Equals("String")))
                    {
                        if (columnHeader.Equals(string.Empty))
                            columnHeader = getWorksheet().Cells[counter, colNumber].Text;
                        else
                            columnHeader += " " + getWorksheet().Cells[counter, colNumber].Text;
                    }
                }
                else
                    columnHeader = string.Empty;
            }
            return columnHeader;
        }

        public String getHeader(int row, int column)
        {
            String columnHeader = getColumnHeader(row, column);
            String rowHeader = getRowHeader(row, column);

            if (columnHeader.Equals("") && rowHeader.Equals(""))
                return "";
            else if (columnHeader.Equals(""))
                return "(" + rowHeader + ")";
            else if (rowHeader.Equals(""))
                return "(" + columnHeader + ")";
            else
                return "(" + getColumnHeader(row, column) + " | " + getRowHeader(row, column) + ")";
        }

        public String getHeader2(int row, int column)
        {
            if (getColumnHeader(row, column).Equals("") && getRowHeader(row, column).Equals(""))
                return "";
            else
                return getColumnHeader(row, column) + " | " + getRowHeader(row, column);
        }
    }
}
