using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace FirstExcelAddin
{
    class LoadXMLClass
    {
        XmlDocument doc = new XmlDocument();

        public void initializeXML(String fileName)
        {
            doc.Load(fileName);
        }

        public GeneralForm loadRootNode()
        {
            XmlNode root = doc.DocumentElement;
            GeneralForm spreadsheetForm = new GeneralForm();
            spreadsheetForm.Text = "Spreadsheet " + root.Attributes[0].Value;
            spreadsheetForm.General_DescriptionGFBox.Text = root.Attributes[1].Value;
            spreadsheetForm.General_DescriptionGFBox.Select(0, 0);

            return spreadsheetForm;
        }

        public XmlNodeList getWorksheets()
        {
            return doc.DocumentElement.ChildNodes;
        }

        public void loadWorksheet(Dictionary<string, object> generalDictionary)
        {
            XmlNode root = doc.DocumentElement;
            if (root != null)
            {
                GeneralForm worksheetForm;

                foreach (XmlNode chNode in root.ChildNodes)
                {
                    worksheetForm = new GeneralForm();
                    String worksheetName = chNode.Attributes[0].Value;
                    worksheetForm.Text = "Worksheet " + worksheetName;
                    worksheetForm.General_DescriptionGFBox.Text = chNode.Attributes[1].Value;
                    worksheetForm.General_DescriptionGFBox.Select(0, 0);
                    generalDictionary.Add(worksheetName, worksheetForm);
                }
            }
        }

        public void loadWorksheetElements(Dictionary<string, object> generalDictionary, Dictionary<string, object> inputDictionary,
            Dictionary<string, object> outputDictionary, Dictionary<string, object> rangeDictionary,
            Dictionary<string, object> cellDictionary, String worksheetName)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                foreach (XmlNode chNode in node.ChildNodes)
                {
                    if (chNode.Name.Equals("Row"))
                        createRowElement(chNode, generalDictionary);

                    if (chNode.Name.Equals("Column"))
                        createColumnElement(chNode, generalDictionary);

                    if (chNode.Name.Equals("Input"))
                        createInputElement(chNode, inputDictionary);

                    if (chNode.Name.Equals("Output"))
                        createOutputElement(chNode, outputDictionary);

                    if (chNode.Name.Equals("Range"))
                        createRangeElement(chNode, rangeDictionary);

                    if (chNode.Name.Equals("Cell"))
                        createCellElement(chNode, cellDictionary);

                }
            }
        }

        private void createRowElement(XmlNode chNode, Dictionary<string, object> generalDictionary)
        {
            GeneralForm rowForm = new GeneralForm();
            String rowName = chNode.Attributes[0].Value;
            rowForm.Text = "Row " + rowName;
            rowForm.General_DescriptionGFBox.Text = chNode.Attributes[1].Value;
            rowForm.General_DescriptionGFBox.Select(0, 0);
            generalDictionary.Add(rowName, rowForm);
        }

        private void createColumnElement(XmlNode chNode, Dictionary<string, object> generalDictionary)
        {
            GeneralForm columnForm = new GeneralForm();
            String columnName = chNode.Attributes[0].Value;
            columnForm.Text = "Column " + columnName;
            columnForm.General_DescriptionGFBox.Text = chNode.Attributes[1].Value;
            columnForm.General_DescriptionGFBox.Select(0, 0);
            generalDictionary.Add(columnName, columnForm);
        }

        private void createInputElement(XmlNode chNode, Dictionary<string, object> inputDictionary)
        {
            InputOutputForm inputForm = new InputOutputForm();
            String inputName = chNode.Attributes[0].Value;
            inputForm.Text = "Input " + inputName;
            inputForm.inputOutputDescriptionBox.Text = chNode.Attributes[1].Value;
            inputForm.inputOutputDescriptionBox.Select(0, 0);
            inputDictionary.Add(inputName, inputForm);
        }

        private void createOutputElement(XmlNode chNode, Dictionary<string, object> outputDictionary)
        {
            InputOutputForm outputForm = new InputOutputForm();
            String outputName = chNode.Attributes[0].Value;
            outputForm.Text = "Output " + outputName;
            outputForm.inputOutputDescriptionBox.Text = chNode.Attributes[1].Value;
            outputForm.inputOutputDescriptionBox.Select(0, 0);
            outputDictionary.Add(outputName, outputForm);
        }

        private void createRangeElement(XmlNode chNode, Dictionary<string, object> rangeDictionary)
        {
            InputOutputForm rangeForm = new InputOutputForm();
            String rangeName = chNode.Attributes[0].Value;
            rangeForm.Text = "Range " + rangeName;
            rangeForm.inputOutputDescriptionBox.Text = chNode.Attributes[1].Value;
            rangeForm.inputOutputDescriptionBox.Select(0, 0);
            rangeForm.createCheckbox();
            rangeDictionary.Add(rangeName, rangeForm);
        }

        private void createCellElement(XmlNode chNode, Dictionary<string, object> cellDictionary)
        {
            //formula
            if (chNode.HasChildNodes)
            {
                CellForm cellForm = new CellForm();
                String cellName = chNode.Attributes[0].Value;
                cellForm.createFormSizes(cellElementList(chNode.ChildNodes));
                cellForm.Text = "Cell " + cellName;
                cellForm.General_DescriptionBox.Text = chNode.Attributes[1].Value;
                cellForm.General_DescriptionBox.Select(0, 0);
                if (chNode.ChildNodes.Count == 2)
                {
                    cellForm.inputBox.Text = chNode.ChildNodes.Item(0).Attributes[1].Value;
                    cellForm.OutputBox.Text = chNode.ChildNodes.Item(1).Attributes[0].Value;
                    cellDictionary.Add(cellName, cellForm);

                }

                else
                {

                    foreach (XmlNode child in chNode.ChildNodes)
                    {
                        if (child.Name.Equals("input"))
                        {
                            foreach (Control ctr in cellForm.Controls)
                            {
                                if (ctr is TextBox && ctr.Name.Equals(child.Attributes[0].Value))
                                {
                                    ctr.Text = child.Attributes[1].Value;
                                    break;
                                }
                            }
                        }
                        if (child.Name.Equals("output"))
                            cellForm.OutputBox.Text = child.Attributes[0].Value;
                    }
                    cellDictionary.Add(cellName, cellForm);
                }
            }
            //value
            else
            {
                GeneralForm cellForm = new GeneralForm();
                String cellName = chNode.Attributes[0].Value;
                cellForm.Text = "Cell " + cellName;
                cellForm.General_DescriptionGFBox.Text = chNode.Attributes[1].Value;
                cellForm.General_DescriptionGFBox.Select(0, 0);
                cellDictionary.Add(cellName, cellForm);
            }
        }

        private List<String> cellElementList(XmlNodeList childNode)
        {
            List<String> list = new List<string>();

            foreach (XmlNode node in childNode)
            {
                if (node.Name.Equals("input"))
                    list.Add(node.Attributes[0].Value);
            }

            return list;
        }

        private XmlNode checkNode(String path, String nodeName)
        {
            XmlNodeList node = doc.SelectNodes(path);

            foreach (XmlNode chNode in node)
            {
                if (chNode.Attributes[0].Value.Equals(nodeName))
                    return chNode;
            }
            return null;
        }

        public void readXMLToSpreadsheet(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            int row = 1;
            int column = 1;
            String path = AppDomain.CurrentDomain.BaseDirectory + "TemporaryXMLFileTemp.xml";

            if (worksheet.Cells[row, column].Value != null)
            {
                using (FileStream fs = File.Create(path))
                {
                    while (worksheet.Cells[row, column].Value != null)
                    {
                        Byte[] info = new UTF8Encoding(true).GetBytes(worksheet.Cells[row, column].Value);
                        // Adicionar informaçao ao ficheiro
                        fs.Write(info, 0, info.Length);

                        Array.Clear(info, 0, info.Length);
                        row++;
                    }
                }
            }
        }

        private String readRootXML(String path)
        {
            StreamReader reader;

            reader = new StreamReader(path);
            String xmlSecondLine = "";
            for (int i = 0; i < 2; i++)
                xmlSecondLine = reader.ReadLine();

            reader.Close();
            reader.Dispose();
            return xmlSecondLine;
        }

        private void clearContent(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            int row = 1;
            int column = 1;

            while (worksheet.Cells[row, column].Value != null)
            {
                worksheet.Cells[row, column].clear();
                row++;
            }

        }

        public void readXMLToInvisibleWorksheet(Microsoft.Office.Interop.Excel.Worksheet worksheet, String path)
        {
            XmlTextReader reader = null;
            int row = 1;
            int column = 1;

            clearContent(worksheet);
            try
            {
                // Variaveis declaradas que vao ser usadas no buffer
                Char[] buffer;
                int iCnt = 0;
                int charbuffersize;

                //Carrega no reader com a data do ficheiro. Ignora os espaços em branco
                reader = new XmlTextReader(path);
                reader.WhitespaceHandling = WhitespaceHandling.None;

                // Set variables used by ReadChars.
                charbuffersize = 35000;
                buffer = new Char[charbuffersize];


                worksheet.Cells[row, column].Value = readRootXML(path);
                row = 2;
                reader.MoveToContent();
                while ((iCnt = reader.ReadChars(buffer, 0, charbuffersize)) > 0)
                {
                    worksheet.Cells[row, column].Value = new String(buffer, 0, iCnt);
                    row++;
                    //Limpa o buffer
                    Array.Clear(buffer, 0, charbuffersize);
                }

            }
            finally
            {
                if (reader != null)
                {
                    worksheet.Cells[row, column].Value += "</Spreadsheet>";
                    Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    reader.Close();
                }
            }
        }
    }
}
