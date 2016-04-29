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
    class SaveXMLClass
    {
        XmlDocument doc = new XmlDocument();
        XmlElement element;


        public void initializeXML()
        {
            XmlElement root = doc.DocumentElement;
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.InsertBefore(xmlDeclaration, root);
        }

        public Boolean docExist()
        {
            if (doc.DocumentElement == null)
                return false;
            return true;
        }

        public void load(String fileName)
        {
            doc.Load(fileName);
        }

        public void loadDoc(String fileName)
        {
            doc.Load(fileName);
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.InsertBefore(xmlDeclaration, doc.DocumentElement);
        }
        public void addRootNode(String sheetName, String description)
        {
            if (doc.DocumentElement == null)
            {
                element = doc.CreateElement(string.Empty, "Spreadsheet", string.Empty);
                doc.AppendChild(createAttrNode(element, sheetName, description));
            }
            else
            {
                XmlElement root = doc.DocumentElement;
                root.Attributes[1].Value = description;
            }
        }

        private XmlElement createAttrNode(XmlElement element, String attrName, String attrDescription)
        {
            XmlAttribute nameAttr = doc.CreateAttribute("name");
            XmlAttribute descriptionAttr = doc.CreateAttribute("description");

            nameAttr.Value = attrName;
            descriptionAttr.Value = attrDescription;

            element.SetAttributeNode(nameAttr);
            element.SetAttributeNode(descriptionAttr);

            return element;
        }

        public void addWorksheetNodeToRoot(String worksheetName, String description)
        {
            if (doc.DocumentElement != null)
            {

                XmlNode node = checkNode("//Worksheet", worksheetName);
                if (node != null)
                    modifyNode(node, description);
                else
                {
                    element = doc.CreateElement(string.Empty, "Worksheet", string.Empty);
                    doc.DocumentElement.AppendChild(createAttrNode(element, worksheetName, description));
                }
            }
        }

        public void addColumnToWorksheet(String elementName, String worksheetName, String description)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyNode(childNode, description);

                else
                {
                    element = doc.CreateElement(string.Empty, "Column", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createAttrNode(element, elementName, description));
                }
            }
        }

        public void addRowToWorksheet(String elementName, String worksheetName, String description)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyNode(childNode, description);

                else
                {
                    element = doc.CreateElement(string.Empty, "Row", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createAttrNode(element, elementName, description));
                }
            }
        }

        public void addRangeToWorksheet(String elementName, String worksheetName, String description)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyNode(childNode, description);

                else
                {
                    element = doc.CreateElement(string.Empty, "Range", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createAttrNode(element, elementName, description));
                }
            }
        }

        public void addInputToWorksheet(String elementName, String worksheetName, String description)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyNode(childNode, description);
                else
                {
                    element = doc.CreateElement(string.Empty, "Input", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createAttrNode(element, elementName, description));
                }
            }
        }

        public void removeFromWorksheet(String elementName, String worksheetName)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    node.RemoveChild(childNode);
            }
        }

        public void removeInexistentWorksheet(Dictionary<Microsoft.Office.Interop.Excel.Worksheet, List<Dictionary<string, object>>> worksheetDictionary)
        {
            Boolean toRemove = false;
            XmlNode root = doc.DocumentElement;
            if (root != null)
            {
                foreach (XmlNode chNode in root.ChildNodes)
                {

                    foreach (KeyValuePair<Microsoft.Office.Interop.Excel.Worksheet, List<Dictionary<string, object>>> displayWorksheet in worksheetDictionary)
                    {

                        if (chNode.Attributes[0].Value.Equals(displayWorksheet.Key.Name))
                        {
                            toRemove = false;
                            break;
                        }

                        else
                            toRemove = true;
                    }

                    if (toRemove)
                        root.RemoveChild(chNode);
                }
            }
        }

        public void addOutputToWorksheet(String elementName, String worksheetName, String description)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyNode(childNode, description);
                else
                {
                    element = doc.CreateElement(string.Empty, "Output", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createAttrNode(element, elementName, description));
                }
            }
        }

        public void addCellWithValueToWorksheet(String elementName, String worksheetName, String description)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyNode(childNode, description);
                else
                {
                    element = doc.CreateElement(string.Empty, "Cell", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createAttrNode(element, elementName, description));
                }
            }
        }

        public void addCellWithFormulaToWorksheet(String elementName, String worksheetName, CellForm form)
        {
            XmlNode node = checkNode("//Worksheet", worksheetName);
            if (node != null)
            {
                XmlNode childNode = getNodeByWorksheet(node.ChildNodes, elementName);

                if (childNode != null)
                    modifyCellWithFormulaNode(childNode, form);
                else
                {
                    element = doc.CreateElement(string.Empty, "Cell", string.Empty);
                    XmlNode worksheet = checkNode("//Worksheet", worksheetName);
                    worksheet.AppendChild(createInputCellAttr(createAttrNode(element, elementName, form.General_DescriptionBox.Text), elementName, form));
                }
            }
        }

        private XmlElement createInputCellAttr(XmlElement parentElement, String parentName, CellForm form)
        {
            XmlElement child;
            XmlAttribute cellAttr;
            XmlAttribute descriptionAttr;

            foreach (KeyValuePair<String, String> children in form.argumentsText)
            {
                if (!children.Key.Equals("General_DescriptionBox") && (!children.Key.Equals("OutputBox")))
                {
                    child = doc.CreateElement(string.Empty, "input", string.Empty);
                    cellAttr = doc.CreateAttribute("cell");
                    descriptionAttr = doc.CreateAttribute("description");
                    cellAttr.Value = children.Key;
                    descriptionAttr.Value = children.Value;
                    child.SetAttributeNode(cellAttr);
                    child.SetAttributeNode(descriptionAttr);
                    parentElement.AppendChild(child);
                }
            }

            child = doc.CreateElement(string.Empty, "output", string.Empty);
            descriptionAttr = doc.CreateAttribute("description");
            descriptionAttr.Value = form.OutputBox.Text;
            child.SetAttributeNode(descriptionAttr);
            parentElement.AppendChild(child);
            return parentElement;
        }

        private XmlNode getNodeByWorksheet(XmlNodeList childNodes, String elementName)
        {
            foreach (XmlNode chNode in childNodes)
            {
                if (chNode.Attributes[0].Value.Equals(elementName))
                    return chNode;
            }
            return null;
        }

        private XmlNode checkNode(String path, String nodeName)
        {
            XmlNodeList node = doc.SelectNodes(path);

            foreach (XmlNode chNode in node)
            {
                System.Diagnostics.Debug.WriteLine(chNode.Name);
                if (chNode.Attributes[0].Value.Equals(nodeName))
                    return chNode;
            }
            return null;
        }

        private void modifyNode(XmlNode node, String attrDescription)
        {
            node.Attributes[1].Value = attrDescription;
        }

        public void modifyNodeNameAndAttribute(Microsoft.Office.Interop.Excel.Worksheet worksheet, String oldWorksheetName, String attrDescription)
        {
            XmlNode node = checkNode("//Worksheet", oldWorksheetName);
            node.Attributes[0].Value = worksheet.Name;
            node.Attributes[1].Value = attrDescription;
        }

        private void modifyCellWithFormulaNode(XmlNode node, CellForm form)
        {
            node.Attributes[1].Value = form.General_DescriptionBox.Text;

            foreach (XmlNode chNode in node.ChildNodes)
            {
                if (!chNode.Name.Equals("output"))
                {
                    if (form.argumentsText.ContainsKey(chNode.Attributes[0].Value))
                        chNode.Attributes[1].Value = form.argumentsText[chNode.Attributes[0].Value];
                }

                else
                    chNode.Attributes[0].Value = form.argumentsText["OutputBox"];
            }

        }

        public void saveXML(String name)
        {
            doc.Save(name);

        }

        public void saveHiddenFile(String name)
        {
            String path = AppDomain.CurrentDomain.BaseDirectory + name;
            try
            {
                doc.Save(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Could not save xml file. Original error: " + ex.Message);
            }
            
        }
        public String readRootXML()
        {
            StreamReader reader;
            String path = AppDomain.CurrentDomain.BaseDirectory + "TemporaryXMLFileTemp.xml";

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

        public void saveXMLToWorksheet(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            XmlTextReader reader = null;
            String path = AppDomain.CurrentDomain.BaseDirectory + "TemporaryXMLFileTemp.xml";
            int row = 1;
            int column = 1;

            clearContent(worksheet);
            try
            {

                // Declare variables used by ReadChars
                Char[] buffer;
                int iCnt = 0;
                int charbuffersize;

                // Load the reader with the data file.  Ignore white space.
                reader = new XmlTextReader(path);
                reader.WhitespaceHandling = WhitespaceHandling.None;

                // Set variables used by ReadChars.
                charbuffersize = 35000;
                buffer = new Char[charbuffersize];

                worksheet.Cells[row, column].Value = readRootXML();
                row = 2;
                reader.MoveToContent();
                while ((iCnt = reader.ReadChars(buffer, 0, charbuffersize)) > 0)
                {
                    worksheet.Cells[row, column].Value = new String(buffer, 0, iCnt);
                    row++;
                    // Clear the buffer.
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
            File.Delete(path);
        }
    }
}
