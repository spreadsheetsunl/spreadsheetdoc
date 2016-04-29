using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FirstExcelAddin
{
    public partial class CellForm : Form
    {
        public String cellType;
        public String formula;
        public TextBox textBox;
        public Label labelBox;
        private bool clearButtonClicked;
        private static String cellFormula = "CELL_FORMULA";
        private static String columnType = "COLUMN";
        private static String rowType = "ROW";

        public Dictionary<String, String> argumentsText = new Dictionary<String, String>();

        public CellForm()
        {
            InitializeComponent();
            this.FormClosing += CellForm_FormClosing;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cellType = getRibbon().cellType();
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            //grava no dicionario global.
            saveArgumentsText();
            getRibbon().saveCellFormula(getRibbon().getCellFormName(), this);
            getRibbon().InputRibbon.Enabled = false;
            getRibbon().OutputRibbon.Enabled = false;
            if (checkIfTextBoxIsEmpty())
                getRibbon().cellDocRibbon.Enabled = false;
            else
                getRibbon().cellDocRibbon.Enabled = true;
            //grava no dicionario do form.
            clearButtonClicked = false;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            undoCreatedTextBoxDynamically();
            if (clearButtonClicked)
                copyFromDictionaryToTexBox();
            clearButtonClicked = false;
        }

        private void remove_button_Click(object sender, EventArgs e)
        {
            getRibbon().removeCell(getRibbon().getCellFormName(), cellFormula);
            getRibbon().InputRibbon.Enabled = true;
            getRibbon().OutputRibbon.Enabled = true;
            getRibbon().cellDocRibbon.Enabled = false;
            clearButtonClicked = false;
        }

        private void clear_button_Click(object sender, EventArgs e)
        {
            clearButtonClicked = true;
            clearCreatedTextBoxDynamically();
        }

        private Ribbon1 getRibbon()
        {
            return Globals.Ribbons.Ribbon1;
        }

        private void CellForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                undoCreatedTextBoxDynamically();

                if (clearButtonClicked)
                    copyFromDictionaryToTexBox();
            }
        }

        //conforme o numero de argumentos o tamanho do form é adaptado de forma a ficar
        //com uma interface mais agradável
        public void createFormSizes(List<String> arguments)
        {
            if (arguments.Count == 2)
            {
                this.Size = new System.Drawing.Size(536, 393);
                createTextBoxes(arguments);
            }
            else if (arguments.Count > 2)
            {
                this.Size = new System.Drawing.Size(536, 474);
                createTextBoxes(arguments);
            }
            else
            {
                this.inputBox.Name = arguments[0];
                this.inputCell.Text = labelText(arguments[0]);
            }

        }

        //Criação das caixas de texto dinamicamente
        private void createTextBoxes(List<String> arguments)
        {
            int heightTextBoxLocation = 115;
            int heightLabelBoxLocation = 118;
            for (int i = 0; i < arguments.Count; i++)
            {
                if (i == 0)
                {
                    this.inputCell.Name = "Label" + i;
                    this.inputCell.Text = labelText(arguments[i]);
                    this.inputBox.Name = arguments[i];
                }
                else
                {
                    //criar dinamicamente caixas para os argumentos
                    textBox = new TextBox();
                    textBox.Name = arguments[i];
                    textBox.Location = new Point(118, heightTextBoxLocation);
                    textBox.AutoSize = false;
                    textBox.Size = new System.Drawing.Size(378, 50);
                    textBox.Multiline = true;
                    textBox.Visible = true;


                    //criar dinamicamente labels dos respetivos argumentos
                    labelBox = new Label();
                    labelBox.AutoSize = true;
                    labelBox.Name = "Label" + i;
                    labelBox.Text = labelText(arguments[i]);
                    labelBox.Location = new Point(15, heightLabelBoxLocation);

                    //adicionar ao controlador as caixas de texto e as labels
                    this.Controls.Add(textBox);
                    this.Controls.Add(labelBox);
                }
                heightTextBoxLocation += 56;
                heightLabelBoxLocation += 56;
            }

            heightTextBoxLocation += 10;
            heightLabelBoxLocation += 13;
            this.Output.Location = new Point(15, heightLabelBoxLocation);
            this.OutputBox.Location = new Point(118, heightTextBoxLocation);
        }

        //Criação da label da respetiva caixa de texto. 
        private String labelText(String argument)
        {
            String labelTextCreation = "";
            if (argument.Contains(":"))
            {
                string[] cellArguments = argument.Split(':');
                String inference = "";
                for (int i = 0; i < cellArguments.Length; i++)
                {
                    String columnLetter = splitCell(cellArguments[i], columnType);
                    int row = Int32.Parse(splitCell(cellArguments[i], rowType));
                    int column = getRibbon().getExcelColumnNumber(columnLetter);
                    if (i == 0)
                        inference = getRibbon().getHeader2(row, column) + " ...\n";
                    else
                        inference += getRibbon().getHeader2(row, column);
                }
                if (inference.Equals(" ...\n"))
                    labelTextCreation += argument + " (" + getRibbon().getCellType(argument) + ")";
                else
                    labelTextCreation += argument + " (" + getRibbon().getCellType(argument) + ")\n(" + inference + ")";
            }
            else
            {
                String columnLetter = splitCell(argument, columnType); // aqui
                int row = Int32.Parse(splitCell(argument, rowType));
                int column = getRibbon().getExcelColumnNumber(columnLetter);
                labelTextCreation += argument + " (" + getRibbon().getCellType(argument) + ")\n" + getRibbon().getHeader(row, column);
            }
            return labelTextCreation;
        }

        //Limpar o conteudo das caixas de texto criadas dinamicamente.
        private void clearCreatedTextBoxDynamically()
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                    ((TextBox)ctr).Clear();
            }
        }

        //Nao gravar o que foi escrito nas caixas de texto criadas dinamicamente (caso nao se carregue ok)
        private void undoCreatedTextBoxDynamically()
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                    ((TextBox)ctr).Undo();
            }
        }

        //Guarda o conteudo de cada caixa de texto. Cada caixa de texto tem o seu respetivo id.
        private void saveArgumentsText()
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                {
                    if (!argumentsText.ContainsKey(ctr.Name))
                        argumentsText.Add(ctr.Name, ctr.Text);
                    else
                        argumentsText[ctr.Name] = ctr.Text;
                }
            }
        }

        //Copiar o conteudo para cada caixa de texto. Quando se faz cancel ou close o utilizador nao pretende que nenhuma
        //alteraçao seja gravada. Entao é preciso garantir que se mantem tudo consistente.
        private void copyFromDictionaryToTexBox()
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                {
                    ((TextBox)ctr).Text = argumentsText[ctr.Name];
                }
            }
        }

        //Verificar se existe alguma caixa de texto com conteudo. Se existir permite o botao de visualizar a documentaçao
        //da respetiva celula. Caso contrario o botao mantem-se inativo apesar da celula em questao estar criada no dicionario
        public Boolean checkIfTextBoxIsEmpty()
        {

            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                {
                    if (!((TextBox)ctr).Text.Equals(String.Empty))
                        return false;
                }
            }
            return true;
        }
        //Metodo para o botão de visualizar a documentaçao de uma formula
        public void readOnlyTextBox()
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                    ((TextBox)ctr).Enabled = false;
            }
            this.remove_buttonCell.Visible = false;
            this.clear_buttonCell.Visible = false;
            this.okButtonCell.Visible = false;
            this.cancelButtonCell.Visible = true;
            this.cancelButtonCell.Text = "CLOSE";
            this.ShowDialog();
        }
        //Metodo para o botão de editar/criar a documentaçao de uma formula
        public void writeReadTextBox()
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is TextBox)
                    ((TextBox)ctr).Enabled = true;
            }
            this.remove_buttonCell.Visible = true;
            this.remove_buttonCell.Enabled = true;
            this.clear_buttonCell.Visible = true;
            this.okButtonCell.Visible = true;
            this.cancelButtonCell.Text = "CANCEL";
            this.ShowDialog();
        }

        //ir ao dicionario local e copiar o valor da caixa de texto para o novo form criado
        public void copyTextBox(String key, String text)
        {
            ((TextBox)this.Controls.Find(key, false).FirstOrDefault()).Text = text;
        }

        //Método para carregar as labels corretas. Quando se faz import nao se coloca logo as labels, só ao abrir é que tudo é
        //calculado. Garante que sempre que o utilizador abrir uma formula ve as labels com a informaçao atual
        public void loadInputLabels(String argument, int i, int row, int column)
        {
            foreach (Control ctr in this.Controls)
            {
                if (ctr is Label)
                {
                    if (ctr.Name.Equals("Label" + i))
                    {
                        String labelTextCreation = "Input \n";
                        labelTextCreation += argument + " (" + getRibbon().getCellType(argument) + ")\n" + getRibbon().getHeader(row, column);
                        ((Label)ctr).Text = labelTextCreation;
                    }
                }
            }
        }

        //Dividir a celula. Por exemplo, D6 divide em D e 6
        private string splitCell(String cell, String delimiter)
        {
            string pattern = "(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)";
            string[] result = Regex.Split(cell, pattern, RegexOptions.IgnoreCase);

            if (delimiter.Equals(columnType))
            {
                if (result[0].Contains('$'))
                {
                    string[] columnAux = result[0].Split('$');
                    return columnAux[1];
                }

                else
                    return result[0];
            }
                
            else
                return result[1];
        }
    }
}
