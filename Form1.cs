using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace DevIntOffice
{
    public partial class Form1 : Form
    {
        //Declaração das variáveis globais.
        string reg, uf, tipo, tec, coord,esp,end, gra, loc, est, cabo, sess, nev, dir = @"c:\DEV_INT", arq;

        private void button3_Click(object sender, EventArgs e)
        {
            GravarFicheiro();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Resete nas caixas para a escolha das opções.
            comboBox1.Text = null;
            comboBox2.Text = null;
            comboBox3.Text = null;
            //Resete nos campos.
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            GravarFicheiro();
            excelApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(ValidarDados() == true)
            {
                Adicionar();//Obtenção da variaveis as caixas de texto
            }
        }
        bool ValidarDados()
        {
            //Verificação do comprimento do campo comboBoxs.
            if (comboBox1.Text.Length == 0)
            {
                MessageBox.Show("Selecione uma regional");
                comboBox1.Focus();//Seleção da caixa de combinação
                return false;//Preenchimento incorreto
            }
            if (comboBox2.Text.Length == 0)
            {
                MessageBox.Show("Selecione uma UF");
                comboBox2.Focus();//Seleção da caixa de combinação
                return false;//Preenchimento incorreto
            }
            if (comboBox3.Text.Length == 0)
            {
                MessageBox.Show("Selecione um tipo");
                comboBox3.Focus();//Seleção da caixa de combinação
                return false;//Preenchimento incorreto
            }
            //Verificação do comprimento da cadeia de caracteres nos campos texBoxs. 
            if (textBox1.Text.Length < 1)
            {
                MessageBox.Show("Insira o nome do técnico");
                textBox1.SelectAll();//Seleção de todo o texto
                textBox1.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox2.Text.Length < 1)
            {
                MessageBox.Show("Insira o nome do coordenador de GA");
                textBox2.SelectAll();//Seleção de todo o texto
                textBox2.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox3.Text.Length < 1)
            {
                MessageBox.Show("Insira o numero espelho");
                textBox3.SelectAll();//Seleção de todo o texto
                textBox3.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox4.Text.Length < 1)
            {
                MessageBox.Show("Insira o endereço");
                textBox4.SelectAll();//Seleção de todo o texto
                textBox4.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox5.Text.Length < 1)
            {
                MessageBox.Show("Insira o GRA");
                textBox5.SelectAll();//Seleção de todo o texto
                textBox5.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox6.Text.Length < 1)
            {
                MessageBox.Show("Insira o local");
                textBox6.SelectAll();//Seleção de todo o texto
                textBox6.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox7.Text.Length < 1)
            {
                MessageBox.Show("Insira o nome da estação");
                textBox7.SelectAll();//Seleção de todo o texto
                textBox7.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox8.Text.Length < 1)
            {
                MessageBox.Show("Insira o numero do cabo");
                textBox8.SelectAll();//Seleção de todo o texto
                textBox8.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox9.Text.Length < 1)
            {
                MessageBox.Show("Insira o numero da sessão");
                textBox9.SelectAll();//Seleção de todo o texto
                textBox9.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            if (textBox10.Text.Length < 1)
            {
                MessageBox.Show("Insira o numero do evento");
                textBox10.SelectAll();
                textBox10.Focus();//Seleção da caixa de texto
                return false;//Preenchimento incorreto
            }
            return true;
        }
        void Adicionar()
        {
            //Obteñção das variaveis
            reg = comboBox1.Text ; 
            uf = comboBox2.Text;
            tipo = comboBox3.Text;
            //Devolução do resultado
            tec = textBox1.Text;
            coord = textBox2.Text;
            esp = textBox3.Text;
            end = textBox4.Text;
            gra = textBox5.Text;
            loc = textBox6.Text;
            est = textBox7.Text;
            cabo = textBox8.Text;
            sess = textBox9.Text;
            nev = textBox10.Text;
            Exportar();
        }
        void Exportar()
        {
            excelApp.Sheets["Planilha1"].Select(); //Seleção da planilha
            //Verificação da ultima linha preenchida.
            int linhaExcel = 2;//Dados apartir da 2 linha
            bool valor = true;//Verificação da célula preenchida

            while(valor == true)//Enquanto a célula da coluna A estiver preenchida ...
            {
                if (excelApp.Range["A" + linhaExcel].Value != null)
                {
                    valor = true;
                    linhaExcel = linhaExcel + 1;//Proxima linha
                }
                else
                {
                    valor = false;
                }
            }
            //Passagem dos dados do formulario pra a planilha 
            excelApp.Range["A" + linhaExcel].Value = reg;
            excelApp.Range["B" + linhaExcel].Value = uf;
            excelApp.Range["C" + linhaExcel].Value = tipo;
            excelApp.Range["D" + linhaExcel].Value = tec;
            excelApp.Range["E" + linhaExcel].Value = coord;
            excelApp.Range["F" + linhaExcel].Value = esp;
            excelApp.Range["G" + linhaExcel].Value = end;
            excelApp.Range["H" + linhaExcel].Value = gra;
            excelApp.Range["I" + linhaExcel].Value = loc;
            excelApp.Range["J" + linhaExcel].Value = est;
            excelApp.Range["K" + linhaExcel].Value = cabo;
            excelApp.Range["L" + linhaExcel].Value = sess;
            excelApp.Range["M" + linhaExcel].Value = nev;
            FormatarFicheiro();

        }
        Microsoft.Office.Interop.Excel.Application excelApp;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {   
            comboBox1.Text = null;
            comboBox2.Text = null;
            comboBox3.Text = null;
            //Carregamento dos comboBoxs
            comboBox1.Items.Add("RNE");
            comboBox1.Items.Add("RBA");
            comboBox1.Items.Add("RNO");

            comboBox2.Items.Add("CE");
            comboBox2.Items.Add("PB");
            comboBox2.Items.Add("PE");
            comboBox2.Items.Add("RN");

            comboBox3.Items.Add("PRIMAÁRIO");
            comboBox3.Items.Add("SECUNDÁRIO");
            comboBox3.Items.Add("RÍGIDO");


            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            //Verificar se a pasta existe,, caso não exista é criada.  
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);//Criação
                MessageBox.Show("Diretori foi criado para salvamento " + dir);
                
            }
          
            IniciarExcel();

            if(VerificarFicheiro() == false)
            {
                CriarFicheiro();
            }
            else
            {
                AbrirFicheiro();
            }
        }
        void IniciarExcel()
        {
            //Nova instância excel.
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Visibilidade.
            excelApp.Visible = true;
        }
        bool VerificarFicheiro()
        {
            //Caminho completo até o ficheiro
            arq = Application.StartupPath + @"\devExcel.xlsx";
            if (System.IO.File.Exists(arq))
            {
                return true;//O arquivo já existe.
            }
            else
            {
                return false;//O arquivo ainda não existe
            }
        }
        void CriarFicheiro()
        {
            excelApp.Workbooks.Add();//Novo arquivo
            excelApp.Sheets["Planilha1"].Select();//Seleção da primeira planilha 
            //Títulos das colunas
            excelApp.Range["A1"].Value = "REGIONAL";
            excelApp.Range["B1"].Value = "UF";
            excelApp.Range["C1"].Value = "TIPO";
            excelApp.Range["D1"].Value = "TÉCNICO";
            excelApp.Range["E1"].Value = "COORD GA";
            excelApp.Range["F1"].Value = "N ESPELHO";
            excelApp.Range["G1"].Value = "ENDEREÇO";
            excelApp.Range["H1"].Value = "GRA";
            excelApp.Range["I1"].Value = "LOCALIDADE";
            excelApp.Range["J1"].Value = "ESTAÇÃO";
            excelApp.Range["K1"].Value = "CABO";
            excelApp.Range["L1"].Value = "SESSÃO";
            excelApp.Range["M1"].Value = "N EVENTO";
            FormatarFicheiro();
        }
        void FormatarFicheiro()
        {
            //Titulo em negrito
            excelApp.Range["A1:M1"].Font.Bold = true;
            //Escala de cinza 50%
            excelApp.Range["A1:M1"].Interior.Color = Microsoft.Office.Interop.Excel.Constants.xlGray50;
            //Alinhamento do texto ao centro da coluna.
            excelApp.Range["A1:M1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            //Largura das colunas
            excelApp.Columns["A:A"].ColumnWidth = 10;
            excelApp.Columns["B:B"].ColumnWidth = 10;
            excelApp.Columns["C:C"].ColumnWidth = 15;
            excelApp.Columns["D:D"].ColumnWidth = 20;
            excelApp.Columns["E:E"].ColumnWidth = 20;
            excelApp.Columns["F:F"].ColumnWidth = 20;
            excelApp.Columns["G:G"].ColumnWidth = 50;
            excelApp.Columns["H:H"].ColumnWidth = 10;
            excelApp.Columns["I:I"].ColumnWidth = 20;
            excelApp.Columns["J:J"].ColumnWidth = 10;
            excelApp.Columns["K:K"].ColumnWidth = 10;
            excelApp.Columns["L:L"].ColumnWidth = 10;
            excelApp.Columns["M:M"].ColumnWidth = 30;

            if(VerificarFicheiro() == false)
            {
                GravarFicheiroComo();
            }
            else
            {
                GravarFicheiro();
            }
        }
        void GravarFicheiroComo()
        {
            //Gravação do arquivo
            excelApp.ActiveWorkbook.SaveAs(arq);
        }
        void GravarFicheiro()
        {
            //Gravação do arquivo
            excelApp.ActiveWorkbook.Save();
        }
        void AbrirFicheiro()
        {
            //Abertura do arquivo
            excelApp.Workbooks.Open(arq);
        }         
        
    }
}
