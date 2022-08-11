using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Text.RegularExpressions;
//using DocumentFormat.OpenXml.Wordprocessing;

namespace WordExcel_Winforms_net6
{
    public partial class Form1 : Form
    {
        WordArgs wordArgs = new WordArgs();

        public Form1()
        {
            InitializeComponent();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 522;
            progressBar1.Step = 1;

        }

        public void IncrementBar()
        {
            progressBar1.Increment(1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (wordArgs.wordFile != String.Empty && wordArgs.books.Length != 0 && wordArgs.shapka.Length != 0 && (checkBox1.Checked || wordArgs.contents_right.Length != 0))
            {
                if (wordArgs.source_ext == ".xlsx")
                {
                    
                    Basics.source_is_Excel(wordArgs);
                }
                else if (wordArgs.source_ext == ".docx")
                {
                    Basics.source_XML_Word(wordArgs);
                }
                else
                {
                    MessageBox.Show("Ой. Кажется, вы не выбрали темплан в качестве источника ППЗ. Это следует сделать прежде, чем мы сможем продолжить:)");
                } 
            }
            else
            {   //проверяем всё, чего может не хватать
                if(wordArgs.wordFile == String.Empty) MessageBox.Show("Ой. Кажется, вы не указали, где сохранить план практических занятий. Их же нужно где-то хранить:)");
                if(wordArgs.books.Length == 0) MessageBox.Show("Необходимо выбрать список литературы:)");
                if (wordArgs.shapka.Length == 0) MessageBox.Show("Нужно выбрать файл, в котором прописана \"шапка\" ППЗ :)");
                if (wordArgs.contents_right.Length == 0) MessageBox.Show("Необходимо выбрать файл с содержанием учебных вопросов, либо нажать галочку \"Взять из темплана\":)");
            }
        }



        private void button1_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Word file (*.docx)|*.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                wordArgs.wordFile = saveFileDialog1.FileName;
                textBox2.Text = saveFileDialog1.FileName;
            }
            Console.WriteLine(wordArgs.wordFile);

            using (var document = DocX.Create(wordArgs.wordFile))
            {
                document.Save();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Filter = "Excel table (*.xlsx)|*.xlsx|Word docx (*.docx)|*.docx|All Files(*.*)|*.*";
                openFileDialog.Title = "Выберите таблицу с ппз";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordArgs.sourceFile = openFileDialog.FileName;
                    textBox1.Text = openFileDialog.FileName;
                    wordArgs.source_ext = Path.GetExtension(openFileDialog.FileName); //определяем, таблица или ворд файл мы открыли
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.RestoreDirectory = true;
                    openFileDialog.Filter = "Текстовый файл (*.txt)|*.txt|All Files(*.*)|*.*";
                    openFileDialog.Title = "Выберите текстовый файл с литаратурой";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox3.Text = openFileDialog.FileName;
                        wordArgs.books = System.IO.File.ReadAllLines(openFileDialog.FileName);
                    }
                }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Filter = "Текстовый файл (*.txt)|*.txt|All Files(*.*)|*.*";
                openFileDialog.Title = "Выберите текстовый файл с шапкой";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordArgs.shapka = System.IO.File.ReadAllLines(openFileDialog.FileName);
                    textBox4.Text = openFileDialog.FileName;
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Filter = "Текстовый файл (*.txt)|*.txt|All Files(*.*)|*.*";
                openFileDialog.Title = "Выберите текстовый файл с содержанием";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordArgs.contents_right = System.IO.File.ReadAllLines(openFileDialog.FileName);
                    textBox5.Text = openFileDialog.FileName;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            switch (checkBox1.CheckState)
            {
                case CheckState.Checked:
                    wordArgs.externalContensRight = true;
                    break;
                case CheckState.Unchecked:
                    wordArgs.externalContensRight = false;
                    break;
            }
        }
    }
}