using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows.Forms;
////using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Text.RegularExpressions;
//using DocumentFormat.OpenXml.Wordprocessing;


using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
//using System.Collections.Generic;
using OfficeOpenXml;
//using Paragraph = Xceed.Document.NET.Paragraph;
using System.Text.RegularExpressions;
using System.Linq;
using System.Windows.Forms;
//using System.IO.Packaging;
//using System.Text.RegularExpressions;
//using Xceed.Document.NET;
////using Xceed.Words.NET;



namespace WordExcel_Winforms_net6
{
    public partial class Form1 : Form
    {


        Basics b1 = new Basics();
        public Form1()
        {
            InitializeComponent();
        }
        


        private void button1_Click(object sender, EventArgs e)
        {
            if (b1.wordFile != String.Empty && b1.books.Length != 0 && b1.shapka.Length != 0 && (checkBox1.Checked || b1.contents_right.Length != 0))
            {
                if (b1.source_ext == ".xlsx")
                {
                    
                    b1.source_is_Excel();
                }
                else if (b1.source_ext == ".docx")
                {
                    b1.source_XML_Word();
                }
                else
                {
                    MessageBox.Show("Ой. Кажется, вы не выбрали темплан в качестве источника ППЗ. Это следует сделать прежде, чем мы сможем продолжить:)");
                } 
            }
            else
            {   //проверяем всё, чего может не хватать
                if(b1.wordFile == String.Empty) MessageBox.Show("Ой. Кажется, вы не указали, где сохранить план практических занятий. Их же нужно где-то хранить:)");
                if(b1.books.Length == 0) MessageBox.Show("Необходимо выбрать список литературы:)");
                if (b1.shapka.Length == 0) MessageBox.Show("Нужно выбрать файл, в котором прописана \"шапка\" ППЗ :)");
                if (b1.contents_right.Length == 0) MessageBox.Show("Необходимо выбрать файл с содержанием учебных вопросов, либо нажать галочку \"Взять из темплана\":)");
            }
        }



        private void button1_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Word file (*.docx)|*.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                b1.wordFile = saveFileDialog1.FileName;
                textBox2.Text = saveFileDialog1.FileName;
            }

            using (var document = DocX.Create(b1.wordFile))
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
                    b1.sourceFile = openFileDialog.FileName;
                    textBox1.Text = openFileDialog.FileName;
                    b1.source_ext = Path.GetExtension(openFileDialog.FileName); //определяем, таблица или ворд файл мы открыли
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
                    openFileDialog.Title = "Выберите текстовый файл с литературой";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox3.Text = openFileDialog.FileName;
                        b1.books = System.IO.File.ReadAllLines(openFileDialog.FileName);
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
                    b1.shapka = System.IO.File.ReadAllLines(openFileDialog.FileName);
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
                    b1.contents_right = System.IO.File.ReadAllLines(openFileDialog.FileName);
                    textBox5.Text = openFileDialog.FileName;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            switch (checkBox1.CheckState)
            {
                case CheckState.Checked:
                    b1.externalContensRight = true;
                    break;
                case CheckState.Unchecked:
                    b1.externalContensRight = false;
                    break;
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
        
    }
}