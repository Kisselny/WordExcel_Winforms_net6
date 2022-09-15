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
        
        public Basics b1 = new Basics();
        private int ops = 0;

        public void TemporaryTestMethod()
        {
            Random random = new Random();
            string st = random.Next(0, 100).ToString();
            b1.wordFile = @$"E:\Doki\�������\{st}.docx";
            textBox2.Text = @$"E:\Doki\�������\{st}.docx";
            using (var document = DocX.Create(b1.wordFile))
            {
                document.Save();
            }
            b1.source_ext = Path.GetExtension(textBox1.Text);
            textBox3.Text = @"C:\����������-�����.txt";
            b1.books = System.IO.File.ReadAllLines(@"C:\����������-�����.txt");
            b1.shapka = System.IO.File.ReadAllLines(@"C:\�����-�����.txt");
            textBox4.Text = @"C:\�����-�����.txt";
            b1.contents_right = System.IO.File.ReadAllLines(@"C:\����������-�����.txt");
            textBox5.Text = @"C:\����������-�����.txt";
        }

        public Form1(Form1 masterForm)
        {
            InitializeComponent();

        }

        public Form1()
        {
            InitializeComponent();
            //label1.Text = "Ready";
            TemporaryTestMethod();// ������ �����, ��� ����� �������
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1(this);
            btn1Func();
        }

        private async void btn1Func()
        {
            if (b1.wordFile != String.Empty && b1.books.Length != 0 && b1.shapka.Length != 0 && (checkBox1.Checked || b1.contents_right.Length != 0))
            {

                if (b1.source_ext == ".xlsx")
                {

                    await b1.source_is_Excel();
                }
                else if (b1.source_ext == ".docx")
                {
                    await source_XML_Word();

                }
                else
                {
                    MessageBox.Show("��. �������, �� �� ������� ������� � �������� ��������� ���. ��� ������� ������� ������, ��� �� ������ ����������:)");
                    return;
                }
                await MasterFnc();
            }
            else
            {   //��������� ��, ���� ����� �� �������
                if (b1.wordFile == String.Empty) MessageBox.Show("��. �������, �� �� �������, ��� ��������� ���� ������������ �������. �� �� ����� ���-�� �������:)");
                if (b1.books.Length == 0) MessageBox.Show("���������� ������� ������ ����������:)");
                if (b1.shapka.Length == 0) MessageBox.Show("����� ������� ����, � ������� ��������� \"�����\" ��� :)");
                if (b1.contents_right.Length == 0) MessageBox.Show("���������� ������� ���� � ����������� ������� ��������, ���� ������ ������� \"����� �� ��������\":)");
            }
            return;
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
                openFileDialog.Title = "�������� ������� � ���";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    b1.sourceFile = openFileDialog.FileName;
                    textBox1.Text = openFileDialog.FileName;
                    b1.source_ext = Path.GetExtension(openFileDialog.FileName); //����������, ������� ��� ���� ���� �� �������
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.RestoreDirectory = true;
                    openFileDialog.Filter = "��������� ���� (*.txt)|*.txt|All Files(*.*)|*.*";
                    openFileDialog.Title = "�������� ��������� ���� � �����������";
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
                openFileDialog.Filter = "��������� ���� (*.txt)|*.txt|All Files(*.*)|*.*";
                openFileDialog.Title = "�������� ��������� ���� � ������";
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
                openFileDialog.Filter = "��������� ���� (*.txt)|*.txt|All Files(*.*)|*.*";
                openFileDialog.Title = "�������� ��������� ���� � �����������";
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

        private void label1_Click(object sender, EventArgs e)
        {

        }

        public async Task MasterFnc()
        {
            //while (b1.argsQ.Count > 0)
            //{
            //    using (b1.document = DocX.Load(b1.wordFile))
            //    {
            //        //var p = b1.argsQ.Dequeue();
            //        await Task.Run(() => b1.WordBuild(b1.argsQ.Dequeue()));
            //        ops++;
            //        label1.Text = String.Format("��������� {0} ������ ������������ �������", ops.ToString());
            //        label2.Text = String.Format("������� �������: " + b1.argsQ.Count.ToString());
            //    }
            //}
            var p = new WordArgs();
            while (b1.argsQ.TryDequeue(out p))
            {
                using (b1.document = DocX.Load(b1.wordFile))
                {
                    await Task.Run(() => b1.WordBuild(p));
                    ops++;
                    //label1.Text = String.Format("��������� {0} ������ ������������ �������", ops.ToString());
                    label2.Text = String.Format("������� �������: " + b1.argsQ.Count.ToString());
                }
            }

        }


        public async Task source_XML_Word()
        {

            int creationOps = 0;
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument src_docx =
                WordprocessingDocument.Open(b1.sourceFile, true))
            {
                //Find the  table in the document.
                b1.wordTable_Global =
                    src_docx.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ElementAt(1);

                int row_count = b1.wordTable_Global.Elements<TableRow>().Count();
                Console.WriteLine("�����: " + row_count);
                TableRow row; //��������� ��� ����� �����, ����� �� ������ ����-�����
                TableCell cell;

                for (int i = 0; i < row_count; i++)
                {
                    await Task.Run(() => insideLoop(i));

                    //label1 = new Label();
                    //creationOps++;
                    //label1.Text = String.Format(creationOps.ToString());


                    progressBar1.PerformStep();
                }
                return;

                void insideLoop(int i)
                {
                    WordArgs wordArgs = new WordArgs();
                    try
                    {
                        row = b1.wordTable_Global.Elements<TableRow>().ElementAt(i);// Find the second row in the table.
                        cell = row.Elements<TableCell>().ElementAt(1);// Find the third cell in the row.
                        wordArgs.clearCell = row.Elements<TableCell>().ElementAt(2).InnerText;

                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        //continue;
                        return;
                    }

                    if (cell.InnerText.ToString().Contains("�������")) wordArgs.semester = int.Parse(Regex.Match(cell.InnerText.ToString(), @"\d+").Value);

                    if (cell.InnerText.ToString().Contains("����"))
                    {
                        wordArgs.topicNow = int.Parse(Regex.Match(cell.InnerText.ToString(), @"\d+").Value);
                        wordArgs.fullTopic = String.Format("�� ���� � {0}. {1};", wordArgs.topicNow, wordArgs.clearCell);
                    }

                    if (cell.InnerText.ToString().Contains("������������ �������"))
                    {
                        wordArgs.lessonNumbers = b1.LessonNumbers(cell.InnerText, wordArgs, i);

                        for (int cnt = 0; cnt < wordArgs.lessonNumbers.Length; cnt++)
                        {
                            wordArgs.regexOperations(wordArgs);
                            wordArgs.topicNumForDequeue = wordArgs.lessonNumbers[cnt];


                            // ���������� � �����: ������� ���������� � ������� �� ������, � �� �� ��������
                            //���-�� ���-�� ����������� � ������������ �������
                            //https://stackoverflow.com/questions/16601750/c-sharp-queue-objects-modified-in-queue-after-being-enqueued
                            //https://stackoverflow.com/questions/78536/deep-cloning-objects/78577#78577
                            //����� ��������, ��������� ����������� �����. �� ��������

                            b1.argsQ.Enqueue(new WordArgs(wordArgs));
                        }
                    }
                    
                }

            }
        }



    }
}