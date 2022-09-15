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
using Xceed.Document.NET;//*
using Xceed.Words.NET;
using System.Collections.Concurrent;

            //"Оператор using — это рекомендуемая альтернатива последовательности методов.Open, .Save, .Close.
            //Это гарантирует, что метод Dispose(внутренний метод, используемый Open XML SDK для очистки ресурсов) 
            //вызывается автоматически при достижении закрывающей фигурной скобки.Блок, который следует за оператором using, 
            //устанавливает область объекта, создаваемого или именуемого в операторе using. 
            //В этом случае это doc.Так как класс WordprocessingDocument в пакете Open XML SDK автоматически сохраняет и закрывает 
            //объект в реализации метода System.IDisposable и поскольку Dispose вызывается автоматически при выходе из блока, 
            //нет необходимости явно вызывать методы Save и Close, если вы используете оператор using."
            // - с сайта docs.microsoft.com про OpenXML


namespace WordExcel_Winforms_net6
{

    
    public class Basics  // TODO подгрузка элементов таблицы из БД, UI, разбивка на документы, ✅конвертировать массивы в List<T>✅, (если 1 учебный вопрос сильно маленький, переразбить ячейку)
	{                       // ✅переделать (ебейшее число аргументов и глобальный треш) в структуру, которую можно передавать функции Word🆗
                            // подумать, можно ли открыть файл один раз вне цикла, а в открытый док всё добавлять


        //[System.STAThread] // вот эта фигня была нужна, иначе не компилировалось
        public DocumentFormat.OpenXml.Wordprocessing.Table wordTable_Global;
        public string sourceFile;
        public string source_ext = "nothing"; //определяем, таблица или ворд файл мы открыли
        public string wordFile = String.Empty;
        public string[] books = new string[0]; //подгружаем список книг
        public string[] shapka = new string[0]; //содержание учебных занятий из источника
        public string[] contents_right = new string[0]; //подгружаем содержание для правого столбца из файла
        public OfficeOpenXml.ExcelWorksheet excelSheet_Global;
        public bool externalContensRight = false; // короче, это переменная обманчивая. на самом деле тут значение true и false должно восприниматься наоборот относительно названия переменной
        
        public ConcurrentQueue<WordArgs> argsQ = new();

        public Xceed.Words.NET.DocX document;





        internal async Task source_is_Excel()
		{
            using (ExcelPackage package = new ExcelPackage(new FileInfo(sourceFile)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Лист2"];
                excelSheet_Global = package.Workbook.Worksheets["Лист2"];
                //  using (wordArgs.the_doc = DocX.Load(wordArgs.wordFile)) // попытка использовать .Load вне цикла: сохраняется только 1й ппз. только он
                {
                    for (int k = 8; k <= 523; k++) // last working value = 89 / 523  // <- не забыть вернуть полный цикл
                    {
                        await Task.Run(() => insideLoopXls(k));
                        //progressBar1.PerformStep();
                    }

                    MessageBox.Show("Готово!");
                }
            }

            void insideLoopXls(int k)
            {
                WordArgs wordArgs = new WordArgs();
                string cell = "B" + k.ToString(); // эти две обязательно находятся внутри цикла
                string exCell = excelSheet_Global.Cells[cell].Value.ToString();
                //while(excelSheet_Global.Cells[cell].last



                //узнаем номер семестра, чтобы раскидывать литературу
                if (Regex.Match(exCell, @"\d\sсеместр").Success) wordArgs.semester = int.Parse(Regex.Match(exCell, @"\d+").Value);

                if (exCell.Contains("Тема ") && Char.IsNumber(exCell[5]) /* && exCell.Contains(". \"")*/)
                {
                    wordArgs.topicNow = int.Parse(Regex.Match(exCell, @"\d+").Value);
                    string topicName = exCell.Remove(0, 7);
                    wordArgs.fullTopic = String.Format("по теме № {0}. {1};", wordArgs.topicNow, topicName);
                }

                if (exCell.Contains("Практическое занятие №"))
                {
                    string pattern = @"Практическое занятие №(\s)?\d{1,4}([\s\p{P}])?([\d\s\p{P}])+\b";
                    string replacement = "";
                    wordArgs.clearCell = Regex.Replace(exCell, pattern, replacement);
                    wordArgs.exCell = exCell;
                    //		wordArgs.lessonNumber = Regex.Match(exCell, @"\d+").Value;
                    /* NEW:*/
                    //		wordArgs.lessonNumbers = new int[] { int.Parse(Regex.Match(exCell, @"\d+").Value) }; // попробуем заменить эту команду на вызов функции, пусть все работает универсально
                    wordArgs.lessonNumbers = LessonNumbers(exCell, wordArgs, k);
                    for (int cnt = 0; cnt < wordArgs.lessonNumbers.Length; cnt++)
                    {   // теперь тут еще и вложенный луп, чтобы экселевский метод работал, как вордовский, используя только массив номеров, а не отдельный номер
                        wordArgs.regexOperations(wordArgs);
                        wordArgs.topicNumForDequeue = wordArgs.lessonNumbers[cnt];
                        argsQ.Enqueue(new WordArgs(wordArgs));
                    }

                }
                return;
            }
        }

        internal Task WordBuild(WordArgs wordArgs) // TODO в методе нужно заменить переменную номера занятия на массив как для ворда так и для экселя
        {
            //выбираем книжки по семестру)))
            int a = 0, b = 0;
            switch (wordArgs.semester)  // перенести всю эту историю в Main
            {
                case 2: a = 0; b = 1; break;
                case 3: a = 1; b = 2; break;
                case 4: a = 1; b = 3; break;
                case 5: a = 2; b = 4; break;
                case 6: a = 3; b = 4; break;
                case 7: a = 4; b = 7; break;
                case 8: a = 5; b = 6; break;
                case 9: a = 5; b = 7; break;
            }

                            

                    document.SetDefaultFont(new Xceed.Document.NET.Font("Times New Roman"), 14);
                    // Create a paragraph and insert text.
                    /*TODO в строчке снизу есть баг с ошибкой оверфлоу, потому что до внедрения очереди 
                     * глобальный каунтер из класса управлялся циклом for вне этого класса. изза этого 
                     * при первоначальном внедрении отдельного метода Dequeue этот каунтер не сбрасывается, потому что*/
                    string header = String.Format(shapka[0] + wordArgs.topicNumForDequeue); //this one
                    document.InsertParagraph(header).Alignment = Alignment.center;
                    document.InsertParagraph(wordArgs.fullTopic).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[1]).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[2]).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[3]).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[4]).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[5]).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[6]).Alignment = Alignment.both;
                    document.InsertParagraph(shapka[7]).Alignment = Alignment.both;
                    document.InsertParagraph("  1. " + books[a] + "\n  2. " + books[b]).Alignment = Alignment.left;
                    document.InsertParagraph(shapka[8]).Alignment = Alignment.both;
                    var p2 = document.InsertParagraph();
                    var t = p2.InsertTableAfterSelf(2, 2);
                    document.InsertParagraph("\n");


                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("Учебные вопросы и время, отведенное на их рассмотрение");
                        t.Rows[0].Cells[1].Paragraphs[0].Append("Методические рекомендации руководителю учебного занятия");
                        t.Rows[1].Cells[0].Paragraphs[0].Append("Вступительная часть (5 минут)");
                        t.Rows[1].Cells[1].Paragraphs[0].Append("Доклад командира группы о готовности к занятиям. Объявление темы и целей занятий.");

                        for (int i = 0; i < wordArgs.splittedText.Count; i++)
                        {
                            Row dynamicRow = t.InsertRow();
                            dynamicRow.Cells[0].Paragraphs.First().Append(wordArgs.parts[i]);
                            dynamicRow.Cells[0].Paragraphs[0].Append(wordArgs.splittedText[i]);
                            //дальше выбираем, будет ли правый столбец будет браться из внешнего файла, или будет дополняться содержанием правого столбца
                            if (externalContensRight == false) //
                            {
                                dynamicRow.Cells[1].Paragraphs.First().Append
                                                    (contents_right[i % contents_right.Length]); // здесь знак "%" возвращает в начало содержания, если в левом столбце больше частей, чем в данном массиве 
                            }
                            else
                            {
                                if (wordArgs.splittedText[i].Contains("активной форме"))
                                {//Если есть активная форма, то не нужно ничего подставлять в начале

                                    dynamicRow.Cells[1].Paragraphs.First().Append(wordArgs.splittedText[i]);
                                }
                                else
                                {
                                    dynamicRow.Cells[1].Paragraphs.First().Append
                                                            ("Группа изучает " + char.ToLower(wordArgs.splittedText[i][0]) + wordArgs.splittedText[i].Substring(1));
                                }

                            }
                            t.Rows.Add(dynamicRow);
                        }

                        Row nextStaticRow = t.InsertRow();
                        nextStaticRow.Cells[0].Paragraphs[0].Append("Заключительная часть (5 минут)");
                        nextStaticRow.Cells[1].Paragraphs[0].Append("Подведение итогов занятия. Ответ на вопросы слушателей. Выставление оценок. Объявление задания для самостоятельной работы.");
                        t.Rows.Add(nextStaticRow);
                    }
                    document.Save(); // Save this document to disk. 


            return Task.CompletedTask;
        }

        public int searchInSource(int column, int row) //
        {
            //Console.WriteLine("вызов поиска");
            string exCell;
            char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray(); // алфавит как раз-таки для EPPlus
            try //как будто бы я здесь определил всю логику работы метода как для вызова из экселя, так и для вызова из ворда
            {
                if (source_ext == ".xlsx")
                {
                    string cell = alpha[column].ToString() + row.ToString(); // ну и эта по идее тоже для Эксель
                    exCell = excelSheet_Global.Cells[cell].Value.ToString(); // эта строка точно для работы с методом Excel в формате EPPlus. для этого же и алфавит вверху
                }
                else
                {
                    var rowHere = wordTable_Global.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ElementAt(row);
                    exCell = rowHere.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(column).InnerText;
                }

                var match = Regex.Match(exCell, @"\d*");
                //if (match.Success) Console.WriteLine("матч: " + match.Value + " - " + int.Parse(match.Value));
                return int.Parse(match.Value);
            }
            catch (ArgumentNullException)
            {
                //Console.WriteLine("EX 1");
                return 0;
            }
            catch (FormatException)
            {
                //Console.WriteLine("EX 2");
                return 0;
            }
        }
        internal int[] LessonNumbers(string celltext, WordArgs wordArgs, int iterOut)
        {
            //сначала берем только максимум два числа после "№", а потом эту строку разбиваем на 2 отдельных числа и всё заебись
            var preMatch = Regex.Match(celltext, @"занятие\s?№?\s?\d{1,3}(-\d{1,3})?").ToString();// "занятие\s?№?\s?\d{1,3}(-\d{1,3})?"gm
            MatchCollection matches = Regex.Matches(preMatch, @"\d+");


            if (matches.Count > 1)
            {
                int rng_one = int.Parse(matches.First().Value);
                int rng_last = int.Parse(matches[1].Value);
                int rng_count = (rng_last - rng_one + 1);
                //Console.WriteLine("первый: " + rng_one); // это всё дебаг инструкции, потом удалить
                //Console.WriteLine("последний: " + rng_last); // это всё дебаг инструкции, потом удалить
                //Console.WriteLine("количество: " + rng_count + "\n"); // это всё дебаг инструкции, потом удалить
                wordArgs.scanLast.lastLesson = rng_last;
                wordArgs.scanLast.lastTopic = wordArgs.topicNow;
                int[] lessonNumbers = new int[rng_count];
                int j = 0;
                for (int i = rng_one; i <= rng_last; i++)
                {
                    Console.Write(i + " ");
                    lessonNumbers[j++] = i;
                }
                return lessonNumbers;
            }
            else if (matches.Count == 1)
            {
                int rng_one = int.Parse(matches.First().Value);
                //Console.WriteLine("единственный: " + rng_one);
                wordArgs.scanLast.lastLesson = rng_one;
                wordArgs.scanLast.lastTopic = wordArgs.topicNow;
                int[] lessonNumbers = new int[] { rng_one };
                return lessonNumbers;
            }
            else
            {
                //Console.WriteLine("изначально отсутствует номер, запускаем просчет");

                List<int> missedNumbers = new List<int>();
                for (int i = 0; i < 14; i++)
                {   // эта функция возвращает наибольшее количество часов в строке, что чаще всего соответствует колонке "практические занятия"
                    missedNumbers.Add(searchInSource(i, iterOut));
                }
                //Console.Write("просчет: ");
                //foreach (int s in missedNumbers) Console.Write(s + "   ");
                //Console.WriteLine("\n длина массива: " + missedNumbers.Count);
                int divided = missedNumbers.Max() / 2;
                int[] lessonNumbers = new int[divided];
                if (wordArgs.topicNow == wordArgs.scanLast.lastTopic)// смотрим, поменялась ли тема
                {
                    //Console.WriteLine("прошлая и текущая темы совпадают");
                    //Console.WriteLine("прошлое занятие: №" + wordArgs.scanLast.lastLesson);
                    wordArgs.scanLast.lastLesson += 1;
                    for (int i = 0; i < lessonNumbers.Length; i++)
                    {
                        //lessonNumbers[i] = wordArgs.scanLast.lastLesson + 1 + i;
                        lessonNumbers[i] = wordArgs.scanLast.lastLesson + i;
                        //Console.WriteLine("занятие: " + lessonNumbers[i]);
                    }
                    wordArgs.scanLast.lastLesson += lessonNumbers.Length - 1; // тут отнимаем 1, т.к. сначала мы искусственно прибавили к прошлому занятию 1, поэтому увеличение на длину массива получается излишним
                }
                else
                {
                    //Console.WriteLine("прошлая и текущая темы отличаются");
                    //Console.WriteLine("прошлая: № " + wordArgs.scanLast.lastTopic);
                    //Console.WriteLine("текущая: № " + wordArgs.topicNow);
                    //Console.WriteLine("прошлое занятие: №" + wordArgs.scanLast.lastLesson);
                    wordArgs.scanLast.lastTopic = wordArgs.topicNow; // либо так, либо наоборот, не допетрил пока. а еще возможно до лупа ее вставить надо
                    for (int i = 0; i < lessonNumbers.Length; i++)
                    {
                        lessonNumbers[i] = i + 1;
                        Console.WriteLine("занятие: " + lessonNumbers[i]);
                    }
                    wordArgs.scanLast.lastLesson = lessonNumbers[^1];
                }
                return lessonNumbers;
            }
            ////    Console.WriteLine()
            ////    Console.WriteLine()
            ////Console.WriteLine("Тут " + ( - ) + 1 + " уроков");
            //foreach (var match in matches)
            //    {
            //        Console.Write(match + " - ");// match.Value will contain one of the matches
            //    }
        }


            

    }
}
