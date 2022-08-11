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
    internal class WordArgs
    {
        public int counter; //счетчик, используется, когда ПЗ пронумерованы пачками, в стиле "занятие №4-7"

        public int semester;
        public string fullTopic;
        public string sourceFile;
        public string wordFile = String.Empty;
        public string exCell;

        public OfficeOpenXml.ExcelWorksheet excelSheet_Global; //эти две херни для экселя и ворда, чтобы мы могли иметь доступ
        public DocumentFormat.OpenXml.Wordprocessing.Table wordTable_Global; //... к соответствующим таблицам из данного класса без этих непонятных using-ов со скобочками()
        //public Table wordTable; //нахера этот был нужен?? верхний я создал попозже, а этот как-будто бы не вызывался ваще ничем


        public string clearCell;
        public string source_ext = "nothing"; //определяем, таблица или ворд файл мы открыли

        public int lessonNow, topicNow; // это когда в ворде (или не только) номер следующего занятия не вписан
        public (int lastTopic, int lastLesson) scanLast; // попробуем prevLessonNum заменить на это, чтобы у предыдущего номера занятия был ассоциированный номер

        public int[] lessonNumbers;

        List<int> matches;

        char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray(); // алфавит как раз-таки для EPPlus
        public string[] books = new string[0]; //подгружаем список книг
        public string[] shapka = new string[0]; //содержание учебных занятий из источника
        public string[] contents_right = new string[0]; //подгружаем содержание для правого столбца из файла
        public bool externalContensRight = false; // короче, это переменная обманчивая. на самом деле тут значение true и false должно восприниматься наоборот относительно названия переменной
        public Xceed.Document.NET.Document the_doc; // это тип документа в var document = DocX.Load(wordArgs.wordFile)
                                                    //он нужен чтобы передать .Load вне цикла внутрь метода Word(), но почему-то сохраняется только  1й ппз


        public int searchInSource(int column, int row) //
        {
            //Console.WriteLine("вызов поиска");
            string exCell;
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

        public void WordBuild(WordArgs wordArgs) // TODO в методе нужно заменить переменную номера занятия на массив как для ворда так и для экселя
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


            //       var document = wordArgs.the_doc;
            using (var document = DocX.Load(wordArgs.wordFile))


            //  using (var document = DocX.Load(@"E:\Repos\WordExcel_net5\bin\Debug\Test.docx"))
            {

                document.SetDefaultFont(new Xceed.Document.NET.Font("Times New Roman"), 14);
                // Create a paragraph and insert text.

                string header = String.Format(wordArgs.shapka[0] + wordArgs.lessonNumbers[counter]);
                document.InsertParagraph(header).Alignment = Alignment.center;
                document.InsertParagraph(wordArgs.fullTopic).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[1]).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[2]).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[3]).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[4]).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[5]).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[6]).Alignment = Alignment.both;
                document.InsertParagraph(wordArgs.shapka[7]).Alignment = Alignment.both;
                document.InsertParagraph("  1. " + wordArgs.books[a] + "\n  2. " + wordArgs.books[b]).Alignment = Alignment.left;
                document.InsertParagraph(wordArgs.shapka[8]).Alignment = Alignment.both;
                var p2 = document.InsertParagraph();
                var t = p2.InsertTableAfterSelf(2, 2);
                document.InsertParagraph("\n");

                //     string[] splittedText, parts;
                List<string> parts = new List<string>();
                List<string> splittedText = new List<string>();
                regexOperations(wordArgs, t, out splittedText, out parts);

                {
                    t.Rows[0].Cells[0].Paragraphs[0].Append("Учебные вопросы и время, отведенное на их рассмотрение");
                    t.Rows[0].Cells[1].Paragraphs[0].Append("Методические рекомендации руководителю учебного занятия");
                    t.Rows[1].Cells[0].Paragraphs[0].Append("Вступительная часть (5 минут)");
                    t.Rows[1].Cells[1].Paragraphs[0].Append("Доклад командира группы о готовности к занятиям. Объявление темы и целей занятий.");
                    
                    for (int i = 0; i < splittedText.Count; i++)
                    {
                        Row dynamicRow = t.InsertRow();
                        dynamicRow.Cells[0].Paragraphs.First().Append(parts[i]);
                        dynamicRow.Cells[0].Paragraphs[0].Append(splittedText[i]);
                        //дальше выбираем, будет ли правый столбец будет браться из внешнего файла, или будет дополняться содержанием правого столбца
                        if (wordArgs.externalContensRight == false) //
                        {
                            dynamicRow.Cells[1].Paragraphs.First().Append
                                                (wordArgs.contents_right[i % wordArgs.contents_right.Length]); // здесь знак "%" возвращает в начало содержания, если в левом столбце больше частей, чем в данном массиве 
                        }
                        else
                        {
                            if (splittedText[i].Contains("активной форме")) 
                            {//Если есть активная форма, то не нужно ничего подставлять в начале
                                
                                dynamicRow.Cells[1].Paragraphs.First().Append(splittedText[i]);
                            }
                            else
                            {
                                dynamicRow.Cells[1].Paragraphs.First().Append
                                                        ("Группа изучает " + char.ToLower(splittedText[i][0]) + splittedText[i].Substring(1));
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
            }
        }

        private static void regexOperations(WordArgs wordArgs, Table t, out List<string> splittedText, out List<string> parts) // это извлечённый метод, поэтому много аргументов. можно было бы поработать нам тем, чтобы все их сделать частями класса
        {

            string pattern2 = @"\n|\;|(?<!форме)\.\s?|(В активной.*)"; //решил не делать 2 этапа разбиения, пусть сразу будет много и тогда если че ужмём
            Regex rgx = new Regex(pattern2);
            splittedText = rgx.Split(wordArgs.clearCell).ToList(); //вот здесь укасывается максимальное количество разбиений. для этого выше создается отдельный объект регекс
            
            //foreach(string s in splittedText) Console.WriteLine("entry: " + s);

            //вот тут что-то связанное с делегатами, хоть я про них и не шарю
            splittedText.RemoveAll(RemoveSpacesFromList); //эта инструкция использует метод из скобок, чтоб найти в каждой строке листа строку, соответствующую условиям return, и удалить такую строку. все просто
            bool RemoveSpacesFromList(string s)
            {
                return (s.Length < 3) && (s.Equals(String.Empty) || s.Contains(" ") || s.Contains(";") || s.Contains("."));
            }
            splittedText.TrimExcess();

            //foreach (string s in splittedText) Console.WriteLine("entry: " + s);

            //снова возвращаем лимитирование, потому что после первого разбиения остаются пустые строки и т.д. и вот после того
            //как они убраны, можно уже ограничивать настоящий текст каким-то максимумом
            int limit = 3; 
            if(splittedText.Count > limit)
            {
                for(int i = limit; i < splittedText.Count;)
                {
                    if (Char.IsLower(splittedText[i][0]))
                    {
                        splittedText[limit - 1] += "; " + splittedText[i];
                    }
                    else
                    {
                        splittedText[limit - 1] += ". " + splittedText[i];
                    }
                    splittedText.RemoveAt(i);
                }
            }

            for (int i = 0; i < splittedText.Count; i++)
            {

                splittedText[i].TrimStart();
                splittedText[i] = char.ToUpper(splittedText[i][0]) + splittedText[i].Substring(1); //делаем каждый учебный вопрос с заглавной
                //TODO(later) шаблон вставки пробелов, когда два слова разделены запятой или точкой без пробела
                //{ //в этом блоке будем дальше украшать ячейку, добавлять пробелы после точек
                //    string pattern = @"\b\p\b";
                //    if(Regex.IsMatch(splittedText[i], pattern)) Console.WriteLine("найдено: " + splittedText[i]);
                    
                //    //string replacement = @"\b\p\s\b";
                //    //string result = Regex.Replace(splittedText[i], pattern, replacement);
                //    //splittedText[i] = result;
                //}
            }


            //for (int i = 0; i < splittedText.Capacity; i++)
            //{
            //    splittedText[i] = splittedText[i].TrimStart(); //?? убираем из начала пробелы??
            //    if (!(Regex.IsMatch(splittedText[i], @"\b\."))) splittedText[i] = $"{splittedText[i]}."; //судя по всему добавляем точку. уже стал забывать
            //}


            ////=================================================================================
            ////=всё, что идёт ниже, это просто жесть))) это я придумал вместо того, чтобы просто
            ////=как раньше разбить на три части с суммой 80. тут минуты прям рассчитываются
            ////=с разными проверками на > или < 80, т.к. apparently я не достаточно умный,
            ////=чтоб сразу приудмать формулу, которая будет упираться в максимум 80 мину.
            ////=================================================================================

            //TODO(actual) минутки то неровно разбиваются, пофиксить
            if (splittedText.Count > 0)
            {
                Random minsRand = new Random();// работаем с таблицей и переменными для разных минут
                List<int> minutes = new List<int>(splittedText.Count);
                int randMax = (80 / splittedText.Count) / 5;

                for (int i = 0; i < splittedText.Count; i++)
                {   // разделить на 5 и умножить на 5 нужно как раз чтобы длительность учебных вопросов была округлена до 5 минут. изи)
                    minutes.Add(80 / splittedText.Count / 5 * 5);
                }

                int allTime;
                int howManyTimes = 0;
                do
                {
                    allTime = 0;
                    foreach (int kk in minutes) allTime += kk;
                    while (allTime > 80)
                    {
                        int fL = findLargest(minutes); // здесь берется нерандомная ячейка, т.к. был случай ухода в минус по минутам))))
                        minutes[fL] -= 5;
                        allTime = 0;
                        foreach (int kk in minutes) allTime += kk;
                        //Console.WriteLine("Перехлёст: " + allTime);
                    }
                    while (allTime < 80)
                    {
                        int randomPart = minsRand.Next(splittedText.Count); // тут можно рандомщину брать, т.к. ниже нуля не уйдем)))))
                        minutes[randomPart] += 5;
                        allTime = 0;
                        foreach (int kk in minutes) allTime += kk;
                        //Console.WriteLine("Недохлёст: " + allTime);
                    }
                    howManyTimes++;
                } while (allTime != 80 && howManyTimes < 10);
                    if (allTime != 80) Console.WriteLine("да блядь это невозможно привести к 80 минутам...");

                parts = new List<string>();
                for (int i = 0; i < splittedText.Count; i++)
                {
                    parts.Add(String.Format("Учебный вопрос № {1}. ({0} минут)\n", minutes[i], (i + 1)));
                }

                return; 
            } //end if

            else
            {
                Console.WriteLine("В темплане отсутствует содержание данного практического занятия");
                parts = new List<string>(splittedText.Capacity);
                for (int i = 0; i < splittedText.Capacity; i++)
                {
                    parts.Add(String.Format("<Учебный вопрос отсутствует>\n"));
                    splittedText.Add(String.Format("<В темплане отсутствует содержание данного практического занятия>"));
                }

                return; 
            }

            int findLargest(List<int> minutes)
            {
                int Largest = 0;
                int returnable = 0;

                for (int i = 0; i < minutes.Capacity; i++)
                {
                    if (minutes[i] > Largest)
                    {
                        Largest = minutes[i];
                        returnable = i;
                    }
                }
                return returnable;
            }
        }
    }
}
