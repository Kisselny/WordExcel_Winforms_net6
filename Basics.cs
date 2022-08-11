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
using Xceed.Words.NET;

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
	static public class Basics  // TODO подгрузка элементов таблицы из БД, UI, разбивка на документы, ✅конвертировать массивы в List<T>✅, (если 1 учебный вопрос сильно маленький, переразбить ячейку)
	{                       // ✅переделать (ебейшее число аргументов и глобальный треш) в структуру, которую можно передавать функции Word🆗
							// подумать, можно ли открыть файл один раз вне цикла, а в открытый док всё добавлять
		[System.STAThread] // вот эта фигня была нужна, иначе не компилировалось

        internal static void source_XML_Word(WordArgs wordArgs) //BUG: в тестовом исходном ворд-доке на теме 12 не пронумерованы ПЗ. Из-за этого данных ПЗ ващщщще нет в итоговом файле, хотя есть метод, который нумерует даже непронумерованные ПЗ
		{
			// Open a WordprocessingDocument for editing using the filepath.
			using (WordprocessingDocument src_docx =
				WordprocessingDocument.Open(wordArgs.sourceFile, true))
			{
				//Find the  table in the document.
					  wordArgs.wordTable_Global =
						  src_docx.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(1);
					//TODO имеет смысл добавить трай-кетч на наличие таблицы
			   
					int row_count = wordArgs.wordTable_Global.Elements<TableRow>().Count();
					Console.WriteLine("строк: " + row_count);
					TableRow row; //объявляем эти штуки здесь, чтобы не внутри трай-кетча
					TableCell cell;

					for (int i = 0; i < row_count; i++)
					{                
						try
						{
							row = wordArgs.wordTable_Global.Elements<TableRow>().ElementAt(i);// Find the second row in the table.
							cell = row.Elements<TableCell>().ElementAt(1);// Find the third cell in the row.
							wordArgs.clearCell = row.Elements<TableCell>().ElementAt(2).InnerText;

						}
						catch (ArgumentOutOfRangeException)
						{
							continue;
						}

						if (cell.InnerText.ToString().Contains("семестр")) wordArgs.semester = int.Parse(Regex.Match(cell.InnerText.ToString(), @"\d+").Value);

						if (cell.InnerText.ToString().Contains("Тема"))
						{
							wordArgs.topicNow = int.Parse(Regex.Match(cell.InnerText.ToString(), @"\d+").Value);
							wordArgs.fullTopic = String.Format("по теме № {0}. {1};", wordArgs.topicNow, wordArgs.clearCell);
						}

						if (cell.InnerText.ToString().Contains("Практическое занятие"))
							{
								wordArgs.lessonNumbers = LessonNumbers(cell.InnerText, wordArgs, i);

								for (wordArgs.counter = 0; wordArgs.counter < wordArgs.lessonNumbers.Length; wordArgs.counter++)
								{
									wordArgs.WordBuild(wordArgs); //здесь мы вызываем главный метод, чтоб строить результирующий документ. в методе нужно заменить переменную номера занятия на массив как для ворда так и для экселя
								}
							}
						}
						return;
					//Paragraph p = cell.Elements<Paragraph>().FirstOrDefault();
					//Run r = p.Elements<Run>().FirstOrDefault();
					//Text t = r.Elements<Text>().FirstOrDefault();

					//Console.WriteLine(t.Text);  
			   // src_docx.Save();
			}
		}

        private static int[] LessonNumbers(string celltext, WordArgs wordArgs, int iterOut)
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
                    {	// эта функция возвращает наибольшее количество часов в строке, что чаще всего соответствует колонке "практические занятия"
                        missedNumbers.Add(wordArgs.searchInSource(i, iterOut));
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


        internal static void source_is_Excel(WordArgs wordArgs)
		{
			using (ExcelPackage package = new ExcelPackage(new FileInfo(wordArgs.sourceFile)))
			{
				ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
				var sheet = package.Workbook.Worksheets["Лист2"];
				wordArgs.excelSheet_Global = package.Workbook.Worksheets["Лист2"];
				//  using (wordArgs.the_doc = DocX.Load(wordArgs.wordFile)) // попытка использовать .Load вне цикла: сохраняется только 1й ппз. только он
				{
					for (int k = 8; k <= 523; k++) // last working value = 89 / 523  // <- не забыть вернуть полный цикл
					{
						string cell = "B" + k.ToString(); // эти две обязательно находятся внутри цикла
						string exCell = wordArgs.excelSheet_Global.Cells[cell].Value.ToString();



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
							for (wordArgs.counter = 0; wordArgs.counter < wordArgs.lessonNumbers.Length; wordArgs.counter++)
							{	// теперь тут еще и вложенный луп, чтобы экселевский метод работал, как вордовский, используя только массив номеров, а не отдельный номер
								wordArgs.WordBuild(wordArgs); // самое важное - вызов функции внутри лупа. заменил кучу аргументов на объект
															  //  wordArgs.the_doc.Save();
								Console.WriteLine(wordArgs.lessonNumbers[wordArgs.counter] + " done");
							}
						}

					}
					MessageBox.Show("Готово!");
				}
			}
		}
	}
}
