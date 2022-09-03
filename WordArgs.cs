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
        public string exCell;
        public string clearCell;
        public int lessonNow, topicNow; // это когда в ворде (или не только) номер следующего занятия не вписан
        public (int lastTopic, int lastLesson) scanLast; // попробуем prevLessonNum заменить на это, чтобы у предыдущего номера занятия был ассоциированный номер
        public int[] lessonNumbers;
        List<int> matches;
        public List<string> parts = new List<string>();
        public List<string> splittedText = new List<string>();
        
        public Xceed.Document.NET.Document the_doc; // это тип документа в var document = DocX.Load(wordArgs.wordFile)
                                                    //он нужен чтобы передать .Load вне цикла внутрь метода Word(), но почему-то сохраняется только  1й ппз


        public void regexOperations(WordArgs wordArgs) // это извлечённый метод, поэтому много аргументов. можно было бы поработать нам тем, чтобы все их сделать частями класса
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
