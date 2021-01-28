using System;
using Microsoft.Office.Interop.Word;

namespace Laba1
{
    class Program
    {
        static WdColorIndex wdNoHighlight;

        /// <summary>
        /// номера раздела, ==0 - нет разделов
        /// </summary>
        static uint _sectionNumber = 0;

        /// <summary>
        /// номера рисунков, ==0 - нет картинок
        /// </summary>
        static uint _pictureNumber = 0;

        /// <summary>
        /// номера таблиц, ==0 - нет таблиц
        /// </summary>
        static uint _tableNumber = 0;

        static void Main(string[] args)
        {
            //путь до исходного шаблона
            string sourcePath = @"C:\Users\User\Desktop\LabRTF\шаблон.rtf";

            //путь до выходного файла
            string distPath = @"C:\Users\User\Desktop\LabRTF\result.rtf";

            //путь до csv файла для создания таблицы
            string csvPath = @"C:\Users\User\Desktop\LabRTF\data.csv";

            //список закладок
            string[] templateStringList =
            {
                "[*имя раздела*]", ///0
                "[*имя рисунка*]", ///1
                "[*ссылка на следующий рисунок*]", ///2
                "[*ссылка на предыдущий рисунок*]", ///3
                "[*ссылка на таблицу*]", ///4
                "[*таблица первая*]" ///5
            };

            var application = new Application();
            application.Visible = true;

            var document = application.Documents.Open(sourcePath);

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                for (int i = 0; i < templateStringList.Length; i++)
                {
                    if (paragraph.Range.Text.Contains(templateStringList[i]))
                    {
                        System.Console.Out.WriteLine(templateStringList[i]);
                    }
                }
            }

            Paragraph prevParagraph = null;
            Object missing = System.Type.Missing;

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                for (int i = 0; i < templateStringList.Length; i++)
                {
                    if (paragraph.Range.Text.Contains(templateStringList[i]))
                    {
                        switch (i)
                        {
                            case 0:
                            {
                                paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                paragraph.Range.Font.Name = "Times New Roman";
                                paragraph.Range.Font.Size = 15;
                                paragraph.Format.SpaceAfter = 12;
                                paragraph.Range.Font.Bold = 1;
                                paragraph.Range.HighlightColorIndex = 0;
                                paragraph.Range.HighlightColorIndex = wdNoHighlight;

                                _sectionNumber++;
                                string replaceString = _sectionNumber.ToString();

                                paragraph.Range.Find.Execute(templateStringList[i],
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2,
                                    ref missing, ref missing,
                                    ref missing, ref missing);
                            }
                                break;
                            case 1:
                            {
                                paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                paragraph.Range.Font.Name = "Times New Roman";
                                paragraph.Range.Font.Size = 12;
                                paragraph.Format.SpaceAfter = 12;
                                paragraph.Range.HighlightColorIndex = 0;

                                if (prevParagraph != null)
                                {
                                    prevParagraph.Format.SpaceBefore = 12;
                                    prevParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                }

                                _pictureNumber++;
                                var replaceString = $"Рисунок {_sectionNumber}.{_pictureNumber} -";

                                paragraph.Range.Find.Execute(templateStringList[i],
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2,
                                    ref missing, ref missing,
                                    ref missing, ref missing);
                            }
                                break;
                            case 2:
                            {
                                var replaceString = _sectionNumber + "." + (_pictureNumber + 1);

                                paragraph.Range.Find.Execute(templateStringList[i],
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2,
                                    ref missing, ref missing,
                                    ref missing, ref missing);
                                
                                paragraph.Range.HighlightColorIndex = wdNoHighlight;
                            }
                                break;
                            case 3:
                            {
                                var replaceString = _sectionNumber + "." + _pictureNumber;

                                paragraph.Range.Find.Execute(templateStringList[i],
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2,
                                    ref missing, ref missing,
                                    ref missing, ref missing);
                                
                                paragraph.Range.HighlightColorIndex = wdNoHighlight;
                            }
                                break;
                            case 4:
                            {
                                _tableNumber++;
                                var replaceString = _sectionNumber + "." + _tableNumber;

                                paragraph.Range.Find.Execute(templateStringList[i],
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2,
                                    ref missing, ref missing,
                                    ref missing, ref missing);
                                
                                paragraph.Range.HighlightColorIndex = wdNoHighlight;
                            }
                                break;
                            case 5:
                            {
                                application.Selection.Find.Execute(templateStringList[i]);
                                var range = application.Selection.Range;
                                range.HighlightColorIndex = 0;

                                string[] 
                                    listRows=System.IO.File.ReadAllText(csvPath).Split("\r\n".ToCharArray(),
                                        StringSplitOptions.RemoveEmptyEntries);

                                string[] 
                                    listTitle=listRows[0].Split(";,".ToCharArray(), 
                                        StringSplitOptions.RemoveEmptyEntries);

                                var wordTable = document.Tables.Add(range, listRows.Length, 
                                    listTitle.Length);

                                for (var k = 0; k < listTitle.Length; k++)
                                {
                                    wordTable.Cell( 1, k + 1).Range.Text = listTitle[k].ToString();
                                }

                                for (var j = 1; j < listRows.Length; j++)
                                {
                                    string[] 
                                        listValues = listRows[j].Split(";,".ToCharArray(),
                                            StringSplitOptions.RemoveEmptyEntries);
                                    for (var k = 0; k < listValues.Length; k++)
                                    {
                                        wordTable.Cell(j + 1, k + 1).Range.Text = listValues[k].ToString();
                                    }
                                }
                            }
                                break;

                        }
                    }
                }

                prevParagraph = paragraph;
            }

            document.SaveAs2(distPath);
            System.Console.In.Read();
            // application.Quit();
        }
    }
}