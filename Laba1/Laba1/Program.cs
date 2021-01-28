using Microsoft.Office.Interop.Word;

namespace Laba1
{
    class Program
    {
        
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
                "[*имя раздела*]",///0
                "[*имя рисунка*]",///1
                "[*ссылка на следующий рисунок*]",///2
                "[*ссылка на предыдущий рисунок*]",///3
                "[*ссылка на таблицу*]",///4
                "[*таблица первая*]"///5
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
                            }
                                break;
                            case 2:
                            {
                                  
                            }
                                break;
                            case 3:
                            {
                                    
                            }
                                break;
                            case 4:
                            {
                                  
                            }
                                break;
                            case 5:
                            {

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