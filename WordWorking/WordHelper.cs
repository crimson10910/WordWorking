using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordsChanger
{
    class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string fileName)  // Конструктор класса с защитой 
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not found");
            }
        }

        internal bool Process(Dictionary<string, string> items, bool showPreview)   //Основной метод работы с документом формата Word
        {
            Word.Application app = null;    
            try
            {
                app = new Word.Application();   // вызываем редактор ворда

                string newFileName = Path.Combine(_fileInfo.DirectoryName, 
                    DateTime.Now.ToString("yyyyMMdd HHmmss ") + _fileInfo.Name); //Задаем новое имя для файла
                File.Copy(_fileInfo.FullName, newFileName);     //копируем старый файл

                Object file = newFileName;  // создаем объект файла

                Object missing = Type.Missing;

                app.Documents.Open(file);   //Открываем документ

                foreach (var item in items) // далее работаем со словарем, который направили в этот метод из
                                            // предыдущего файла программы
                {
                    Word.Find find = app.Selection.Find;    //обозначаем, что будем искать и выделять некий фрагмент
                    find.Text = item.Key;                   //обозначаем, что выделенным фрагментом будет ключ словаря
                    find.Replacement.Text = item.Value;     // замененный текст будет содержимым из словаря под определенным ключем

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,    // это настройки поиска, вы можете сами в                    
                        MatchCase: false,                   // них разобраться и подобрать под конкретную задачу
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

                app.ActiveDocument.Save();

                if (showPreview)
                {
                    app.Visible = true;
                    app.ActiveDocument.PrintPreview();
                }
                else
                {
                    app.ActiveDocument.Close();
                }

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (app != null && !showPreview)
                {
                    app.Quit();
                }
            }

            return false;
        }
    }
}