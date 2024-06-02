using CreaterFromVSU.ViewModel.Utilites.Structs;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static CreaterFromVSU.ViewModel.Utilites.MicrosoftWordProcessor;

namespace CreaterFromVSU.ViewModel.Utilites
{
    class PDFProcessor
    {
        /*public async void StartCreateDocuments(Create_for is_create_doc)
        {
            await AddTextToTextBoxAsync("Начало создания... (несколько минут будет чтение файлов с данными, после чего запустится создание документов)");
            // участники дистанционные
             List<PlayersListStruct> playersDist = new List<PlayersListStruct>();
             CreateListForDiploms(playersDist, PathOpen, 1);

             // участники очные
             List<PlayersListStruct> playersOchno = new List<PlayersListStruct>();
             CreateListForDiploms(playersOchno, PathOpen, 2);

             // словарь соревнований
             Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences = new Dictionary<string, ReferenceMaterialDictionary>();
             CreateDictionaryReference(dictionaryReferences, PathOpen);

             // словарь городов
             Dictionary<string, string> dictionaryCities = new Dictionary<string, string>();
             CreateDictionaryCity(dictionaryCities, PathOpen);

             city_dist = new Thread(() => CreateCityes(FolderSave, playersDist, dictionaryCities, "города-дистант"));
             city_ochno = new Thread(() => CreateCityes(FolderSave, playersDist, dictionaryCities, "города-очно"));
 

                  diplom_dist = new Thread(() => CreateDiploms("дипломы-дистант", playersDist, dictionaryReferences, dictionaryCities));
                  diplom_ochno = new Thread(() => CreateDiploms("дипломы-очно", playersOchno, dictionaryReferences, dictionaryCities));
                  sertific_dist = new Thread(() => CreateCertificate("сертификаты-дистант", playersDist, dictionaryReferences, dictionaryCities));
                  sertific_ochno = new Thread(() => CreateCertificate("сертификаты-очно", playersOchno, dictionaryReferences, dictionaryCities));
                  sertificFrom_dist = new Thread(() => CreateCertificateWithBacking("сертификаты_с_подложкой-дистант", playersDist, dictionaryReferences, dictionaryCities));
                  sertificFrom_ochno = new Thread(() => CreateCertificateWithBacking("сертификаты_с_подложкой-очно", playersOchno, dictionaryReferences, dictionaryCities));

                  moder_dist = new Thread(() => CreateModer(FolderSave, playersDist, dictionaryReferences, "модерам-дистант"));
                  moder_ochno = new Thread(() => CreateModer(FolderSave, playersOchno, dictionaryReferences, "модерам-очно"));
      
            if (is_create_doc.city_dist)
            {
                city_dist.Start();
            }
            if (is_create_doc.city_ochno)
            {
                city_ochno.Start();
            }
            if (is_create_doc.diplom_dist)
            {
                diplom_dist.Start();
            }
            if (is_create_doc.diplom_ochno)
            {
                diplom_ochno.Start();
            }
            if (is_create_doc.sertific_dist)
            {
                sertific_dist.Start();
            }
            if (is_create_doc.sertific_ochno)
            {
                sertific_ochno.Start();
            }
            if (is_create_doc.sertificFrom_dist)
            {
                sertificFrom_dist.Start();
            }
            if (is_create_doc.sertificFrom_ochno)
            {
                sertificFrom_ochno.Start();
            }
            if (is_create_doc.moder_dist)
            {
                moder_dist.Start();
            }
            if (is_create_doc.moder_ochno)
            {
                moder_ochno.Start();
            }
        }*/
        /*public void CreateCertificateWithBacking(string type, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            CreateWordPage(CreateSertificatWithBacking, type, players, dictionaryReferences, dictionaryCities);
        }*/
        /*public void CreateCertificate(string type, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            CreateWordPage(CreateSertificat, type, players, dictionaryReferences, dictionaryCities);
        }*/
        /*public void CreateDiploms(string type, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            CreateWordPage(CreateDiplom, type, players, dictionaryReferences, dictionaryCities);
        }*/
        /*private void CreateParagrahp(ref Document document, string text, int countParagraph,
                                                int spaceBefor, int spaceAfter)
        {
            Word.Paragraph paragraph = document.Paragraphs.Add();
            Word.Range range = paragraph.Range;
            range.Font.Size = 16;
            range.Font.Name = "Calibri";
            range.Text = text.Trim().Length == 0 ? " " : text.Trim();

            range.InsertParagraphAfter();

            document.Paragraphs[countParagraph].SpaceBefore = spaceBefor;
            document.Paragraphs[countParagraph].SpaceAfter = spaceAfter;
            document.Paragraphs[countParagraph].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }*/
        /*private void CreateDiplom(DiplomStruct diplom, string SavePath)
        {
            var wordApp = new Word.Application();
            // Добавляем новый документ
            Word.Document doc = wordApp.Documents.Add();

            CreateParagrahp(ref doc, diplom.competition, 1, 360, 0);

            CreateParagrahp(ref doc, "возрастная категория", 2, 6, 0);
            CreateParagrahp(ref doc, diplom.age.Substring(20), 3, 0, 0);

            CreateParagrahp(ref doc, diplom.fio, 4, 60, 12);
            doc.Paragraphs[4].Range.Font.Size = 26;
            doc.Paragraphs[4].Range.Font.Bold = 1;
            CreateParagrahp(ref doc, diplom.birthday, 5, 6, 0);
            CreateParagrahp(ref doc, diplom.city, 6, 0, 8);
            CreateParagrahp(ref doc, diplom.teacher, 7, 18, 8);
            doc.Paragraphs[6].Range.Font.Size = 15;

            doc.PageSetup.TopMargin = 0;
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;
            doc.PageSetup.BottomMargin = 0;

            // Сохраняем документ
            doc.SaveAs(SavePath);
            doc.Close();
            wordApp.Quit();
        }*/
        /*private void CreateSertificat(DiplomStruct diplom, string SavePath)
        {
            var wordApp = new Word.Application();
            // Добавляем новый документ
            Word.Document doc = wordApp.Documents.Add();

            CreateParagrahp(ref doc, diplom.fio, 1, 283, 4);
            doc.Paragraphs[1].Range.Font.Size = 26;
            doc.Paragraphs[1].Range.Font.Bold = 1;


            if (diplom.city.Length >= 44)
            {
                CreateParagrahp(ref doc, diplom.birthday, 2, 0, 0);
                string first = "";
                string[] words = diplom.city.Split(' ');
                int i = 0;
                while (words.Length > i && (first + words[i]).Length < 45)
                {
                    first += words[i] + " ";
                    i++;
                }
                string last = "";
                for (int k = i; words.Length > k; k++)
                    last += words[k] + " ";

                CreateParagrahp(ref doc, first, 3, 0, 0);
                CreateParagrahp(ref doc, last, 4, 0, 0);

                CreateParagrahp(ref doc, diplom.competition, 5, 120, 8);
                CreateParagrahp(ref doc, diplom.age, 6, 0, 8);
                CreateParagrahp(ref doc, diplom.teacher, 7, 36, 8);
                doc.Paragraphs[7].Range.Font.Size = 15;
            }
            else
            {
                CreateParagrahp(ref doc, diplom.birthday, 2, 0, 8);
                CreateParagrahp(ref doc, diplom.city, 3, 0, 8);
                CreateParagrahp(ref doc, diplom.competition, 4, 120, 8);
                CreateParagrahp(ref doc, diplom.age, 5, 0, 8);
                CreateParagrahp(ref doc, diplom.teacher, 6, 36, 8);
                doc.Paragraphs[6].Range.Font.Size = 15;
            }


            doc.PageSetup.TopMargin = 0;
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;
            doc.PageSetup.BottomMargin = 0;

            // Сохраняем документ
            doc.SaveAs(SavePath);
            doc.Close();
            wordApp.Quit();
        }*/
        /* private void CreateSertificatWithBacking(DiplomStruct diplom, string SavePath)
        {
            var wordApp = new Word.Application();
            // Добавляем новый документ
            Word.Document doc = wordApp.Documents.Add();

            CreateParagrahp(ref doc, diplom.fio, 1, 283, 4);
            doc.Paragraphs[1].Range.Font.Size = 26;
            doc.Paragraphs[1].Range.Font.Bold = 1;
            if (diplom.city.Length >= 44)
            {
                CreateParagrahp(ref doc, diplom.birthday, 2, 0, 0);
                string first = "";
                string[] words = diplom.city.Split(' ');
                int i = 0;
                while (words.Length > i && (first + words[i]).Length < 45)
                {
                    first += words[i] + " ";
                    i++;
                }
                string last = "";
                for (int k = i; words.Length > k; k++)
                    last += words[k] + " ";

                CreateParagrahp(ref doc, first, 3, 0, 0);
                CreateParagrahp(ref doc, last, 4, 0, 0);

                CreateParagrahp(ref doc, diplom.competition, 5, 120, 8);
                CreateParagrahp(ref doc, diplom.age, 6, 0, 8);
                CreateParagrahp(ref doc, diplom.teacher, 7, 36, 8);
                doc.Paragraphs[7].Range.Font.Size = 15;
            }
            else
            {
                CreateParagrahp(ref doc, diplom.birthday, 2, 0, 8);
                CreateParagrahp(ref doc, diplom.city, 3, 0, 8);
                CreateParagrahp(ref doc, diplom.competition, 4, 120, 8);
                CreateParagrahp(ref doc, diplom.age, 5, 0, 8);
                CreateParagrahp(ref doc, diplom.teacher, 6, 36, 8);
                doc.Paragraphs[6].Range.Font.Size = 15;
            }

            doc.PageSetup.TopMargin = 0;
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;
            doc.PageSetup.BottomMargin = 0;


            Word.Shape shape = doc.Shapes.AddPicture(PathPicture, false, true, 0, 0, 0, 0);
            shape.Fill.UserPicture(PathPicture);
            shape.Width = doc.PageSetup.PageWidth;
            shape.Height = doc.PageSetup.PageHeight;
            shape.Top = 0;
            shape.Left = 0;

            shape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;

            // Сохраняем документ
            doc.SaveAs(SavePath);
            doc.Close();
            wordApp.Quit();
        }*/
        /*private async void CreateWordPage(DiplomCreationDelegate creationDelegate, string folderMain, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            try
            {
                int i = 0;
                string currentPath;
                DiplomStruct diplomStruct = new DiplomStruct();
                while (i < players.Count)
                {
                    PlayersListStruct people = players[i];
                    await AddTextToTextBoxAsync(i.ToString().PadRight(4, ' ') + " | " + people.ToString().PadRight(50, ' ') + " | " + folderMain);
                    if ((!string.IsNullOrEmpty(people.CodeCompetition) && !people.CodeCompetition.Contains("не участвую"))
                        || (!string.IsNullOrEmpty(people.CodeContest) && !people.CodeContest.Contains("не участвую"))
                        || (!string.IsNullOrEmpty(people.CodeExhibition) && !people.CodeExhibition.Contains("не участвую"))
                        || (!String.IsNullOrEmpty(people.OlympicsContest) && !people.OlympicsContest.Contains("не участвую")))
                    {
                        currentPath = FolderSave + @"\" + folderMain + @"\" + people.CityPlayers + @"\" + people.FioPlayers;
                        if (!Directory.Exists(FolderSave + @"\" + folderMain))
                        {
                            // Создаем папку, если она не существует
                            Directory.CreateDirectory(FolderSave + @"\" + folderMain);
                        }
                        if (!Directory.Exists(FolderSave + @"\" + folderMain + @"\" + people.CityPlayers))
                        {
                            // Создаем папку, если она не существует
                            Directory.CreateDirectory(FolderSave + @"\" + folderMain + @"\" + people.CityPlayers);
                        }

                        diplomStruct.fio = (people.FioPlayers.Split(' ').Length > 0 ? people.FioPlayers.Split(' ')[0] : "") + " " + (people.FioPlayers.Split(' ').Length > 1 ? people.FioPlayers.Split(' ')[1] : "");

                        if (people.SchoolPlayers == "Индивидуальный участник")
                        {
                            diplomStruct.birthday = "";
                        }
                        else
                        {
                            diplomStruct.birthday = String.Format(SCHOOL, people.isMen ? "учащийся" : "учащаяся", people.SchoolPlayers);
                        }
                        diplomStruct.city = dictionaryCities[people.CityPlayers];
                        diplomStruct.teacher = String.Format(TEACHER, people.TeacherPlayers);

                        if (!String.IsNullOrEmpty(people.CodeCompetition) && !people.CodeCompetition.Contains("не участвую"))
                        {
                            diplomStruct.competition = String.Format(COMPETITION, "соревновании", dictionaryReferences[people.CodeCompetition].NameCompetition);
                            diplomStruct.age = String.Format(AGE, dictionaryReferences[people.CodeCompetition].AgeRank);
                            creationDelegate(diplomStruct, currentPath + people.CodeCompetition);
                        }
                        if (!String.IsNullOrEmpty(people.CodeContest) && !people.CodeContest.Contains("не участвую"))
                        {
                            diplomStruct.competition = String.Format(COMPETITION, "конкурсе", dictionaryReferences[people.CodeContest].NameCompetition);
                            diplomStruct.age = String.Format(AGE, dictionaryReferences[people.CodeContest].AgeRank);
                            creationDelegate(diplomStruct, currentPath + people.CodeContest);
                        }
                        if (!String.IsNullOrEmpty(people.CodeExhibition) && !people.CodeExhibition.Contains("не участвую"))
                        {
                            diplomStruct.competition = String.Format(COMPETITION, "выставке", dictionaryReferences[people.CodeExhibition].NameCompetition);
                            diplomStruct.age = String.Format(AGE, dictionaryReferences[people.CodeExhibition].AgeRank);
                            creationDelegate(diplomStruct, currentPath + people.CodeExhibition);
                        }
                        if (!String.IsNullOrEmpty(people.OlympicsContest) && !people.OlympicsContest.Contains("не участвую"))
                        {
                            diplomStruct.competition = String.Format(OLIMPIC, "олимпиаде", dictionaryReferences[people.OlympicsContest].NameCompetition);
                            diplomStruct.age = String.Format(AGE, dictionaryReferences[people.OlympicsContest].AgeRank);
                            creationDelegate(diplomStruct, currentPath + people.OlympicsContest);
                        }

                    }
                    i++;
                }
            }
            catch (Exception e)
            {
                await AddTextToTextBoxAsync(String.Format("Произошла ошибка при создании {0}!\n Ошибка: ", folderMain) + e);
            }
        }*/
    }
}
