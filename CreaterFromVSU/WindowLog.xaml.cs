﻿using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Reflection;


namespace CreaterFromVSU
{
    public struct ReferenceMaterialDictionary
    {
        public string TypeCompetition;
        public string AgeRank;
        public string NameCompetition;
    }
    public struct DiplomStruct
    {
        public string competition;
        public string age;
        public string fio;
        public string birthday;
        public string city;
        public string teacher;
    }

    public struct PlayersListStruct
    {
        public string? CodeCompetition;
        public string? NameCommand;
        public string? eMail;
        public string? CodeExhibition;
        public string? CodeContest;
        public string? OlympicsContest;
        public string FioPlayers;
        public DateTime BirthdayPlayers;
        public bool isMen;
        public string SchoolPlayers;
        public string? CityPlayers;
        public string? TeacherPlayers;

        public override string ToString()
        {
            return FioPlayers.PadRight(35, ' ') + " | " + CityPlayers;
        }
    }
    public partial class WindowLog : Window
    {
        public Thread city_dist;
        public Thread city_ochno;
        public Thread diplom_dist;
        public Thread diplom_ochno;
        public Thread sertific_dist;
        public Thread sertific_ochno;
        public Thread sertificFrom_dist;
        public Thread sertificFrom_ochno;
        public Thread moder_dist;
        public Thread moder_ochno;

        private const string COMPETITION = "в {0} «{1}»,";
        private const string OLIMPIC = "в {0} {1},";
        private const string AGE = "возрастная категория «{0}»";
        private const string SCHOOL = "{0} {1}";
        private const string BIRTHDAY_SCHOOL = "{0} {1} {2}";
        private const string TEACHER = "Педагог: {0}";

        private delegate void DiplomCreationDelegate(DiplomStruct diplom, string savePath);

        public struct Create_for
        {
            public bool city_dist;
            public bool city_ochno;
            public bool diplom_dist;
            public bool diplom_ochno;
            public bool sertific_dist;
            public bool sertific_ochno;
            public bool sertificFrom_dist;
            public bool sertificFrom_ochno;
            public bool moder_dist;
            public bool moder_ochno;
        }
        private Thread createrDoc;

        private string PathOpen = "";
        private string FolderSave = "";
        private string PathPicture = "";
        public WindowLog(Create_for cr, string open, string picture, string folder)
        {
            PathPicture = picture;
            PathOpen = open;
            FolderSave = folder;
            InitializeComponent();
            TextBoxLog.Text = "Процесс создания запущен!\n";
            createrDoc = new Thread(() => StartCreateDocuments(cr));
            createrDoc.Start();
        }

        public async Task AddTextToTextBoxAsync(string text)
        {
            await Task.Run(() =>
            {
                lock (locker)
                {
                    Dispatcher.Invoke(() =>
                    {
                        TextBoxLog.AppendText(text + Environment.NewLine);
                    });
                }
            });
        }
        private object locker = new object();

        public async void StartCreateDocuments(Create_for is_create_doc)
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
        }

        public async void CreateModer(string foldelPathOut,
            List<PlayersListStruct> playerList, Dictionary<string, ReferenceMaterialDictionary> referencesDic, string path)
        {
            try
            {
                int k = 0;
                foreach (KeyValuePair<string, ReferenceMaterialDictionary> kvpair in referencesDic)
                {
                    Console.WriteLine("Создание для модераторов файла: " + kvpair.Value.TypeCompetition + " " + kvpair.Value.NameCompetition);
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = workbook.ActiveSheet;

                    // Объединяем ячейки A1 и B1
                    Excel.Range range = worksheet.Range["A1", "I1"];
                    range.Merge();
                    range.Value = "X Межрегиональный открытый фестиваль научно-технического творчества «РОБОАРТ-2024»";
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    range = worksheet.Range["A2", "I2"];
                    range.Merge();
                    range.Value = "Список участников";
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    range = worksheet.Range["A3", "I3"];
                    range.Merge();
                    range.Value = kvpair.Value.TypeCompetition + " " + kvpair.Value.NameCompetition + ", код " + kvpair.Key;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    range = worksheet.Range["A4", "I4"];
                    range.Merge();
                    range.Value = kvpair.Value.AgeRank;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    range = worksheet.Range["A6"]; range.Value = "№ П/П";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["B6"]; range.Value = "Название команды";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["C6"]; range.Value = "ФИО участника";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["D6"]; range.Value = "Дата рождения";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["E6"]; range.Value = "Учебное заведение";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["F6"]; range.Value = "ФИО руководителя";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["G6"]; range.Value = "Населенный пункт";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["H6"]; range.Value = "Отметка о прибытии";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range = worksheet.Range["I6"]; range.Value = "e-mail";
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;


                    int i = 6;
                    foreach (PlayersListStruct people in playerList)
                    {
                        if ((people.CodeContest is not null ? people.CodeContest.Contains(kvpair.Key) : false)
                            || (people.CodeCompetition is not null ? people.CodeCompetition.Contains(kvpair.Key) : false)
                            || (people.CodeExhibition is not null ? people.CodeExhibition.Contains(kvpair.Key) : false)
                            || (people.OlympicsContest is not null ? people.OlympicsContest.Contains(kvpair.Key) : false))
                        {
                            i++;
                            range = worksheet.Range["A" + i.ToString()]; range.Value = (i - 6).ToString();
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["B" + i.ToString()]; range.Value = people.NameCommand;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["C" + i.ToString()]; range.Value = people.FioPlayers;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["D" + i.ToString()]; range.Value = people.BirthdayPlayers;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["E" + i.ToString()]; range.Value = people.SchoolPlayers;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["F" + i.ToString()]; range.Value = people.TeacherPlayers;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["G" + i.ToString()]; range.Value = people.CityPlayers;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["H" + i.ToString()];
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range = worksheet.Range["I" + i.ToString()]; range.Value = people.eMail;
                            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                    }

                    Excel.Range columnB = worksheet.Columns["A"];
                    columnB.ColumnWidth = 6;
                    columnB = worksheet.Columns["B"];
                    columnB.ColumnWidth = 15;
                    columnB = worksheet.Columns["C"];
                    columnB.ColumnWidth = 33;
                    columnB = worksheet.Columns["D"];
                    columnB.ColumnWidth = 13.5;
                    columnB = worksheet.Columns["E"];
                    columnB.ColumnWidth = 25;
                    columnB = worksheet.Columns["F"];
                    columnB.ColumnWidth = 33;
                    columnB = worksheet.Columns["G"];
                    columnB.ColumnWidth = 16;
                    columnB = worksheet.Columns["H"];
                    columnB.ColumnWidth = 18;
                    columnB = worksheet.Columns["I"];
                    columnB.ColumnWidth = 30;

                    if (!Directory.Exists(foldelPathOut + @"\" + path))
                    {
                        Directory.CreateDirectory(foldelPathOut + @"\" + path);
                    }
                    workbook.SaveAs(foldelPathOut + @"\" + path + @"\" + kvpair.Key.Replace(@"\", "").Replace("\"", ""));
                    workbook.Close();
                    excelApp.Quit();
                    await AddTextToTextBoxAsync(("Создан файл: " + kvpair.Value.TypeCompetition.Replace(@"\", "").Replace("\"", "") + " " + kvpair.Value.NameCompetition.Replace(@"\", "").Replace("\"", "") + " -> " + (i - 6).ToString() + " участников") + " | " + path);
                    k = k + i - 6;
                }
                await AddTextToTextBoxAsync("Файлы " + path + " созданы, кол-во участников: " + k.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void CreateCertificateWithBacking(string type, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            CreateWord(CreateSertificatWithBacking, type, players, dictionaryReferences, dictionaryCities);
        }
        public void CreateCertificate(string type, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            CreateWord(CreateSertificat, type, players, dictionaryReferences, dictionaryCities);
        }
        public void CreateDiploms(string type, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
                                        Dictionary<string, string> dictionaryCities)
        {
            CreateWord(CreateDiplom, type, players, dictionaryReferences, dictionaryCities);
        }
        private void CreateParagrahp(ref Word.Document document, string text, int countParagraph,
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
        }
        private void CreateDiplom(DiplomStruct diplom, string SavePath)
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
        }
        private void CreateSertificat(DiplomStruct diplom, string SavePath)
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
        }

        private void CreateSertificatWithBacking(DiplomStruct diplom, string SavePath)
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
        }

        private async void CreateWord(DiplomCreationDelegate creationDelegate, string folderMain, List<PlayersListStruct> players, Dictionary<string, ReferenceMaterialDictionary> dictionaryReferences,
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
        }
        public async void CreateCityes(string foldelPathOut, List<PlayersListStruct> playerList, Dictionary<string, string> citiesDic, string path)
        {
            try
            {
                Dictionary<string, HashSet<string>> schools = new Dictionary<string, HashSet<string>>();
                foreach (KeyValuePair<string, string> kvpair in citiesDic)
                {
                    schools.Add(kvpair.Key, new HashSet<string>());
                    foreach (PlayersListStruct people in playerList)
                    {
                        if (kvpair.Key.Equals(people.CityPlayers))
                        {
                            schools[kvpair.Key].Add(people.SchoolPlayers);
                        }
                    }
                }

                int k = 0;
                foreach (KeyValuePair<string, string> kvpair in citiesDic)
                {
                    foreach (string currentSchool in schools[kvpair.Key])
                    {
                        await AddTextToTextBoxAsync("Создание файла городу: " + kvpair.Key);
                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook workbook = excelApp.Workbooks.Add();
                        Excel.Worksheet worksheet = workbook.ActiveSheet;

                        // Объединяем ячейки A1 и B1
                        Excel.Range range = worksheet.Range["A1", "J1"];
                        range.Merge();
                        range.Value = "X Межрегиональный открытый фестиваль научно-технического творчества «РОБОАРТ-2024»";
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        range = worksheet.Range["A2", "J2"];
                        range.Merge();
                        range.Value = "Лист первичной регистрации участников";
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        range = worksheet.Range["A3", "J3"];
                        range.Merge();
                        range.Value = kvpair.Value; range.Font.Bold = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        range = worksheet.Range["A4", "J4"];
                        range.Merge();
                        range.Value = currentSchool; range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                        Excel.Range row1 = worksheet.Rows[6];
                        row1.RowHeight = 28.8;

                        range = worksheet.Range["A6", "A7"]; range.Merge(); range.Value = "№ П/П";
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["B6", "B7"]; range.Merge(); range.Value = "Фамилия, Имя, Отчество участника";
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["C6", "C7"]; range.Merge(); range.Value = "Фамилия, Имя, Отчество руководителя команды "; range.WrapText = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["D6", "D7"]; range.Merge(); range.Value = "Согласие на обработку персональных данных участника"; range.WrapText = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["E6", "E7"]; range.Merge(); range.Value = "Согласие на обработку персональных данных руководителя"; range.WrapText = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["F6", "F7"]; range.Merge(); range.Value = "Приказ или расписка ответственности"; range.WrapText = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["G6", "H6"]; range.Merge(); range.Value = "Талоны на питание"; range.WrapText = true;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["H7"]; range.Value = "завтрак";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["G7"]; range.Value = "обед";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["I6", "I7"]; range.Merge(); range.Value = "Значки";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range = worksheet.Range["J6", "J7"]; range.Merge(); range.Value = "Подпись педагога";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;


                        int i = 7;
                        HashSet<string> swap = new HashSet<string>();
                        foreach (PlayersListStruct people in playerList)
                        {
                            if (people.CityPlayers.Equals(kvpair.Key) && people.SchoolPlayers.Equals(currentSchool) && (!swap.Contains(people.FioPlayers)))
                            {
                                swap.Add(people.FioPlayers);
                                i++;
                                range = worksheet.Range["A" + i.ToString()]; range.Value = (i - 7).ToString();
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["B" + i.ToString()]; range.Value = people.FioPlayers;
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["C" + i.ToString()]; range.Value = people.TeacherPlayers;
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["D" + i.ToString()];
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["E" + i.ToString()];
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["F" + i.ToString()];
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["G" + i.ToString()]; range.Value = "1";
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["H" + i.ToString()]; range.Value = "1";
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["I" + i.ToString()];
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range = worksheet.Range["J" + i.ToString()];
                                range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            }
                        }
                        k = k + i - 7;
                        i++;

                        range = worksheet.Range["A" + i.ToString(), "J" + i.ToString()];
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.Weight = Excel.XlBorderWeight.xlThin;

                        range = worksheet.Range["B" + i.ToString()]; range.Value = "Руководитель команды";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["G" + i.ToString()]; range.Value = "1";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["H" + i.ToString()]; range.Value = "1";
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;


                        i++;

                        range = worksheet.Range["A" + i.ToString(), "J" + i.ToString()];
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.Weight = Excel.XlBorderWeight.xlThin;

                        range = worksheet.Range["B" + i.ToString()]; range.Value = "Итого"; range.Font.Bold = true;
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["G" + i.ToString()]; range.Value = (i - 8).ToString();
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["H" + i.ToString()]; range.Value = range.Value = (i - 8).ToString();
                        range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        i = i + 1;
                        range = worksheet.Range["A" + i.ToString(), "J" + i.ToString()];
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.Weight = Excel.XlBorderWeight.xlThin;

                        i = i + 3;
                        range = worksheet.Range["B" + i.ToString()]; range.Value = "Сведения заполнил";

                        range = worksheet.Range["C" + i.ToString(), "E" + i.ToString()]; range.Merge();
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        range = worksheet.Range["J" + i.ToString(), "I" + i.ToString()]; range.Merge();
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        i = i + 1;
                        range = worksheet.Range["C" + i.ToString(), "E" + i.ToString()]; range.Merge(); range.Value = "ФИО";
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        range = worksheet.Range["J" + i.ToString(), "I" + i.ToString()]; range.Merge(); range.Value = "Подпись";
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                        Excel.Range columnB = worksheet.Columns["A"];
                        columnB.ColumnWidth = 6;
                        columnB = worksheet.Columns["B"];
                        columnB.ColumnWidth = 33;
                        columnB = worksheet.Columns["C"];
                        columnB.ColumnWidth = 33;
                        columnB = worksheet.Columns["D"];
                        columnB.ColumnWidth = 21;
                        columnB = worksheet.Columns["E"];
                        columnB.ColumnWidth = 21;
                        columnB = worksheet.Columns["F"];
                        columnB.ColumnWidth = 15;
                        columnB = worksheet.Columns["G"];
                        columnB.ColumnWidth = 7;
                        columnB = worksheet.Columns["H"];
                        columnB.ColumnWidth = 7;
                        columnB = worksheet.Columns["I"];
                        columnB.ColumnWidth = 7;
                        columnB = worksheet.Columns["J"];
                        columnB.ColumnWidth = 16;


                        if (!Directory.Exists(foldelPathOut + @"\" + path))
                        {
                            Directory.CreateDirectory(foldelPathOut + @"\" + path);
                        }
                        workbook.SaveAs(foldelPathOut + @"\" + path + @"\" + kvpair.Key.Replace(@"\", "").Replace("\"", "") + " " + currentSchool.Replace(@"\", "").Replace("\"", "") + @".xlsx");
                        workbook.Close();
                        excelApp.Quit();
                        await AddTextToTextBoxAsync(("Создан файл: " + kvpair.Value + " -> " + (i - 6).ToString() + " участников") + " | " + path);

                    }
                }
                await AddTextToTextBoxAsync("Всего: " + k.ToString() + " участников\n Успешное создание всех файлов - " + path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public async void CreateDictionaryCity(Dictionary<string, string> dictionary, string filePath)
        {
            int i = 0;
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet excelTable = excelWorkbook.Sheets[3];

                for (i = 1; excelTable.Cells[i, 2].Value is not null; i++)
                {
                    dictionary.Add(excelTable.Cells[i, 2].Value.ToString(),
                            excelTable.Cells[i, 3].Value.ToString());
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
            catch (Exception e)
            {
                await AddTextToTextBoxAsync("Произошла ошибка с файлом: " + filePath + " \n" + e);
                if (i == 0)
                {
                    await AddTextToTextBoxAsync("Возможно он используется или не правильно составлен. Попробуйте закрыть файл и перезапустить программу!");
                }
                else
                {
                    await AddTextToTextBoxAsync("Cтрока " + i.ToString() + " содержит неполные данные. Удалите или добавьте данные!");
                }
            }
        }
        public async void CreateDictionaryReference(Dictionary<string, ReferenceMaterialDictionary> dictionary, string filePath)
        {
            int i = 0;
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet excelTable = excelWorkbook.Sheets[4];

                for (i = 1; excelTable.Cells[i, 1].Value is not null; i++)
                {
                    dictionary.Add(excelTable.Cells[i, 1].Value.ToString(),
                        new ReferenceMaterialDictionary()
                        {
                            TypeCompetition = excelTable.Cells[i, 2].Value.ToString(),
                            AgeRank = excelTable.Cells[i, 3].Value.ToString(),
                            NameCompetition = excelTable.Cells[i, 4].Value.ToString()
                        }
                    );
                }

                excelWorkbook.Close();
                excelApp.Quit();
                await AddTextToTextBoxAsync("Успешное чтение файла: " + filePath);
            }
            catch (Exception e)
            {
                await AddTextToTextBoxAsync("Произошла ошибка с файлом: " + filePath + "\n" + e);
                if (i == 0)
                {
                    await AddTextToTextBoxAsync("Возможно он используется или не правильно составлен. Попробуйте закрыть файл и перезапустить программу!");
                }
                else
                {
                    await AddTextToTextBoxAsync("Cтрока " + i.ToString() + " содержит неполные данные. Удалите или добавьте данные!");
                }
            }
        }
        public async void CreateListForDiploms(List<PlayersListStruct> list, string filePath, int number)
        {
            int i = 0;
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet excelTable = excelWorkbook.Sheets[number];
                for (i = 2; excelTable.Cells[i, 4].Value is not null; i++)
                {
                    list.Add(
                        new PlayersListStruct()
                        {
                            CodeCompetition = excelTable.Cells[i, 13].Value is not null ? excelTable.Cells[i, 13].Value.ToString() : null,
                            CodeExhibition = excelTable.Cells[i, 14].Value is not null ? excelTable.Cells[i, 14].Value.ToString() : null,
                            CodeContest = excelTable.Cells[i, 15].Value is not null ? excelTable.Cells[i, 15].Value.ToString() : null,
                            OlympicsContest = excelTable.Cells[i, 16].Value is not null ? excelTable.Cells[i, 16].Value.ToString() : null,
                            FioPlayers = excelTable.Cells[i, 4].Value.ToString(),
                            BirthdayPlayers = Convert.ToDateTime(excelTable.Cells[i, 5].Value),
                            SchoolPlayers = excelTable.Cells[i, 12].Value.ToString(),
                            CityPlayers = excelTable.Cells[i, 11].Value is not null ? excelTable.Cells[i, 11].Value.ToString() : null,
                            TeacherPlayers = excelTable.Cells[i, 19].Value is not null ? excelTable.Cells[i, 19].Value.ToString() : null,
                            isMen = excelTable.Cells[i, 22].Value.ToString().Equals("М") ? true : false,
                            NameCommand = excelTable.Cells[i, 8].Value is not null ? excelTable.Cells[i, 8].Value.ToString() : null,
                            eMail = excelTable.Cells[i, 21].Value is not null ? excelTable.Cells[i, 21].Value.ToString() : null,
                        }
                    );
                }
                excelWorkbook.Close();
                excelApp.Quit();
                await AddTextToTextBoxAsync("Успешное чтение файла: " + filePath);
            }
            catch (Exception e)
            {
                await AddTextToTextBoxAsync("Произошла ошибка с файлом: " + filePath + "\n" + e);
                if (i == 0)
                {
                    await AddTextToTextBoxAsync("Возможно он используется или не правильно составлен. Попробуйте закрыть файл и перезапустить программу!");
                }
                else
                {
                    await AddTextToTextBoxAsync("Cтрока " + i.ToString() + " содержит неполные данные. Удалите или добавьте данные!");
                }
            }
        }

        private void ButtonExsit_MouseLeave(object sender, MouseEventArgs e)
        {
            ButtonExsit.Background = new SolidColorBrush(Color.FromRgb(255, 0, 0));
        }

        private void ButtonExsit_MouseEnter(object sender, MouseEventArgs e)
        {
            ButtonExsit.Background = new SolidColorBrush(Color.FromRgb(255, 80, 80));
        }

        private bool isDragging = false;
        private Point lastPosition;

        private void TopPanel_MouseUp(object sender, MouseButtonEventArgs e)
        {
            isDragging = false;
        }
        private void TopPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentPosition = e.GetPosition(this);
                Left = Left - (lastPosition.X - currentPosition.X);
                Top = Top - (lastPosition.Y - currentPosition.Y);
            }
        }

        private void TopPanel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isDragging = true;
            lastPosition = e.GetPosition(this);
        }

        private void ButtonExsit_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Close();
        }

        private void ButtonCollapse_MouseLeave(object sender, MouseEventArgs e)
        {
            ButtonCollapse.Background = new SolidColorBrush(Color.FromRgb(62, 62, 246));
        }

        private void ButtonCollapse_MouseEnter(object sender, MouseEventArgs e)
        {
            ButtonCollapse.Background = new SolidColorBrush(Color.FromRgb(121, 121, 227));
        }
        private void Image_MouseDown_1(object sender, MouseButtonEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }
    }
}