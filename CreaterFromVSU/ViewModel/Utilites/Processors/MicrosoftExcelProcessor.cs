using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using CreaterFromVSU.ViewModel.Utilites.Structs;

namespace CreaterFromVSU.ViewModel.Utilites
{
    class MicrosoftExcelProcessor
    {
        public async void CreateListSummarySheetsCityes(string foldelPathOut, List<PlayersListStruct> playerList, Dictionary<string, string> citiesDic, string path)
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
                        /*await AddTextToTextBoxAsync("Создание файла городу: " + kvpair.Key);
                        */Excel.Application excelApp = new Excel.Application();
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
                        /*await AddTextToTextBoxAsync(("Создан файл: " + kvpair.Value + " -> " + (i - 6).ToString() + " участников") + " | " + path);*/

                    }
                }
                /*await AddTextToTextBoxAsync("Всего: " + k.ToString() + " участников\n Успешное создание всех файлов - " + path);*/
            }
            catch (Exception ex)
            {
               /* MessageBox.Show(ex.Message);*/
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
                /*await AddTextToTextBoxAsync("Произошла ошибка с файлом: " + filePath + " \n" + e);
                if (i == 0)
                {
                    await AddTextToTextBoxAsync("Возможно он используется или не правильно составлен. Попробуйте закрыть файл и перезапустить программу!");
                }
                else
                {
                    await AddTextToTextBoxAsync("Cтрока " + i.ToString() + " содержит неполные данные. Удалите или добавьте данные!");
                }*/
            }
        }
        public async void CreateDictionaryCompetitions(Dictionary<string, ReferenceMaterialDictionary> dictionary, string filePath)
        {
            /*int i = 0;
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
            }*/
        }
        public async void CreateListPersonForDiploms(List<PlayersListStruct> list, string filePath, int number)
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
            /*    excelWorkbook.Close();
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
                }*/
            }
            catch { }
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
                    /*await AddTextToTextBoxAsync(("Создан файл: " + kvpair.Value.TypeCompetition.Replace(@"\", "").Replace("\"", "") + " " + kvpair.Value.NameCompetition.Replace(@"\", "").Replace("\"", "") + " -> " + (i - 6).ToString() + " участников") + " | " + path);
                    */k = k + i - 6;
                }
                /*await AddTextToTextBoxAsync("Файлы " + path + " созданы, кол-во участников: " + k.ToString());
           */ }
            catch (Exception ex)
            {
                /* MessageBox.Show(ex.Message);*/
            }
        }
    }
}
