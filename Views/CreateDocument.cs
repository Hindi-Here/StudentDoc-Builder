using System;
using System.Collections.Generic;
using System.Linq;

using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Access.Dao;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;

namespace StudentDoc_Builder.Views
{
    public class CreateDocument
    {
        private readonly string _outputPath;
        private readonly AccessInfo _getTableInfo;
        private readonly MainWindow _mainWindow;

        public CreateDocument(string outputPath, AccessInfo getTableInfo, MainWindow mainWindow)
        {
            _outputPath = outputPath;
            _getTableInfo = getTableInfo;
            _mainWindow = mainWindow;
        }

        public CreateDocument(AccessInfo getTableInfo, MainWindow mainWindow)
        {
            _getTableInfo = getTableInfo;
            _outputPath = string.Empty;
            _mainWindow = mainWindow;
        }

        public void FillGradeStats() // заполнение статистики оценок в БД
        {
            double[] Avg = AvgGrade();
            int[] Three = GradesCount(3);
            int[] Good = GradesCount(4);
            int[] Excellent = GradesCount(5);

            Database database = _getTableInfo.OpenBase();
            Recordset recordset = database.OpenRecordset(_getTableInfo._dbTable);

            recordset.MoveLast();
            for (int i = 2; i < _getTableInfo.GetColumnCount(); i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    if (!recordset.BOF)
                    {
                        recordset.Edit();
                        switch (j)
                        {
                            case 0:
                                recordset.Fields[i].Value = Excellent[i];
                                break;
                            case 1:
                                recordset.Fields[i].Value = Good[i];
                                break;
                            case 2:
                                recordset.Fields[i].Value = Three[i];
                                break;
                            case 3:
                                recordset.Fields[i].Value = Avg[i];
                                break;
                        }

                        recordset.Update();
                        recordset.MovePrevious();
                    }
                }

                recordset.MoveLast();
            }

            AccessInfo.CloseBase(recordset, database);
            _mainWindow.LogTextBox.Text += $"[Запись статистики] Статистика оценок для группы {_getTableInfo._dbTable} была успешно внесена в таблицу\n";
        }

        private double[] AvgGrade() // средние баллы по студентам
        {
            double[] AvgList = new double[_getTableInfo.GetColumnCount()];

            for (int i = 2; i < AvgList.Length; i++)
            {
                AvgList[i] = Math.Round(_getTableInfo.ConvertToNumber(i).Average(), 2);
            }

            return AvgList;
        }

        private int[] GradesCount(int grade) // подсчет количества оценок
        {
            int[] GradesCount = new int[_getTableInfo.GetColumnCount()];
            for (int i = 2; i < GradesCount.Length; i++)
            {
                List<int> numbers = _getTableInfo.ConvertToNumber(i);
                GradesCount[i] = numbers.Count(n => n == grade);
            }

            return GradesCount;
        }





        public void CreateReference() // создание документов по шаблону "Справка ПО ВО"
        {
            List<string> students = _getTableInfo.GetColumnTitle();

            AccessInfo InfoFromDisciplineTable = new(_getTableInfo._path, $"D={_getTableInfo._dbTable}");
            List<string> d_disciplines = InfoFromDisciplineTable.GetColumnValues(2);
            List<string> d_credit_units = InfoFromDisciplineTable.GetColumnValues(5);

            int credit_units_sum = GetSumWithoutFTD(d_disciplines, d_credit_units);
            double contact_hours_sum = Math.Round(InfoFromDisciplineTable.GetColumnValues(6).Sum(double.Parse), 2);

            // создание кол-ва документов = кол-ву студентов группы
            for (int i = 2; i < _getTableInfo.GetColumnCount(); i++)
            {
                string wordPath = $@"{_outputPath}\{students[i]}.docx";
                List<string> gradesStudent = _getTableInfo.ConvertToFullString(i);

                SynchronizeSortedFtd(d_disciplines, d_credit_units, gradesStudent);
                SynchronizeUnique(d_disciplines, d_credit_units, gradesStudent);

                Spire.Doc.Document document = new(@"Template\Справка ПО ВО.docx");
                document.Replace("{{ФИО студента}}", students[i], true, true);
                Section section = document.Sections[1];
                Table table = section.AddTable(true);

                const int FooterCount = 2;
                float cellWidth = section.PageSetup.ClientWidth;
                table.ResetCells(d_disciplines.Count + FooterCount, 4);

                SetHeaderReferenceTable(table, cellWidth);
                SetBodyReferenceTable(table, cellWidth, d_disciplines, d_credit_units, gradesStudent);
                SetFooterReferenceTable(table, cellWidth, d_disciplines.Count, credit_units_sum, contact_hours_sum);

                DrawingReferenceTableLine(table, d_disciplines);

                _mainWindow.LogTextBox.Text += $"[Создание справки] Справка ПО ВО для студента {students[i]} из группы {_getTableInfo._dbTable} была успешно создана!\n";
                document.SaveToFile(wordPath, FileFormat.Docx);
            }
        }

        private static void SetHeaderReferenceTable(Table table, float cellWidth) // шапка Reference таблицы
        {
            float[] cellWidths = [cellWidth * 0.05f, cellWidth * 0.60f, cellWidth * 0.10f, cellWidth * 0.25f];
            string[] headers = ["№", "Наименование дисциплины (модуля)", "Объем дисциплины (модуля) в зачетных единицах", "Оценка по дисциплине (модулю)"];

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = table.Rows[0].Cells[i];
                cell.SetCellWidth(cellWidths[i], CellWidthType.Point);

                var paragraph = cell.AddParagraph();
                paragraph.AppendText(headers[i]).CharacterFormat.Bold = true;
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                cell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
            }
        }

        private static void SetBodyReferenceTable(Table table, float cellWidth, List<string> discipline, List<string> credit_units, List<string> contact_hours) // заполнение столбцов в созданных документах
        {
            float[] cellWidths = [cellWidth * 0.05f, cellWidth * 0.60f, cellWidth * 0.10f, cellWidth * 0.25f];
            for (int i = 0; i < discipline.Count; i++)
            {
                for (int cellNumber = 0; cellNumber < 4; cellNumber++)
                {
                    var cell = table.Rows[i + 1].Cells[cellNumber];
                    cell.SetCellWidth(cellWidths[cellNumber], CellWidthType.Point);

                    string[] temp = [(i + 1).ToString(), discipline[i], credit_units[i], contact_hours[i]];
                    var paragraph = cell.AddParagraph();
                    paragraph.AppendText(temp[cellNumber]);

                    if (cellNumber != 1)
                        paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    var topBorder = table.Rows[i + 1].Cells[cellNumber].CellFormat.Borders.Top;
                    topBorder.Color = System.Drawing.Color.White;

                    var BottomBorder = table.Rows[i + 1].Cells[cellNumber].CellFormat.Borders.Bottom;
                    BottomBorder.Color = System.Drawing.Color.White;
                }
            }
        }

        private static void SetFooterReferenceTable(Table table, float cellWidth, int size, int credit_units_sum, double contact_hours_sum) // заполнение строк отчета о количестве часов и единиц
        {
            float[] cellWidths = [cellWidth * 0.05f, cellWidth * 0.60f, cellWidth * 0.10f, cellWidth * 0.25f];

            int footerRow1 = size;
            int footerRow2 = size + 1;

            string[] footerRow1Text = ["Объем образовательной программы (без ФТД)", credit_units_sum.ToString() + " з.е,"];
            string[] footerRow2Text = ["Академических часов", contact_hours_sum.ToString() + " ак.час"];

            for (int rowNumber = footerRow1; rowNumber <= footerRow2; rowNumber++)
            {
                for (int cellNumber = 0; cellNumber < 4; cellNumber++)
                {
                    var cell = table.Rows[rowNumber].Cells[cellNumber];
                    cell.SetCellWidth(cellWidths[cellNumber], CellWidthType.Point);

                    if (cellNumber == 1 || cellNumber == 2)
                    {
                        var paragraph = cell.AddParagraph();

                        if (rowNumber == footerRow1)
                        {
                            paragraph.AppendText(footerRow1Text[cellNumber - 1]);
                            if (cellNumber == 2)
                                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                        }
                        else if (rowNumber == footerRow2)
                        {
                            paragraph.AppendText(footerRow2Text[cellNumber - 1]);
                            if (cellNumber == 2)
                                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                        }
                    }

                    var topBorder = table.Rows[rowNumber].Cells[cellNumber].CellFormat.Borders.Top;
                    topBorder.Color = System.Drawing.Color.White;

                    var bottomBorder = table.Rows[rowNumber].Cells[cellNumber].CellFormat.Borders.Bottom;
                    bottomBorder.Color = System.Drawing.Color.White;
                }
            }
        }

        private static void DrawingReferenceTableLine(Table table, List<string> discipline) // отрисовка линий после заполнения таблицы
        {
            int count = discipline.Where(s => s.StartsWith("ФТД")).Count();
            int size = discipline.Count;

            for (int cellNumber = 0; cellNumber < 4; cellNumber++)
            {
                table.Rows[0].Cells[cellNumber].CellFormat.Borders.Top.Color = System.Drawing.Color.Black;
                table.Rows[0].Cells[cellNumber].CellFormat.Borders.Bottom.Color = System.Drawing.Color.Black;

                table.Rows[size - count].Cells[cellNumber].CellFormat.Borders.Bottom.Color = System.Drawing.Color.Black;
                table.Rows[size + 1].Cells[cellNumber].CellFormat.Borders.Bottom.Color = System.Drawing.Color.Black;
            }
        }

        private static void SynchronizeSortedFtd(List<string> d_disciplines, List<string> d_credit_units, List<string> gradesStudent) // сортировка и синхронизация на ФТД
        {
            var sortedIndices = d_disciplines.Select((discipline, index) => new { Discipline = discipline, Index = index })
                .OrderBy(x => x.Discipline.StartsWith("ФТД") ? 1 : 0)
                .Select(x => x.Index)
                .ToList();

            List<string> sortedDisciplines = [];
            List<string> sortedCreditUnits = [];
            List<string> sortedGrades = [];

            foreach (var index in sortedIndices)
            {
                sortedDisciplines.Add(d_disciplines[index]);
                sortedCreditUnits.Add(d_credit_units[index]);
                sortedGrades.Add(gradesStudent[index]);
            }

            d_disciplines.Clear();
            d_disciplines.AddRange(sortedDisciplines);

            d_credit_units.Clear();
            d_credit_units.AddRange(sortedCreditUnits);

            gradesStudent.Clear();
            gradesStudent.AddRange(sortedGrades);
        }

        private static void SynchronizeUnique(List<string> disciplines, List<string> creditUnits, List<string> grades) // уникальные значения и синхронизация данных
        {
            HashSet<string> seen = [];

            List<string> uniqueDisciplines = [];
            List<string> uniqueCreditUnits = [];
            List<string> uniqueGrades = [];

            for (int i = disciplines.Count - 1; i >= 0; i--)
            {
                string discipline = disciplines[i];

                if (seen.Contains(discipline))
                    continue;

                seen.Add(discipline);
                uniqueDisciplines.Add(discipline);
                uniqueCreditUnits.Add(creditUnits[i]);
                uniqueGrades.Add(grades[i]);
            }

            uniqueDisciplines.Reverse();
            uniqueCreditUnits.Reverse();
            uniqueGrades.Reverse();

            disciplines.Clear();
            disciplines.AddRange(uniqueDisciplines);

            creditUnits.Clear();
            creditUnits.AddRange(uniqueCreditUnits);

            grades.Clear();
            grades.AddRange(uniqueGrades);
        }

        static int GetSumWithoutFTD(List<string> d_SourceTable, List<string> SourceContent) // подсчет без учета ФТД
        {
            return Enumerable.Range(0, d_SourceTable.Count)
                             .Where(X => !d_SourceTable[X].StartsWith("ФТД"))
                             .Sum(X => int.Parse(SourceContent[X]));
        }





        public void CreatePortfolio() // создание документов по шаблону "Портфолио"
        {
            List<string> students = _getTableInfo.GetColumnTitle();
            AccessInfo InfoFromDisciplineTable = new(_getTableInfo._path, $"D={_getTableInfo._dbTable}");
            List<string> d_disciplines = InfoFromDisciplineTable.GetColumnValues(2);
            List<string> d_halfYears = InfoFromDisciplineTable.GetColumnValues(3);

            for (int i = 2; i < _getTableInfo.GetColumnCount(); i++)
            {
                string wordPath = $@"{_outputPath}\{students[i]}.docx";
                List<string> gradesStudent = _getTableInfo.ConvertToFullString(i);

                Spire.Doc.Document document = new(@"Template\Портфолио.docx");
                document.Replace("{{ФИО студента}}", students[i], true, true);

                // создание таблицы с оценками по семестрам и манипуляция данными
                BuildGradesTable(document, d_disciplines, d_halfYears, gradesStudent);

                // создание таблицы курсовых работ и манипуляция данными
                SynchronizeCourseWork(d_disciplines, gradesStudent);
                BuildCourseTable(document, d_disciplines, gradesStudent);
                ReturnPortfolioValues(d_disciplines, gradesStudent, i);

                // создание таблицы практик и манипуляция данными
                SynchronizePractice(d_disciplines, gradesStudent);
                BuildPracticeTable(document, d_disciplines, gradesStudent);
                ReturnPortfolioValues(d_disciplines, gradesStudent, i);

                _mainWindow.LogTextBox.Text += $"[Создание портфолио] Портфолио для студента {students[i]} из группы {_getTableInfo._dbTable} было успешно создано!\n";
                document.SaveToFile(wordPath, FileFormat.Docx);
            }
        }

        private static void BuildGradesTable(Spire.Doc.Document document, List<string> disciplines, List<string> halfYears, List<string> grades) // создание таблицы оценок студентов по семестрам
        {
            Section section = document.Sections[1];
            TextRange headerText = section.AddParagraph().AppendText("1. Динамика успеваемости обучающегося");
            headerText.CharacterFormat.Bold = true;
            SetFont(headerText);

            var uniqueHalfYears = halfYears.Distinct().ToList();
            foreach (var halfYear in uniqueHalfYears)
            {
                section.AddParagraph(); // пустой параграф для отступа между таблицами
                TextRange textRange = section.AddParagraph().AppendText($"Успеваемость обучающегося по результатам {halfYear} сессии ");
                textRange.CharacterFormat.Bold = true;
                SetFont(textRange);

                Table table = section.AddTable(true);
                int rowCount = halfYears.Count(h => h == halfYear);
                table.ResetCells(rowCount + 1, 3);

                float cellWidth = section.PageSetup.ClientWidth;
                float[] cellWidths = [cellWidth * 0.05f, cellWidth * 0.70f, cellWidth * 0.25f];
                string[] headers = ["Номер", "Дисциплина", "Оценка"];

                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = table.Rows[0].Cells[i];
                    cell.SetCellWidth(cellWidths[i], CellWidthType.Point);

                    var paragraph = cell.AddParagraph();
                    textRange = paragraph.AppendText(headers[i]);
                    textRange.CharacterFormat.Bold = true;
                    SetFont(textRange);
                    SetTableSetting(paragraph, cell);
                }

                int rowIndex = 1;
                for (int i = 0; i < disciplines.Count; i++)
                {
                    if (halfYears[i] == halfYear)
                    {
                        for (int j = 0; j < headers.Length; j++)
                        {
                            var cell = table.Rows[rowIndex].Cells[j];
                            cell.SetCellWidth(cellWidths[j], CellWidthType.Point);

                            string[] temp = [rowIndex + ".", disciplines[i], grades[i]];
                            var paragraph = cell.AddParagraph();
                            textRange = paragraph.AppendText(temp[j]);

                            SetFont(textRange);
                            SetTableSetting(paragraph, cell);
                        }
                        rowIndex++;
                    }
                }
            }
        }

        private static void BuildCourseTable(Spire.Doc.Document document, List<string> disciplines, List<string> grades) // создание таблицы с данными курсовых
        {
            Section section = document.Sections[1];
            section.AddParagraph();

            TextRange headerText = section.AddParagraph().AppendText("2. Сведения о курсовых работах и курсовых проектах");
            headerText.CharacterFormat.Bold = true;
            SetFont(headerText);

            Table table = section.AddTable(true);
            table.ResetCells(disciplines.Count + 1, 4);

            float cellWidth = section.PageSetup.ClientWidth;
            float[] cellWidths = [cellWidth * 0.05f, cellWidth * 0.35f, cellWidth * 0.35f, cellWidth * 0.25f];
            string[] headers = ["№ п/п", "Дисциплина", "Тема работы", "Оценка"];

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = table.Rows[0].Cells[i];
                cell.SetCellWidth(cellWidths[i], CellWidthType.Point);

                var paragraph = cell.AddParagraph();
                TextRange textRange = paragraph.AppendText(headers[i]);

                textRange.CharacterFormat.Bold = true;
                SetFont(textRange);
                SetTableSetting(paragraph, cell);
            }

            List<string> courseNames = GetNameCourseWork(disciplines);
            for (int i = 0; i < disciplines.Count; i++)
            {
                for (int j = 0; j < headers.Length; j++)
                {
                    var cell = table.Rows[i + 1].Cells[j];
                    cell.SetCellWidth(cellWidths[j], CellWidthType.Point);

                    string[] temp = [(i + 1).ToString() + ".", courseNames[i], "", grades[i]];
                    var paragraph = cell.AddParagraph();
                    TextRange textRange = paragraph.AppendText(temp[j]);

                    SetFont(textRange);
                    SetTableSetting(paragraph, cell);
                }
            }
        }

        private static void BuildPracticeTable(Spire.Doc.Document document, List<string> disciplines, List<string> grades) // создание таблицы с данными о практиках
        {
            Section section = document.Sections[1];
            section.AddParagraph();

            TextRange headerText = section.AddParagraph().AppendText("3. Сведения о практиках");
            headerText.CharacterFormat.Bold = true;
            SetFont(headerText);

            Table table = section.AddTable(true);
            table.ResetCells(disciplines.Count + 1, 5);

            float cellWidth = section.PageSetup.ClientWidth;
            float[] cellWidths = [cellWidth * 0.05f, cellWidth * 0.30f, cellWidth * 0.30f, cellWidth * 0.10f, cellWidth * 0.25f];
            string[] headers = ["№ п/п", "Вид практики", "Место прохождения", "Сроки прохождения", "Оценка"];

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = table.Rows[0].Cells[i];
                cell.SetCellWidth(cellWidths[i], CellWidthType.Point);

                var paragraph = cell.AddParagraph();
                TextRange textRange = paragraph.AppendText(headers[i]);

                textRange.CharacterFormat.Bold = true;
                SetFont(textRange);
                SetTableSetting(paragraph, cell);
            }

            for (int i = 0; i < disciplines.Count; i++)
            {
                for (int j = 0; j < headers.Length; j++)
                {
                    var cell = table.Rows[i + 1].Cells[j];
                    cell.SetCellWidth(cellWidths[j], CellWidthType.Point);

                    string[] temp = [(i + 1).ToString() + ".", disciplines[i], "", "", grades[i]];
                    var paragraph = cell.AddParagraph();
                    TextRange textRange = paragraph.AppendText(temp[j]);

                    SetFont(textRange);
                    SetTableSetting(paragraph, cell);
                }
            }
        }

        private static void SetFont(TextRange textRange) // назначение фонта для текста
        {
            textRange.CharacterFormat.FontName = "Times New Roman";
            textRange.CharacterFormat.FontSize = 12;
        }

        private static void SetTableSetting(Paragraph paragraph, TableCell cell) // центрирование по центру и интервал
        {
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            cell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

            paragraph.Format.LineSpacingRule = 0;
            paragraph.Format.AfterSpacing = 0;
        }

        private void ReturnPortfolioValues(List<string> disciplines, List<string> grades, int index) // вернуть начальные значения
        {
            grades.Clear();
            grades.AddRange(_getTableInfo.ConvertToFullString(index));

            AccessInfo InfoFromDisciplineTable = new(_getTableInfo._path, $"D={_getTableInfo._dbTable}");
            disciplines.Clear();
            disciplines.AddRange(InfoFromDisciplineTable.GetColumnValues(2));
        }

        private static void SynchronizeCourseWork(List<string> disciplines, List<string> grades) // синхронизация оценок и курсовых работ
        {
            var coursework = disciplines.Select((disciplines, index) => new { Discipline = disciplines, Grade = grades.ElementAtOrDefault(index) })
                                        .Where(d => d.Discipline.StartsWith("Курсовая"))
                                        .ToList();

            disciplines.Clear();
            grades.Clear();

            foreach (var item in coursework)
            {
                disciplines.Add(item.Discipline);
                if (item.Grade != null)
                    grades.Add(item.Grade);
            }
        }

        private static List<string> GetNameCourseWork(List<string> disciplines) // получить список дисциплин
        {
            string pattern = "\"([^\"]+)\"";
            Regex regex = new(pattern);
            return disciplines.Where(d => d.StartsWith("Курсовая"))
                              .Select(d => regex.Match(d).Groups[1].Value)
                              .ToList();
        }

        private static void SynchronizePractice(List<string> disciplines, List<string> grades) // синхронизация оценок и курсовых работ
        {
            var coursework = disciplines.Select((disciplines, index) => new { Discipline = disciplines, Grade = grades.ElementAtOrDefault(index) })
                                        .Where(d => d.Discipline.StartsWith("Учебная") || d.Discipline.StartsWith("Производственная"))
                                        .ToList();

            disciplines.Clear();
            grades.Clear();

            foreach (var item in coursework)
            {
                disciplines.Add(item.Discipline);
                if (item.Grade != null)
                    grades.Add(item.Grade);
            }
        }




        public void CreatePersonalCard() // создание документов по шаблону "Личная справка"
        {
            List<string> students = _getTableInfo.GetColumnTitle();

            AccessInfo InfoFromDisciplineTable = new(_getTableInfo._path, $"D={_getTableInfo._dbTable}");
            List<string> d_disciplines = InfoFromDisciplineTable.GetColumnValues(2);
            List<string> d_halfYears = InfoFromDisciplineTable.GetColumnValues(3);
            List<string> d_hours = InfoFromDisciplineTable.GetColumnValues(4);
            List<string> d_creditUnits = InfoFromDisciplineTable.GetColumnValues(5);

            for (int i = 2; i < _getTableInfo.GetColumnCount(); i++)
            {
                string wordPath = $@"{_outputPath}\{students[i]}.docx";

                Spire.Doc.Document document = new(@"Template\личная-карточка.docx");
                document.Replace("{{ФИО студента}}", students[i], true, true);

                BuildCardTable(document, d_halfYears, d_disciplines, d_hours, d_creditUnits);

                _mainWindow.LogTextBox.Text += $"[Создание карточки] Личная карточка для студента {students[i]} из группы {_getTableInfo._dbTable} было успешно создано!\n";
                document.SaveToFile(wordPath, FileFormat.Docx);
            }
        }

        private void BuildCardTable(Spire.Doc.Document document, List<string> halfYears, List<string> disciplines, List<string> hours, List<string> creditUnits) // создание таблиц личных-карточек
        {
            Section section = document.Sections[1];
            float cellWidth = section.PageSetup.ClientWidth;
            float[] cellWidths = [cellWidth * 0.05f,cellWidth * 0.45f,cellWidth * 0.06f,cellWidth * 0.06f, cellWidth * 0.06f,
                                  cellWidth * 0.06f,cellWidth * 0.06f,cellWidth * 0.06f,cellWidth * 0.14f];

            for (int year = 1; year <= 5; year++)
            {
                SynchronizeSemester(year, disciplines, halfYears, hours, creditUnits);

                Table table = section.AddTable(true);
                table.ResetCells(disciplines.Count + 3, 9);

                for (int row = 0; row < disciplines.Count + 3; row++)
                    for (int column = 0; column < cellWidths.Length; column++)
                        table.Rows[row].Cells[column].SetCellWidth(cellWidths[column], CellWidthType.Point);

                SetHeaderCardTable(document, table, year);
                SetBodyCardTable(table, disciplines, halfYears, hours, creditUnits);

                ReturnCardValues(disciplines, halfYears, hours, creditUnits);

                section.AddParagraph();
                section.AddParagraph();
                section.AddParagraph().AppendText("Директор института ____________________________________");
                section.AddParagraph();
                section.AddParagraph();
            }
        }

        private static void SetHeaderCardTable(Spire.Doc.Document document, Table table, int year) // шапка Card таблиц
        {
            MergeCells(table);

            TableCell headerCell = table.Rows[0].Cells[0];
            Paragraph paragraph = headerCell.AddParagraph();

            CharacterFormat charFormat = new(document)
            {
                Bold = true,
                FontName = "Times New Roman",
                FontSize = 12
            };

            paragraph.AppendText($"{year} курс 2023-2024 учебного года").ApplyCharacterFormat(charFormat);
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            SetCellText(table, charFormat);
        }

        private static void SetBodyCardTable(Table table, List<string> disciplines, List<string> halfYears, List<string> hours, List<string> creditUnits) // заполнение тела таблицы
        {
            for (int i = 0; i < disciplines.Count; i++)
            {
                for (int j = 0; j < 9; j++)
                {
                    var cell = table.Rows[i + 3].Cells[j];

                    string[] temp = ["", disciplines[i], hours[i], creditUnits[i], "", "", "", "", ""];
                    var paragraph = cell.AddParagraph();
                    paragraph.AppendText(temp[j]);

                    if (j != 1)
                    {
                        paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                        cell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    }
                }
            }

            MergeSemesterCells(table, halfYears);
        }

        private static void CellText(TableCell cell, string text, CharacterFormat charFormat) // форматирование текста
        {
            Paragraph paragraph = cell.AddParagraph();
            paragraph.AppendText(text).ApplyCharacterFormat(charFormat);
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            cell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
        }

        private static void SetCellText(Table table, CharacterFormat charFormat) // запись данных в таблицу
        {
            CellText(table.Rows[1].Cells[0], "Наименование дисциплины\n(без сокращения)", charFormat);
            CellText(table.Rows[1].Cells[2], "Кол-во", charFormat);
            CellText(table.Rows[1].Cells[4], "Экзамен. оценка", charFormat);
            CellText(table.Rows[1].Cells[6], "Отметка о зачете", charFormat);
            CellText(table.Rows[1].Cells[8], "Дата и № ведомости", charFormat);

            CellText(table.Rows[2].Cells[2], "час", charFormat);
            CellText(table.Rows[2].Cells[3], "зач. ед.", charFormat);
            CellText(table.Rows[2].Cells[4], "оц.", charFormat);
            CellText(table.Rows[2].Cells[5], "балл", charFormat);
            CellText(table.Rows[2].Cells[6], "оц.", charFormat);
            CellText(table.Rows[2].Cells[7], "балл", charFormat);
        }

        private static void MergeCells(Table table) // объединение ячеек
        {
            table.ApplyHorizontalMerge(0, 0, 8);

            table.ApplyHorizontalMerge(1, 0, 1);
            table.ApplyHorizontalMerge(2, 0, 1);
            table.ApplyHorizontalMerge(1, 2, 3);
            table.ApplyHorizontalMerge(1, 4, 5);
            table.ApplyHorizontalMerge(1, 6, 7);

            table.ApplyVerticalMerge(0, 1, 2);
            table.ApplyVerticalMerge(8, 1, 2);
        }

        private static void MergeSemesterCells(Table table, List<string> halfYears) // объединение ячеек семестров
        {
            if (halfYears.Count == 0) return;

            int startRow = 3;
            string currentSemester = halfYears[0];
            int count = 1;

            for (int i = 1; i < halfYears.Count; i++)
            {
                if (halfYears[i] == currentSemester)
                {
                    count++;
                }
                else
                {
                    table.ApplyVerticalMerge(0, startRow, startRow + count - 1);

                    currentSemester = halfYears[i];
                    startRow = i + 3;
                    count = 1;
                }
            }

            table.ApplyVerticalMerge(0, startRow, startRow + count - 1);
        }

        private static void SynchronizeSemester(int year, List<string> disciplines, List<string> halfYears, List<string> hours, List<string> creditUnits) // синхронизация данных по семестрам
        {
            int startSemester = (year - 1) * 2 + 1;
            int endSemester = startSemester + 1;

            var coursework = disciplines.Select((discipline, index) => new
            {
                Discipline = discipline,
                Semester = halfYears.ElementAtOrDefault(index),
                Hour = hours.ElementAtOrDefault(index),
                CreditUnit = creditUnits.ElementAtOrDefault(index)
            })
            .Where(d => int.TryParse(d.Semester, out int semester) && semester >= startSemester && semester <= endSemester)
            .ToList();

            disciplines.Clear();
            hours.Clear();
            creditUnits.Clear();
            halfYears.Clear();

            foreach (var item in coursework)
            {
                disciplines.Add(item.Discipline);
                if (item.Semester != null)
                    halfYears.Add(item.Semester);

                if (item.Hour != null)
                    hours.Add(item.Hour);

                if (item.CreditUnit != null)
                    creditUnits.Add(item.CreditUnit);
            }
        }

        private void ReturnCardValues(List<string> disciplines, List<string> halfYears, List<string> hours, List<string> creditUnits) // вернуть начальные значения
        {
            AccessInfo InfoFromDisciplineTable = new(_getTableInfo._path, $"D={_getTableInfo._dbTable}");

            disciplines.Clear();
            hours.Clear();
            creditUnits.Clear();
            halfYears.Clear();

            disciplines.AddRange(InfoFromDisciplineTable.GetColumnValues(2));
            halfYears.AddRange(InfoFromDisciplineTable.GetColumnValues(3));
            hours.AddRange(InfoFromDisciplineTable.GetColumnValues(4));
            creditUnits.AddRange(InfoFromDisciplineTable.GetColumnValues(5));
        }

    }
}