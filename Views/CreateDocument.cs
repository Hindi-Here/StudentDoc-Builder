using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Interop.Access.Dao;
using Spire.Doc;

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
            int[] Three = ThreeCount();
            int[] Good = GoodCount();
            int[] Excellent = ExcellentCount();

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
                AvgList[i] = _getTableInfo.ConvertToNumber(i).Average();
            }

            return AvgList;
        }

        private int[] ThreeCount() // колличество троек у студента
        {
            int[] ThreeCount = new int[_getTableInfo.GetColumnCount()];
            for (int i = 2; i < ThreeCount.Length; i++)
            {
                List<int> numbers = _getTableInfo.ConvertToNumber(i);
                ThreeCount[i] = numbers.Count(n => n == 3);
            }

            return ThreeCount;
        }

        private int[] GoodCount() // колличество четверок у студента
        {
            int[] GoodCount = new int[_getTableInfo.GetColumnCount()];
            for (int i = 2; i < GoodCount.Length; i++)
            {
                List<int> numbers = _getTableInfo.ConvertToNumber(i);
                GoodCount[i] = numbers.Count(n => n == 4);
            }

            return GoodCount;
        }

        private int[] ExcellentCount() // колличество пятерок у студента
        {
            int[] ExcellentCount = new int[_getTableInfo.GetColumnCount()];
            for (int i = 2; i < ExcellentCount.Length; i++)
            {
                List<int> numbers = _getTableInfo.ConvertToNumber(i);
                ExcellentCount[i] = numbers.Count(n => n == 5);
            }

            return ExcellentCount;
        }





        public void CreateReference() // создание документов по шаблону "Справка ПО ВО"
        {
            List<string> students = _getTableInfo.GetColumnTitle();

            AccessInfo InfoFromDisciplineTable = new(_getTableInfo._path, $"D={_getTableInfo._dbTable}");
            List<string> d_disciplines = InfoFromDisciplineTable.GetColumnValues(2);
            List<string> d_credit_units = InfoFromDisciplineTable.GetColumnValues(5);

            int credit_units_sum = GetSumWithoutFTD(d_disciplines, d_credit_units);
            double contact_hours_sum = InfoFromDisciplineTable.GetColumnValues(6).Sum(double.Parse);

            // создание кол-ва документов = кол-ву студентов группы
            for (int i = 2; i < _getTableInfo.GetColumnCount(); i++)
            {
                string wordPath = $@"{_outputPath}\{students[i]}.docx";
                List<string> gradesStudent = _getTableInfo.ConvertToFullString(i);

                SynchronizeSorted(d_disciplines, d_credit_units, gradesStudent);
                SynchronizeUnique(d_disciplines, d_credit_units, gradesStudent);

                Spire.Doc.Document document = new(@"Template\Справка ПО ВО.docx");
                document.Replace("{{ФИО студента}}", students[i], true, true);
                Section section = document.Sections[1];
                Table table = section.AddTable(true);

                const int FooterCount = 2;
                float cellWidth = section.PageSetup.ClientWidth;
                table.ResetCells(d_disciplines.Count + FooterCount, 4);

                SetHeaderTable(table, cellWidth);
                SetBodyTable(table, cellWidth, d_disciplines, d_credit_units, gradesStudent);
                SetFooterTable(table, cellWidth, d_disciplines.Count, credit_units_sum, contact_hours_sum);

                DrawingLine(table, d_disciplines);

                document.SaveToFile(wordPath, FileFormat.Docx);
            }
        }

        private static void SetHeaderTable(Table table, float cellWidth) // шапка таблицы
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

        private static void SetBodyTable(Table table, float cellWidth, List<string> discipline, List<string> credit_units, List<string> contact_hours) // заполнение столбцов в созданных документах
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

        private static void SetFooterTable(Table table, float cellWidth, int size, int credit_units_sum, double contact_hours_sum) // заполнение строк отчета о количестве часов и единиц
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

        private static void DrawingLine(Table table, List<string> discipline) // отрисовка линий после заполнения таблицы
        {
            int count = discipline.Where(s => s.StartsWith("ФТД")).Count(); // раньше эти переменные я в аргументы писал, а теперь так
            int size = discipline.Count;

            for (int cellNumber = 0; cellNumber < 4; cellNumber++)
            {
                table.Rows[0].Cells[cellNumber].CellFormat.Borders.Top.Color = System.Drawing.Color.Black; // ОК отображает
                table.Rows[0].Cells[cellNumber].CellFormat.Borders.Bottom.Color = System.Drawing.Color.Black;

                table.Rows[size - count].Cells[cellNumber].CellFormat.Borders.Bottom.Color = System.Drawing.Color.Black; // использую данные, полученные из метода в аргументах
                table.Rows[size + 1].Cells[cellNumber].CellFormat.Borders.Bottom.Color = System.Drawing.Color.Black;
            }
        }

        private static void SynchronizeSorted(List<string> d_disciplines, List<string> d_credit_units, List<string> gradesStudent) // сортировка и синхронизация на ФТД
        {
            var sortedIndices = d_disciplines
                .Select((discipline, index) => new { Discipline = discipline, Index = index })
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
    }
}