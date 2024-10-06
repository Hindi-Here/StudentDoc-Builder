using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;

namespace StudentDoc_Builder.Views
{
    public class AccessInfo(string path, string dbTable)
    {
        public readonly string _path = path;
        public readonly string _dbTable = dbTable;
        public const int _statictisRow = 4;

        public Database OpenBase()
        {
            var dbEngine = new DBEngine();
            return dbEngine.OpenDatabase(_path);
        }

        public static void CloseBase(Recordset recordset, Database database)
        {
            recordset.Close();
            database.Close();
        }

        public int GetRowCount() // получить количество записей
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            int row = tabledef.RecordCount;

            database.Close();
            return row;
        }

        public int D_GetRowCount() // получить количество записей для d метки (используется в Exception для доп. проверки)
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[$"D={_dbTable}"];

            int row = tabledef.RecordCount;

            database.Close();
            return row;
        }

        public int GetColumnCount() // получить количество стобцов
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            int column = tabledef.Fields.Count;

            database.Close();
            return column;
        }

        public int D_GetColumnCount() // получить количество столбцов для d метки (используется в Exception для доп. проверки)
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[$"D={_dbTable}"];

            int column = tabledef.Fields.Count;

            database.Close();
            return column;
        }

        public List<string> GetColumnTitle() // получить заголовки столбцов
        { 
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            List<string> columnTitle = [];
            for (int i = 0; i < tabledef.Fields.Count; i++)
                columnTitle.Add(tabledef.Fields[i].Name);

            database.Close();
            return columnTitle;
        }

        public List<string> GetColumnValues(int columnIndex) // получить данные конкретного столбца
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            int rowCount = tabledef.RecordCount;
            List<string> columnValues = [];

            string query = $"SELECT * FROM [{_dbTable}]";
            Recordset recordset = database.OpenRecordset(query);

            for (int i = 0; i < rowCount; i++)
            {
                var fieldValue = recordset.Fields[columnIndex].Value.ToString();
                columnValues.Add(fieldValue != null ? fieldValue.ToString() : string.Empty);
                recordset.MoveNext();
            }

            CloseBase(recordset, database);
            return columnValues;
        }

        public List<string> GetRowValues(int rowIndex) // получить данные конкретной записи
        { 
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            int columnCount = tabledef.Fields.Count;
            List<string> rowValues = [];

            string query = $"SELECT * FROM [{_dbTable}]";
            Recordset recordset = database.OpenRecordset(query);
            recordset.Move(rowIndex);

            for (int i = 0; i < columnCount; i++)
            {
                var fieldValue = recordset.Fields[i].Value.ToString();
                rowValues.Add(fieldValue != null ? fieldValue.ToString() : string.Empty);
            }

            CloseBase(recordset, database);
            return rowValues;
        }

        private readonly Dictionary<string, int> ConvertNumber = new() // словарь для подсчета статистики
        {
            ["3"] = 3,
            ["4"] = 4,
            ["5"] = 5,
        };

        public List<int> ConvertToNumber(int columnIndex) // получение числовых оценок для указанного столбца
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            string query = $"SELECT * FROM [{_dbTable}]";
            Recordset recordset = database.OpenRecordset(query);

            List<int> ints = [];

            recordset.MoveFirst();
            for (int i = 0; i < tabledef.RecordCount - (_statictisRow - 1); i++)
            {
                var fieldValue = recordset.Fields[columnIndex].Value.ToString();

                if (fieldValue != null && ConvertNumber.TryGetValue(fieldValue, out int numericValue))
                {
                    ints.Add(numericValue);
                }

                recordset.MoveNext(); 
            }

            CloseBase(recordset, database);
            return ints;
        }

        private readonly Dictionary<string, string> ConvertFullString = new() // словарь полных названий
        {
            ["3"] = "Удовлетворительно",
            ["4"] = "Хорошо",
            ["5"] = "Отлично",
            ["зач"] = "Зачтено",
            ["незач"] = "Не зачтено",
            [string.Empty] = "Не оценено"
        };

        public List<string> ConvertToFullString(int columnIndex) // получение полных названий значений столбцов
        {
            Database database = OpenBase();
            TableDef tabledef = database.TableDefs[_dbTable];

            string query = $"SELECT * FROM [{_dbTable}]";
            Recordset recordset = database.OpenRecordset(query);

            List<string> strings = [];

            recordset.MoveFirst();
            for (int i = 0; i < tabledef.RecordCount - (_statictisRow - 1); i++)
            {
                var fieldValue = recordset.Fields[columnIndex].Value.ToString();

                if (fieldValue != null && ConvertFullString.TryGetValue(fieldValue, out var fullString))
                    strings.Add(fullString);

                recordset.MoveNext();
            }

            CloseBase(recordset, database);
            return strings;
        }

        public bool CheckOnDisciplineTable()  // проверка на существование аналога с меткой D=
        {
            Database database = OpenBase();

            foreach (TableDef tableDef in database.TableDefs)
            {
                if (tableDef.Name.Equals($"D={_dbTable}", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        public (bool isMatch, int row) CheckOnName() // проверка на совпадение дисциплин из двух таблиц
        {
            List<string> disciplines = GetColumnValues(1);
            AccessInfo InfoFromDisciplineTable = new(_path, $"D={_dbTable}");
            List<string> d_disciplines = InfoFromDisciplineTable.GetColumnValues(2);

            int validDisciplineCount = disciplines.Count - _statictisRow;

            HashSet<string> disciplineSet = new(d_disciplines);
            for (int i = 0; i < validDisciplineCount; i++)
                if (!disciplineSet.Contains(disciplines[i]))
                    return (false, i + 1);

            return (true, 0);
        }

        public (bool isMatch, int column) CheckOnValues() // проверка на числовые данные для D= таблиц
        {
            AccessInfo InfoFromDisciplineTable = new(_path, $"D={_dbTable}");
            for (int i = 3; i <= 6; i++) // задаем диапазон числовых столбцов
            {
                List<string> columnValues = InfoFromDisciplineTable.GetColumnValues(i);

                foreach (var value in columnValues)
                    if (string.IsNullOrWhiteSpace(value) || !double.TryParse(value, out _))
                        return (false, i);
            }
            return (true, 0);
        }
    }
}