using System.Collections.Generic;

namespace StudentDoc_Builder.Views
{
    public class Exception(MainWindow mainWindow)
    {
        private readonly MainWindow _mainWindow = mainWindow;

        public bool ExceptionDetection(AccessInfo accessInfo, string format) // обработка ошибок
        {
            _mainWindow.WarningText.Text = string.Empty;
            if (string.IsNullOrWhiteSpace(_mainWindow.AccessFilePath.Text))
            {
                ShowErrorMessage("Не была выбрана база данных [I01]");
                return true;
            }

            if (_mainWindow.TableList.SelectedItem == null)
            {
                ShowErrorMessage("Не выбрана таблица, принадлежащая выбранному Access файлу [I02]");
                return true;
            }

            if (_mainWindow.OutputFormatList.SelectedItem == null)
            {
                ShowErrorMessage("Не выбран формат вывода для создаваемых документов [I03]");
                return true;
            }
            else if (format == "Статистика успеваемости" && accessInfo.GetRowCount() < AccessInfo._statictisRow)
            {
                ShowErrorMessage("Выбранная таблица слишком мала для перезаписи статистических данных. \n" +
                    "Выделите последние четыре строки для записи статистики [T02]");
                return true;
            }
            else if (format == "Статистика успеваемости")
            {
                return false;
            }


            if (string.IsNullOrWhiteSpace(_mainWindow.OutputFilePath.Text))
            {
                ShowErrorMessage("Не была выбрана папка, в которой будут сохранены результаты преобразований [I04]");
                return true;
            }

            if (accessInfo.GetRowCount() < AccessInfo._statictisRow)
            {
                ShowErrorMessage($"Недостаточно строк в таблице. Выделите больше пространства для {accessInfo._dbTable} [T01]");
                return true;
            }
            else if (!accessInfo.CheckOnDisciplineTable())
            {
                ShowErrorMessage("Не существует таблицы с меткой D= для выбранной таблицы [T03]");
                return true;
            }
            else if (accessInfo.D_GetColumnCount() != 7)
            {
                ShowErrorMessage($"Таблица D={accessInfo._dbTable} не соотвествует формату: \n" +
                    "[Номер, Код, Дисциплины, Семестр, Отведенные часы, З.Е., К.Е.] [D01]");
                return true;
            }
            else if (accessInfo.GetRowCount() - AccessInfo._statictisRow != accessInfo.D_GetRowCount())
            {
                ShowErrorMessage("Имеются расхождения в количестве дисцилин между таблицами. \n" +
                    $"Проверьте строки на соответствие: кол-во дисциплин (без учета 4 полей статистики) \n" +
                    $"для {accessInfo._dbTable} = кол-ву дисциплин для таблицы с меткой D={accessInfo._dbTable} [T04]");
                return true;
            }
            else if (!accessInfo.CheckOnName().isMatch)
            {
                ShowErrorMessage("Имеются расхождения в наименовании дисциплин. Рекомендуем \n" +
                    "сверить наименование в выбранной таблице с аналогочной таблицей \n" +
                    $"D={accessInfo._dbTable} в строке {accessInfo.CheckOnName().row} [T05]");
                return true;
            }
            else
            {
                for (int i = 2; i < accessInfo.GetColumnCount(); i++)
                {
                    List<string> ConvertCheck = accessInfo.ConvertToFullString(i);
                    if (ConvertCheck.Count != accessInfo.GetRowCount() - AccessInfo._statictisRow)
                    {
                        ShowErrorMessage($"Некоторые данные таблицы в столбце {i + 1} имеют неккорректные значения [T06]");
                        return true;
                    }
                }
            } 

            return false; 
        } 

        public void ShowErrorMessage(string message) => _mainWindow.WarningText.Text = message;

    }
}
