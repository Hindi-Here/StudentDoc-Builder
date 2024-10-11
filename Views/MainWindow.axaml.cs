using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Media;

using Microsoft.Office.Interop.Access.Dao;
using Avalonia.Platform;

namespace StudentDoc_Builder.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void LogPanel_Click(object sender, RoutedEventArgs e) // открытие и зыкрытие панели логов
        {
            var logPanelButton = this.FindControl<Button>("LogPanel");
            if (logPanelButton != null)
            {
                if (Width == 700)
                {
                    Width += 450;
                    logPanelButton.Content = "<";
                }
                else
                {
                    Width = 700;
                    logPanelButton.Content = ">";
                }
            }
        }

        private async void ChooseAccessFile_Click(object sender, RoutedEventArgs e) // выбор Access файла
        {
            var topLevel = GetTopLevel(this);

            if (topLevel == null)
                return;

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Выберите файл Access",
                AllowMultiple = false,
                FileTypeFilter = [new("Access Files") { Patterns = ["*.accdb", "*.mdb"] }]
            });

            if (files.Count > 0)
            {
                AccessFilePath.Text = files[0].Path.LocalPath;
                LoadTables();
            }

        }

        private async void ChooseOutputFormat_Click(object sender, RoutedEventArgs e) // выбор папки вывода
        {
            var topLevel = GetTopLevel(this);

            if (topLevel == null)
                return;

            var folder = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Open Text File"
            });

            if (folder.Count > 0)
                OutputFilePath.Text = folder[0].Path.LocalPath;
        }

        private void LoadTables() // подключаемся к Access и ищем имена таблиц
        {
            TableList?.Items.Clear();

            var dbEngine = new DBEngine();
            Database database = dbEngine.OpenDatabase(AccessFilePath.Text);

            foreach (TableDef tableDef in database.TableDefs)
            {
                if (!tableDef.Name.StartsWith("MSys") && !tableDef.Name.StartsWith("D=") && !tableDef.Name.StartsWith('~'))
                {
                    TableList?.Items.Add(new ListBoxItem { Content = tableDef.Name });
                }
            }
            database.Close();
        }

        struct TableInfo {
            public string DbTable { get; set; }
            public string DbFormat { get; set; }
        }
        TableInfo tableInfo;

        private void TableList_SelectionChanged(object? sender, SelectionChangedEventArgs e) // определение выбранной таблицы
        {
            if (sender is ListBox { SelectedItem: ListBoxItem selectedItem })
            {
                if (selectedItem.Content is string content)
                {
                    tableInfo.DbTable = content;
                    selectedItem.FontWeight = Avalonia.Media.FontWeight.Bold;
                }
            }
        }

        private void OutputFormatList_SelectionChanged(object? sender, SelectionChangedEventArgs e) // определение выбранного формата
        {
            if (sender is ListBox { SelectedItem: ListBoxItem selectedItem })
            {
                if (selectedItem.Content is string content)
                {
                    tableInfo.DbFormat = content;
                    selectedItem.FontWeight = Avalonia.Media.FontWeight.Bold;
                }
            }
        }

        private void CreateDocument_Click(object sender, RoutedEventArgs e) // создание документа по шаблону
        {
            WarningText.Foreground = Brushes.Red;
            AccessInfo accessInfo = new(AccessFilePath.Text, tableInfo.DbTable);

            Exception exception = new(this);
            if (exception.ExceptionDetection(accessInfo, tableInfo.DbFormat))
                return;

            WarningText.Foreground = Brushes.Green;
            switch (tableInfo.DbFormat)
            {
                case "Статистика успеваемости":
                    CreateDocument WriteStats = new(accessInfo, this);
                    WriteStats.FillGradeStats();
                    WarningText.Text = $"Данные были успешно занесены в таблицу {tableInfo.DbTable}!";
                    break;
                case "Справка ПО ВО":
                    CreateDocument Reference = new(OutputFilePath.Text, accessInfo, this);
                    Reference.CreateReference();
                    WarningText.Text = $"Документы для группы {tableInfo.DbTable} были созданы!";
                    break;
                case "Личная карточка":
                    CreateDocument Card = new(OutputFilePath.Text, accessInfo, this);
                    Card.CreatePersonalCard();
                    WarningText.Text = $"Документы для группы {tableInfo.DbTable} были созданы!";
                    break;
                case "Портфолио":
                    CreateDocument Portfolio = new(OutputFilePath.Text, accessInfo, this);
                    Portfolio.CreatePortfolio();
                    WarningText.Text = $"Документы для группы {tableInfo.DbTable} были созданы!";
                    break;
            }
        }

    }
}