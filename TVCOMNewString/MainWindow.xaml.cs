using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Path = System.IO.Path;
using Excel = Microsoft.Office.Interop.Excel;

namespace TVCOMNewString
{
    
    public partial class MainWindow : Window
    {
        public int Number = 1;
        public ObservableCollection<DateControl> Controls { get; set; } = new ObservableCollection<DateControl>();
        private bool isCollapsed = false;
        private ObservableCollection<Advertisement> _advertisements;
        public ObservableCollection<Advertisement> Advertisements
        {
            get => _advertisements;
            set
            {
                if (_advertisements != null)
                {
                    _advertisements.CollectionChanged -= Advertisements_CollectionChanged;
                }

                _advertisements = value;

                if (_advertisements != null)
                {
                    _advertisements.CollectionChanged += Advertisements_CollectionChanged;

                    // Подписываемся на изменения существующих элементов
                    foreach (Advertisement item in _advertisements)
                    {
                        item.PropertyChanged += Advertisement_PropertyChanged;
                    }
                }

                OnPropertyChanged(nameof(Advertisements));
            }
        }

        private bool _isUpdating = false;
        public string connStr = @"Data Source=C:\TVCOMString\TVCOMNEWSTRING.db;Version=3;";
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            _advertisements = new ObservableCollection<Advertisement>();
            dataGrid1.ItemsSource = Advertisements;
            _advertisements.CollectionChanged += Advertisements_CollectionChanged;
            NumberOfControls.Text = Number.ToString();
            LoadTable();
        }
        private void Advertisements_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (Advertisement item in e.NewItems)
                {
                    item.PropertyChanged += Advertisement_PropertyChanged;
                }
            }
        }

        private void Advertisement_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!_isUpdating)
            {
                var advertisement = sender as Advertisement;
                AutoSaveRowChanges(advertisement);
            }
        }

        private void AutoSaveRowChanges(Advertisement advertisement)
        {
            try
            {
                _isUpdating = true;
                string updateQuery = @"UPDATE [Объявления] 
                         SET [Текст_объявления] = @text, 
                             [заказчик] = @customer, 
                             [дата_подачи] = @dateOpen, 
                             [дата_закрытия] = @dateClose, 
                             [цвет] = @color, 
                             [телефон] = @phone 
                         WHERE [Код_объявления] = @code";

                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(updateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@text", advertisement.ТекстОбъявления ?? "");
                        cmd.Parameters.AddWithValue("@customer", advertisement.Заказчик ?? "");
                        cmd.Parameters.AddWithValue("@dateOpen", advertisement.ДатаПодачи.Date);
                        cmd.Parameters.AddWithValue("@dateClose", advertisement.ДатаЗакрытия.Date);
                        cmd.Parameters.AddWithValue("@color", string.IsNullOrEmpty(advertisement.Цвет) ? "100,143,143,143" : advertisement.Цвет);
                        cmd.Parameters.AddWithValue("@phone", string.IsNullOrEmpty(advertisement.Телефон) ? (object)DBNull.Value : advertisement.Телефон);
                        cmd.Parameters.AddWithValue("@code", advertisement.КодОбъявления);

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected == 0)
                        {
                            MessageBox.Show("Не удалось обновить объявление. Возможно, оно было удалено другим пользователем.", "Предупреждение");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка автосохранения: {ex.Message}", "Ошибка");
            }
            finally
            {
                _isUpdating = false;
            }
        }
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(NumberOfControls.Text, out int numberOfControls))
            {
                while (Controls.Count < numberOfControls)
                {
                    Controls.Add(new DateControl());
                }
                while (Controls.Count > numberOfControls)
                {
                    Controls.RemoveAt(Controls.Count - 1);
                }
            }
        }
        private void LeftNumber_Click(object sender, RoutedEventArgs e)
        {
            if (Number == 1)
            {
                LeftNumber.IsEnabled = false;
                return;
            }
            else
            {
                RightNumber.IsEnabled = true;
            }
            Number -= 1;
            NumberOfControls.Text = Number.ToString();
        }

        private void RightNumber_Click(object sender, RoutedEventArgs e)
        {
            if (Number == 10)
            {
                RightNumber.IsEnabled = false;
                return;
            }
            else
            {
                LeftNumber.IsEnabled = true;
            }
            Number += 1;
            NumberOfControls.Text = Number.ToString();
        }

        private void hideBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!isCollapsed) // Сворачиваем
            {
                // Анимация грида
                DoubleAnimation MenuAnimation = new DoubleAnimation();
                MenuAnimation.From = 275; // или ваша текущая высота
                MenuAnimation.To = 0;
                MenuAnimation.Duration = TimeSpan.FromSeconds(0.5);
                MenuAnimation.EasingFunction = new CubicEase();

                // Анимация кнопки
                ThicknessAnimation ButtonAnimation = new ThicknessAnimation();
                ButtonAnimation.From = hideBtn.Margin;
                ButtonAnimation.To = new Thickness(10, 10, 10, 0); // Новая позиция кнопки
                ButtonAnimation.Duration = TimeSpan.FromSeconds(0.5);
                ButtonAnimation.EasingFunction = new CubicEase();
                
                // Анимация Поиска
                ThicknessAnimation SearchAnimation = new ThicknessAnimation();
                SearchAnimation.From = SearchPanel.Margin;
                SearchAnimation.To = new Thickness(10, 40, 20, 0);
                SearchAnimation.Duration = TimeSpan.FromSeconds(0.5);
                SearchAnimation.EasingFunction = new CubicEase();
                
                // Анимация Кнопок
                ThicknessAnimation ButtonsAnimation = new ThicknessAnimation();
                ButtonsAnimation.From = ButtonPanel.Margin;
                ButtonsAnimation.To = new Thickness(10, 155, 20, 0);
                ButtonsAnimation.Duration = TimeSpan.FromSeconds(0.5);
                ButtonsAnimation.EasingFunction = new CubicEase();

                // Анимация датагрида
                DoubleAnimation DataGridAnimation = new DoubleAnimation();
                DataGridAnimation.From = 200; // или ваша текущая высота
                DataGridAnimation.To = 450;
                DataGridAnimation.Duration = TimeSpan.FromSeconds(0.5);
                DataGridAnimation.EasingFunction = new CubicEase();

                // Анимация кнопки удаления
                ThicknessAnimation ButtonDelAnimation = new ThicknessAnimation();
                ButtonDelAnimation.From = deleteBtn.Margin;
                ButtonDelAnimation.To = new Thickness(0, 0, 10, 465);
                ButtonDelAnimation.Duration = TimeSpan.FromSeconds(0.5);
                ButtonDelAnimation.EasingFunction = new CubicEase();

                AddAd.BeginAnimation(HeightProperty, MenuAnimation);
                hideBtn.BeginAnimation(MarginProperty, ButtonAnimation);
                SearchPanel.BeginAnimation(MarginProperty, SearchAnimation);
                ButtonPanel.BeginAnimation(MarginProperty, ButtonsAnimation);
                dataGrid1.BeginAnimation(HeightProperty, DataGridAnimation);
                deleteBtn.BeginAnimation(MarginProperty, ButtonDelAnimation);

                hideBtn.Content = "Показать";
                isCollapsed = true;
            }
            else // Разворачиваем
            {
                // Анимация грида
                DoubleAnimation MenuAnimation = new DoubleAnimation();
                MenuAnimation.From = 0;
                MenuAnimation.To = 275;
                MenuAnimation.Duration = TimeSpan.FromSeconds(0.5);
                MenuAnimation.EasingFunction = new CubicEase();

                // Анимация кнопки
                ThicknessAnimation ButtonAnimation = new ThicknessAnimation();
                ButtonAnimation.From = hideBtn.Margin;
                ButtonAnimation.To = new Thickness(10, 255, 10, 0); // Исходная позиция кнопки
                ButtonAnimation.Duration = TimeSpan.FromSeconds(0.5);
                ButtonAnimation.EasingFunction = new CubicEase();

                // Анимация Поиска
                ThicknessAnimation SearchAnimation = new ThicknessAnimation();
                SearchAnimation.From = SearchPanel.Margin;
                SearchAnimation.To = new Thickness(10, 285, 20, 0);
                SearchAnimation.Duration = TimeSpan.FromSeconds(0.5);
                SearchAnimation.EasingFunction = new CubicEase();

                // Анимация Кнопок
                ThicknessAnimation ButtonsAnimation = new ThicknessAnimation();
                ButtonsAnimation.From = ButtonPanel.Margin;
                ButtonsAnimation.To = new Thickness(10, 400, 20, 0);
                ButtonsAnimation.Duration = TimeSpan.FromSeconds(0.5);
                ButtonsAnimation.EasingFunction = new CubicEase();

                // Анимация датагрида
                DoubleAnimation DataGridAnimation = new DoubleAnimation();
                DataGridAnimation.From = 450; // или ваша текущая высота
                DataGridAnimation.To = 200;
                DataGridAnimation.Duration = TimeSpan.FromSeconds(0.5);
                DataGridAnimation.EasingFunction = new CubicEase();

                // Анимация кнопки удаления
                ThicknessAnimation ButtonDelAnimation = new ThicknessAnimation();
                ButtonDelAnimation.From = deleteBtn.Margin;
                ButtonDelAnimation.To = new Thickness(0, 0, 10, 222);
                ButtonDelAnimation.Duration = TimeSpan.FromSeconds(0.5);
                ButtonDelAnimation.EasingFunction = new CubicEase();

                AddAd.BeginAnimation(HeightProperty, MenuAnimation);
                hideBtn.BeginAnimation(MarginProperty, ButtonAnimation);
                SearchPanel.BeginAnimation(MarginProperty, SearchAnimation);
                ButtonPanel.BeginAnimation(MarginProperty, ButtonsAnimation);
                dataGrid1.BeginAnimation(HeightProperty, DataGridAnimation);
                deleteBtn.BeginAnimation(MarginProperty, ButtonDelAnimation);

                hideBtn.Content = "Скрыть";
                isCollapsed = false;
            }
        }

        private void colorBtn_Click(object sender, RoutedEventArgs e)
        {
            ColorPickerWindow colorPicker = new ColorPickerWindow();
            colorPicker.Owner = this; // Устанавливаем родительское окно

            if (colorPicker.ShowDialog() == true)
            {
                Color selectedColor = colorPicker.SelectedWpfColor;
                string colorFormat = $"100,{selectedColor.R},{selectedColor.G},{selectedColor.B}";
                colorTB.Text = colorFormat;
            }
        }
        private void LoadTable()
        {
            try
            {
                string query = @"SELECT Текст_объявления, заказчик, дата_подачи, дата_закрытия, цвет, телефон, Код_объявления 
                       FROM Объявления";

                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand(query, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            _advertisements.Clear();
                            DateTime currentDate = DateTime.Now.Date;

                            while (reader.Read())
                            {
                                DateTime dateOpen, dateClose;

                                // Безопасный парсинг дат
                                if (!DateTime.TryParse(reader["дата_подачи"]?.ToString(), out dateOpen))
                                    continue;
                                if (!DateTime.TryParse(reader["дата_закрытия"]?.ToString(), out dateClose))
                                    continue;

                                // Проверка условия в коде C#
                                if (currentDate >= dateOpen.Date && currentDate <= dateClose.Date)
                                {
                                    var advertisement = new Advertisement
                                    {
                                        ТекстОбъявления = reader["Текст_объявления"]?.ToString() ?? "",
                                        Заказчик = reader["заказчик"]?.ToString() ?? "",
                                        ДатаПодачи = dateOpen,
                                        ДатаЗакрытия = dateClose,
                                        Цвет = reader["цвет"]?.ToString() ?? "100,143,143,143",
                                        Телефон = reader["телефон"]?.ToString() ?? "",
                                        КодОбъявления = reader["Код_объявления"]?.ToString() ?? ""
                                    };
                                    _advertisements.Add(advertisement);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка");
            }
        }
        public void FilterByDateRange(DateTime startDate, DateTime endDate)
        {
            try
            {
                // Фильтрация по диапазону дат
                string query = @"SELECT Текст_объявления, заказчик, дата_подачи, дата_закрытия, цвет, телефон, Код_объявления 
                               FROM Объявления 
                               WHERE DATE(дата_подачи) >= DATE(@startDate) AND DATE(дата_закрытия) <= DATE(@endDate)";

                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();

                    using (var cmd = new SQLiteCommand(query, conn))
                    {
                        // Параметры дат в формате ISO8601
                        cmd.Parameters.AddWithValue("@startDate", startDate.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@endDate", endDate.ToString("yyyy-MM-dd"));

                        using (var reader = cmd.ExecuteReader())
                        {
                            _advertisements.Clear();

                            while (reader.Read())
                            {
                                var advertisement = new Advertisement
                                {
                                    ТекстОбъявления = reader["Текст_объявления"]?.ToString() ?? "",
                                    Заказчик = reader["заказчик"]?.ToString() ?? "",
                                    ДатаПодачи = DateTime.Parse(reader["дата_подачи"].ToString()),
                                    ДатаЗакрытия = DateTime.Parse(reader["дата_закрытия"].ToString()),
                                    Цвет = reader["цвет"]?.ToString() ?? "100,143,143,143",
                                    Телефон = reader["телефон"]?.ToString() ?? "",
                                    КодОбъявления = reader["Код_объявления"]?.ToString() ?? ""
                                };

                                _advertisements.Add(advertisement);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка фильтрации: {ex.Message}", "Ошибка");
            }
        }
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (dataGrid1.SelectedItem is Advertisement selectedAd)
            {
                var result = MessageBox.Show(
                    $"Вы уверены, что хотите удалить объявление:\n\n\"{selectedAd.ТекстОбъявления}\"?",
                    "Подтверждение удаления",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        string deleteQuery = "DELETE FROM [Объявления] WHERE [Код_объявления] = @code";
                        using (SQLiteConnection conn = new SQLiteConnection(connStr))
                        {
                            conn.Open();
                            using (SQLiteCommand cmd = new SQLiteCommand(deleteQuery, conn))
                            {
                                cmd.Parameters.AddWithValue("@code", selectedAd.КодОбъявления);
                                int rowsAffected = cmd.ExecuteNonQuery();
                                if (rowsAffected > 0)
                                {
                                    _advertisements.Remove(selectedAd);
                                    MessageBox.Show("Объявление успешно удалено.", "Успех");
                                }
                                else
                                {
                                    MessageBox.Show("Объявление не найдено или уже удалено.", "Предупреждение");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите объявление для удаления.", "Предупреждение");
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Проверка заполнения обязательных полей
                if (string.IsNullOrWhiteSpace(adTB.Text))
                {
                    MessageBox.Show("Заполните текст объявления", "Ошибка");
                    return;
                }

                if (string.IsNullOrWhiteSpace(orderCB.Text))
                {
                    MessageBox.Show("Выберите заказчика", "Ошибка");
                    return;
                }

                // Получение количества записей для добавления
                if (!int.TryParse(NumberOfControls.Text, out int numberOfRecords) || numberOfRecords <= 0)
                {
                    MessageBox.Show("Введите корректное количество записей", "Ошибка");
                    return;
                }

                // Получение общих данных
                string adText = adTB.Text.Trim();
                string customer = orderCB.Text.Trim();
                string phone = phoneTB.Text.Trim();
                string color = string.IsNullOrWhiteSpace(colorTB.Text) ? "100,143,143,143" : colorTB.Text.Trim();

                // Найти все DateControl на форме
                var dateControls = FindDateControls();

                if (dateControls.Count < numberOfRecords)
                {
                    MessageBox.Show($"Недостаточно элементов управления датами. Найдено: {dateControls.Count}, требуется: {numberOfRecords}", "Ошибка");
                    return;
                }

                // Добавление записей в базу данных
                string insertQuery = @"INSERT INTO [Объявления] 
                             ([Текст_объявления], [заказчик], [дата_подачи], [дата_закрытия], [цвет], [телефон]) 
                             VALUES (@text, @customer, @dateOpen, @dateClose, @color, @phone)";

                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            for (int i = 0; i < numberOfRecords; i++)
                            {
                                var dateControl = dateControls[i];
                                var dates = GetDatesFromControl(dateControl);

                                if (!dates.dateOpen.HasValue || !dates.dateClose.HasValue)
                                {
                                    MessageBox.Show($"Не заполнены даты в DateControl {i + 1}", "Ошибка");
                                    transaction.Rollback();
                                    return;
                                }

                                if (dates.dateOpen.Value > dates.dateClose.Value)
                                {
                                    MessageBox.Show($"Дата открытия не может быть позже даты закрытия (DateControl {i + 1})", "Ошибка");
                                    transaction.Rollback();
                                    return;
                                }

                                using (var cmd = new SQLiteCommand(insertQuery, conn, transaction))
                                {
                                    cmd.Parameters.AddWithValue("@text", adText);
                                    cmd.Parameters.AddWithValue("@customer", customer);
                                    cmd.Parameters.AddWithValue("@dateOpen", dates.dateOpen.Value.Date);
                                    cmd.Parameters.AddWithValue("@dateClose", dates.dateClose.Value.Date);
                                    cmd.Parameters.AddWithValue("@color", color);
                                    cmd.Parameters.AddWithValue("@phone", string.IsNullOrWhiteSpace(phone) ? (object)DBNull.Value : phone);

                                    cmd.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                            MessageBox.Show($"Успешно добавлено {numberOfRecords} объявлений", "Успех");

                            // Очистить поля после успешного добавления
                            ClearFields();

                            // Обновить список объявлений
                            LoadTable();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            throw ex;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении объявлений: {ex.Message}", "Ошибка");
            }
        }

        // Метод для поиска DateControl элементов на форме
        private List<DateControl> FindDateControls()
        {
            var dateControls = new List<DateControl>();

            // Поиск всех DateControl элементов на форме
            foreach (var child in GetAllChildren(this))
            {
                if (child is DateControl dateControl)
                {
                    dateControls.Add(dateControl);
                }
            }

            // Сортировка по имени для правильного порядка
            return dateControls.OrderBy(dc => dc.Name).ToList();
        }

        // Рекурсивный поиск всех дочерних элементов
        private IEnumerable<FrameworkElement> GetAllChildren(DependencyObject parent)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is FrameworkElement frameworkElement)
                    yield return frameworkElement;

                foreach (var descendant in GetAllChildren(child))
                    yield return descendant;
            }
        }

        // Альтернативный метод для получения дат из DateControl
        private (DateTime? dateOpen, DateTime? dateClose) GetDatesFromControl(DateControl dateControl)
        {
            try
            {
                var dateOpenPicker = FindChildByName<DatePicker>(dateControl, "dateOpen");
                var dateClosePicker = FindChildByName<DatePicker>(dateControl, "dateClose");

                return (dateOpenPicker?.SelectedDate, dateClosePicker?.SelectedDate);
            }
            catch
            {
                return (null, null);
            }
        }

        // Поиск дочернего элемента по имени
        private T FindChildByName<T>(DependencyObject parent, string name) where T : FrameworkElement
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T element && element.Name == name)
                    return element;

                var result = FindChildByName<T>(child, name);
                if (result != null)
                    return result;
            }

            return null;
        }

        // Очистка полей после добавления
        private void ClearFields()
        {
            adTB.Clear();
            phoneTB.Clear();
            colorTB.Clear();
            orderCB.SelectedIndex = -1;

            // Очистить DateControl элементы
            var dateControls = FindDateControls();
            foreach (var dateControl in dateControls)
            {
                var dateOpenPicker = FindChildByName<DatePicker>(dateControl, "dateOpen");
                var dateClosePicker = FindChildByName<DatePicker>(dateControl, "dateClose");

                if (dateOpenPicker != null)
                    dateOpenPicker.SelectedDate = null;
                if (dateClosePicker != null)
                    dateClosePicker.SelectedDate = null;
            }
        }
        private void ExportFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (dataGrid1.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.");
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = System.IO.Path.Combine(desktopPath, "Бегунки");
            Directory.CreateDirectory(folderPath);

            string timestamp = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            string filePath = System.IO.Path.Combine(folderPath, $"{timestamp}.txt");

            Encoding ansiEncoding = Encoding.GetEncoding(1251);

            using (StreamWriter begunok = new StreamWriter(filePath, false, ansiEncoding))
            {
                foreach (var item in dataGrid1.Items)
                {
                    if (item == CollectionView.NewItemPlaceholder || item == null)
                        continue;

                    var type = item.GetType();

                    string text = type.GetProperty("ТекстОбъявления")?.GetValue(item)?.ToString() ?? "";
                    string phone = type.GetProperty("Телефон")?.GetValue(item)?.ToString() ?? "";
                    string color = type.GetProperty("Цвет")?.GetValue(item)?.ToString() ?? "";

                    // Очистка текста
                    text = text.Replace("\"", "").Replace("\r", " ").Replace("\n", " ");
                    phone = phone.Replace("\r", "").Replace("\n", "");
                    color = color.Replace("\r", "").Replace("\n", "");

                    if (string.IsNullOrWhiteSpace(phone))
                    {
                        begunok.WriteLine($" {text}");
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(color))
                        {
                            color = "100,143,143,143";
                        }

                        begunok.WriteLine($" {text}|{phone}<pb {color}>");
                    }
                }
            }

            MessageBox.Show($"Файл успешно создан:\n{filePath}", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void exportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("Excel не установлен на этом компьютере.");
                return;
            }

            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            // Настройка ширины колонок
            worksheet.Columns["A"].ColumnWidth = 100;
            worksheet.Columns["B"].ColumnWidth = 12;
            worksheet.Columns["C"].ColumnWidth = 13;
            worksheet.Columns["D"].ColumnWidth = 14;
            worksheet.Columns["E"].ColumnWidth = 14;
            worksheet.Columns["F"].ColumnWidth = 20;

            // Заголовки
            for (int i = 0; i < dataGrid1.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = dataGrid1.Columns[i].Header.ToString();
            }

            // Данные
            for (int i = 0; i < dataGrid1.Items.Count; i++)
            {
                var item = dataGrid1.Items[i];

                // Пропускаем пустые строки или новые строки
                if (item == null || item == CollectionView.NewItemPlaceholder)
                    continue;

                for (int j = 0; j < dataGrid1.Columns.Count; j++)
                {
                    // Получаем значение ячейки
                    var cellValue = GetCellValue(item, dataGrid1.Columns[j]);

                    if (cellValue != null)
                    {
                        Excel.Range cell = worksheet.Cells[i + 2, j + 1];
                        string stringValue;

                        // Проверяем, является ли значение датой
                        if (cellValue is DateTime dateValue)
                        {
                            // Для дат убираем время, оставляем только дату
                            stringValue = dateValue.ToString("dd.MM.yyyy");
                        }
                        else
                        {
                            stringValue = cellValue.ToString();
                        }

                        stringValue = CleanQuotes(stringValue);

                        // Удаляем переносы строк
                        stringValue = stringValue.Replace("\r", " ").Replace("\n", " ");

                        cell.Clear();
                        cell.ClearFormats();
                        cell.NumberFormat = "@";
                        cell.WrapText = false;
                        cell.ShrinkToFit = false;
                        cell.Formula = "'" + stringValue;
                    }
                }
            }

            // Создаем папку "бегунки" на рабочем столе
            string datetext = Convert.ToString(DateTime.Now);
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = Path.Combine(desktopPath, "Бегунки Excel");

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string dateText = datetext.Replace(" ", "_").Replace(":", "_").Replace(".", "_");
            string filePath = Path.Combine(folderPath, $"{dateText}.xlsx");

            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();

            // Освобождение COM объектов
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            MessageBox.Show($"Данные успешно экспортированы в Excel!\nФайл сохранен в: {folderPath}", "Экспорт",
                           MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Метод для получения значения ячейки из DataGrid
        private object GetCellValue(object item, DataGridColumn column)
        {
            if (column is DataGridBoundColumn boundColumn)
            {
                var binding = boundColumn.Binding as System.Windows.Data.Binding;
                if (binding != null && !string.IsNullOrEmpty(binding.Path.Path))
                {
                    var propertyInfo = item.GetType().GetProperty(binding.Path.Path);
                    if (propertyInfo != null)
                    {
                        return propertyInfo.GetValue(item);
                    }
                }
            }
            else if (column is DataGridTextColumn textColumn)
            {
                var binding = textColumn.Binding as System.Windows.Data.Binding;
                if (binding != null && !string.IsNullOrEmpty(binding.Path.Path))
                {
                    var propertyInfo = item.GetType().GetProperty(binding.Path.Path);
                    if (propertyInfo != null)
                    {
                        return propertyInfo.GetValue(item);
                    }
                }
            }

            return null;
        }

        private string CleanQuotes(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Заменяем экранированные кавычки на обычные
            string cleaned = input.Replace("\\\"", "\"");  // \" -> "

            // Заменяем угловые кавычки на обычные
            cleaned = cleaned.Replace("«", "\"");          // « -> "
            cleaned = cleaned.Replace("»", "\"");          // » -> "
            cleaned = cleaned.Replace("„", "\"");          // „ -> "
            cleaned = cleaned.Replace("‚", "'");           // ‚ -> '
            cleaned = cleaned.Replace("'", "'");           // ' -> '
            cleaned = cleaned.Replace("'", "'");           // ' -> '

            return cleaned;
        }
        private string CleanText(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            return input.Replace("\"", "").Trim();
        }

        private void searchButton_Click(object sender, RoutedEventArgs e)
        {
            string query = "SELECT [Текст_объявления], [заказчик], [дата_подачи], [дата_закрытия], [цвет], [телефон], [Код_объявления] FROM Объявления";

            if (string.IsNullOrEmpty(filterTB.Text) && dateRB2.IsChecked == true)
            {
                // Для SQLite используем правильный формат даты
                string dateValue = dateFilter.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");
                query = $"SELECT [Текст_объявления], [заказчик], [дата_подачи], [дата_закрытия], [цвет], [телефон], [Код_объявления] FROM [Объявления] WHERE date('{dateValue}') >= date([дата_подачи]) AND date('{dateValue}') <= date([дата_закрытия])";
            }
            else
            {
                bool hasWhereClause = false;

                if (filterRB1.IsChecked == true && !string.IsNullOrEmpty(filterTB.Text))
                {
                    query += $" WHERE [Текст_объявления] LIKE '%{filterTB.Text.Replace("'", "''")}%'";
                    hasWhereClause = true;
                }
                else if (filterRB2.IsChecked == true && !string.IsNullOrEmpty(filterTB.Text))
                {
                    query += $" WHERE [заказчик] LIKE '%{filterTB.Text.Replace("'", "''")}%'";
                    hasWhereClause = true;
                }

                if (dateRB1.IsChecked == true)
                {
                    string dateValue1 = dateFilter.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");
                    string dateValue2 = dateFilter2.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");

                    if (hasWhereClause)
                        query += $" AND date([дата_подачи]) BETWEEN date('{dateValue1}') AND date('{dateValue2}')";
                    else
                    {
                        query += $" WHERE date([дата_подачи]) BETWEEN date('{dateValue1}') AND date('{dateValue2}')";
                        hasWhereClause = true;
                    }
                }

                if (dateRB3.IsChecked == true)
                {
                    string dateValue = dateFilter.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");

                    if (hasWhereClause)
                        query += $" AND date([дата_закрытия]) = date('{dateValue}')";
                    else
                    {
                        query += $" WHERE date([дата_закрытия]) = date('{dateValue}')";
                        hasWhereClause = true;
                    }
                }

                if (dateRB4.IsChecked == true)
                {
                    string dateValue1 = dateFilter.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");
                    string dateValue2 = dateFilter2.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");

                    if (hasWhereClause)
                        query += $" AND date([дата_закрытия]) BETWEEN date('{dateValue1}') AND date('{dateValue2}')";
                    else
                        query += $" WHERE date([дата_закрытия]) BETWEEN date('{dateValue1}') AND date('{dateValue2}')";
                }
            }

            try
            {
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand(query, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            _advertisements.Clear();
                            DateTime currentDate = DateTime.Now.Date;

                            while (reader.Read())
                            {
                                DateTime dateOpen, dateClose;

                                // Безопасный парсинг дат
                                if (!DateTime.TryParse(reader["дата_подачи"]?.ToString(), out dateOpen))
                                    continue;
                                if (!DateTime.TryParse(reader["дата_закрытия"]?.ToString(), out dateClose))
                                    continue;
                                var advertisement = new Advertisement
                                {
                                    ТекстОбъявления = reader["Текст_объявления"]?.ToString() ?? "",
                                    Заказчик = reader["заказчик"]?.ToString() ?? "",
                                    ДатаПодачи = dateOpen,
                                    ДатаЗакрытия = dateClose,
                                    Цвет = reader["цвет"]?.ToString() ?? "100,143,143,143",
                                    Телефон = reader["телефон"]?.ToString() ?? "",
                                    КодОбъявления = reader["Код_объявления"]?.ToString() ?? ""
                                };
                                _advertisements.Add(advertisement);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выполнении запроса: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Вспомогательный метод для конвертации дат
        private string ConvertToShortDateString(object dateValue)
        {
            if (dateValue == null || dateValue == DBNull.Value)
                return "";

            if (DateTime.TryParse(dateValue.ToString(), out DateTime date))
                return date.ToShortDateString();

            return dateValue.ToString();
        }

        private void dateRB1_Checked(object sender, RoutedEventArgs e)
        {
            if (dateRB1.IsChecked == true && dateFilter2.Width == 0)
            {
                DoubleAnimation RBAnimation = new DoubleAnimation();
                RBAnimation.From = 0;
                RBAnimation.To = 130;
                RBAnimation.Duration = TimeSpan.FromSeconds(0.5);
                RBAnimation.EasingFunction = new CubicEase();

                dateFilter2.BeginAnimation(WidthProperty, RBAnimation);
            }
        }
        private void dateRB2_Checked(object sender, RoutedEventArgs e)
        {
            if (dateRB2.IsChecked == true && dateFilter2.Width == 130)
            {
                DoubleAnimation RBAnimation = new DoubleAnimation();
                RBAnimation.From = 130;
                RBAnimation.To = 0;
                RBAnimation.Duration = TimeSpan.FromSeconds(0.5);
                RBAnimation.EasingFunction = new CubicEase();

                dateFilter2.BeginAnimation(WidthProperty, RBAnimation);
            }
        }
        private void dateRB3_Checked(object sender, RoutedEventArgs e)
        {
            if (dateRB3.IsChecked == true && dateFilter2.Width == 130)
            {
                DoubleAnimation RBAnimation = new DoubleAnimation();
                RBAnimation.From = 130; 
                RBAnimation.To = 0;
                RBAnimation.Duration = TimeSpan.FromSeconds(0.5);
                RBAnimation.EasingFunction = new CubicEase();

                dateFilter2.BeginAnimation(WidthProperty, RBAnimation);
            }
        }
        private void dateRB4_Checked(object sender, RoutedEventArgs e)
        {
            if (dateRB4.IsChecked == true && dateFilter2.Width == 0)
            {
                DoubleAnimation RBAnimation = new DoubleAnimation();
                RBAnimation.From = 0;
                RBAnimation.To = 130;
                RBAnimation.Duration = TimeSpan.FromSeconds(0.5);
                RBAnimation.EasingFunction = new CubicEase();

                dateFilter2.BeginAnimation(WidthProperty, RBAnimation);
            }
        }
    }
    public class Advertisement : INotifyPropertyChanged
    {
        private string _текстОбъявления;
        private string _заказчик;
        private DateTime _датаПодачи;
        private DateTime _датаЗакрытия;
        private string _цвет;
        private string _телефон;
        private string _кодОбъявления;

        public string ТекстОбъявления
        {
            get => _текстОбъявления;
            set
            {
                _текстОбъявления = value;
                OnPropertyChanged();
            }
        }

        public string Заказчик
        {
            get => _заказчик;
            set
            {
                _заказчик = value;
                OnPropertyChanged();
            }
        }
        public DateTime ДатаПодачи
        {
            get => _датаПодачи;
            set
            {
                _датаПодачи = value;
                OnPropertyChanged();
            }
        }
        public DateTime ДатаЗакрытия
        {
            get => _датаЗакрытия;
            set
            {
                _датаЗакрытия = value;
                OnPropertyChanged();
            }
        }
        public string КодОбъявления
        {
            get => _кодОбъявления;
            set
            {
                _кодОбъявления = value;
                OnPropertyChanged();
            }
        }
        public string Цвет
        {
            get => _цвет;
            set
            {
                _цвет = value;
                OnPropertyChanged();
            }
        }
        public string Телефон
        {
            get => _телефон;
            set
            {
                _телефон = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    public class Obyavlenie
    {
        public string Текст_объявления { get; set; }
        public string Телефон { get; set; }
        public string Цвет { get; set; }
    }

}
