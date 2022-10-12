using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TextEditor
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private Dictionary<string, string[]> FileFilter = new Dictionary<string, string[]>()
        {
            { "Документы Word", new[] { "*.docx", "*.doc" } },
            { "Текстовые файлы", new[] { "*.txt" } }
        };

        private string Filter => string.Join("|", FileFilter.Select(keyPair => $"{keyPair.Key} ({string.Join(", ", keyPair.Value)})|{string.Join(";", keyPair.Value)}"));
        private string DefaultFilter => FileFilter.Select(keyPair => $"{keyPair.Key} ({string.Join(", ", keyPair.Value)})|{string.Join(";", keyPair.Value)}").First();

        private readonly MenuItem[] menuItems;

        public const int DefaultScale = 14;
        private int _scale = DefaultScale;

        public int Scale
        {
            get => _scale;
            set
            {
                _scale = value;
                TB.FontSize = _scale;
            }
        }


        private readonly Microsoft.Office.Interop.Word.Application _application;
        private Document _activeDocument;

        public Document ActiveDocument
        {
            get => _activeDocument;
            set
            {
                _activeDocument = value;
                _isSaved = false;
            }
        }

        private bool _isSaved = false;
        private string _path = null;

        private FindWindow _findWindow;

        public MainWindow()
        {
            InitializeComponent();

            menuItems = new MenuItem[TopMenu.Items.Count];
            for (int i = 0; i < menuItems.Length; i++)
            {
                if (TopMenu.Items[i] is MenuItem menuItem)
                    if (menuItem.IsEnabled == false)
                        menuItems[i] = menuItem;
            
            }

            _application = new Microsoft.Office.Interop.Word.Application();

            _findWindow = new FindWindow();
            _findWindow.Show();
        }

        /// <summary>
        /// Создаёт новый документ, если в текущем окне уже есть созданный документ, 
        /// то пытается сохранить и замещает новым документом, иначе ничего не создаёт
        /// </summary>
        private void CreateNewDocument()
        {
            if (ActiveDocument == null)
            {
                AddDocument();
                TB.Text = string.Empty;
            }
            else
            {
                SaveDocument();
                if (_isSaved == false)
                    return;

                CloseDocument();

                AddDocument();
                ReadDocument();
            }
        }

        /// <summary>
        /// Добавляет новый документ в application
        /// </summary>
        private void AddDocument()
        {
            object template = Type.Missing;
            object newTemplate = false;
            object documentType = WdNewDocumentType.wdNewBlankDocument;
            object visible = true;
            ActiveDocument = _application.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
        }

        /// <summary>
        /// Открывает документ
        /// </summary>
        /// <exception cref="FileNotFoundException">Если пытаемся открыть не существующий файл</exception>
        private void OpenDocument()
        {
            if (ActiveDocument != null)
            {
                SaveDocument();
                if (_isSaved == false)
                    return;

                CloseDocument();
            }

            if ((_path = CreateOpenFileDialogMenu()) == null)
                return;

            if (File.Exists(_path) == false)
                throw new FileNotFoundException();

            _activeDocument = _application.Documents.Open(_path);
            ReadDocument();
        }

        private void SaveDocument()
        {
            if (_isSaved)
                return;

            if (ActiveDocument == null)
                return;

            if (_path == null)
            {
                SaveAsDocument();
            }
            else
            {
                ActiveDocument.Content.Text = TB.Text;
                try
                {
                    ActiveDocument.Save();
                    _isSaved = true;
                }
                catch(Exception)
                {
                    return;
                }
            }
        }

        private void SaveAsDocument()
        {
            if (_isSaved)
                return;

            if (ActiveDocument == null)
                return;

            _path = CreateSaveFileDialogMenu();
            if (_path == null)
                return;

            ActiveDocument.Content.Text = TB.Text;
            object pathObject = _path;
            ActiveDocument.SaveAs2(ref pathObject);
            _isSaved = true;
        }

        private string CreateSaveFileDialogMenu()
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Filter = Filter,
                DefaultExt = DefaultFilter,
            };

            if (dialog.ShowDialog() == true)
                return dialog.FileName;
            return null;
        }

        private string CreateOpenFileDialogMenu()
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Filter = Filter,
                DefaultExt = DefaultFilter
            };

            if (dialog.ShowDialog() == true)
                return dialog.FileName;
            return null;
        }

        private void ReadDocument()
        {
            if (ActiveDocument == null)
                return;

            TB.Text = ActiveDocument.Content.Text;
        }

        private void CloseDocument()
        {
            if (ActiveDocument == null)
                return;

            MessageBoxResult wantSave = MessageBox.Show("Хотите сохранить?", "Сохранить изменения?", MessageBoxButton.YesNoCancel);
            if (wantSave == MessageBoxResult.Cancel)
                return;
            else if (wantSave == MessageBoxResult.Yes)
                SaveDocument();

            ActiveDocument.Close();
            ActiveDocument = null;
        }

        private void QuitApplication()
        {
            CloseDocument();
            _application.Quit();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CloseDocument();
            if (ActiveDocument == null)
                _application.Quit();
            else
                e.Cancel = true;

            if (_findWindow != null)
                _findWindow.Close();
        }

        private void TB_TextChanged(object sender, TextChangedEventArgs e)
        {
            _isSaved = false;
            if (ActiveDocument == null)
                CreateNewDocument();
        }

        private void CreateNewDocument_Click(object sender, RoutedEventArgs e) => CreateNewDocument();
        private void NewWindow_Click(object sender, RoutedEventArgs e) => new MainWindow().Show();
        private void OpenDocument_Click(object sender, RoutedEventArgs e) => OpenDocument();
        private void SaveDocument_Click(object sender, RoutedEventArgs e) => SaveDocument();
        private void SaveAsDocument_Click(object sender, RoutedEventArgs e) => SaveAsDocument();
        private void Exit_Click(object sender, RoutedEventArgs e) => Close();
        private void Undo_Click(object sender, RoutedEventArgs e) => TB.Undo();
        private void Cut_Click(object sender, RoutedEventArgs e) => TB.Cut();
        private void Copy_Click(object sender, RoutedEventArgs e) => TB.Copy();
        private void Paste_Click(object sender, RoutedEventArgs e) => TB.Paste();
        private void Delete_Click(object sender, RoutedEventArgs e) => TB.Text = TB.Text.Remove(TB.SelectionStart, TB.SelectionLength);

        private void TB_SelectionChanged(object sender, RoutedEventArgs e)
        {
            bool enabled = TB.SelectionLength > 0;
            if (enabled == false)
            {
                foreach (var item in menuItems)
                    if (item != null)
                        item.IsEnabled = false;
                return;
            }
            foreach (var menuItem in menuItems)
                if (menuItem != null)
                    menuItem.IsEnabled = enabled;
        }

        private void ZoomIn_Click(object sender, RoutedEventArgs e) => Scale = (int)(Scale * 1.3);
        private void ZoomOut_Click(object sender, RoutedEventArgs e) => Scale = (int)(Scale / 1.3);
        private void DefaultZomm_Click(object sender, RoutedEventArgs e) => Scale = DefaultScale;
    }
}
