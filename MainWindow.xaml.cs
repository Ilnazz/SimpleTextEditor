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

        public MainWindow()
        {
            InitializeComponent();
            _application = new Microsoft.Office.Interop.Word.Application();
        }

        private void CreateNewDocument()
        {

        }

        private void OpenDocument()
        {
            _path = CreateOpenFileDialogMenu();
            if (_path == null)
                return;

            if (File.Exists(_path) == false)
                throw new FileNotFoundException();

            _activeDocument = _application.Documents.Open(_path);
            ReadDocument();
        }

        private void SaveDocument()
        {
            if (_path == null)
                SaveAsDocument();
            else
                ActiveDocument.Save();
        }

        private void SaveAsDocument()
        {
            _path = CreateSaveFileDialogMenu();
            if (_path == null)
                return;

            ActiveDocument.SaveAs2(_path);
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

            ActiveDocument.Close();
            ActiveDocument = null;
        }

        private void QuitApplication()
        {
            CloseDocument();
            _application.Quit();
        }

        private void TB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            var prevSelectionStart = TB.SelectionStart;
            TB.Text = TB.Text.Insert(TB.SelectionStart, "\n");
            TB.SelectionStart = prevSelectionStart + 1;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            QuitApplication();
        }
    }
}
