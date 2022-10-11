using Microsoft.Office.Interop.Word;
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
        private readonly Microsoft.Office.Interop.Word.Application _application = new Microsoft.Office.Interop.Word.Application();
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

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenDocument(string path)
        {
            if (File.Exists(path) == false)
                throw new FileNotFoundException();

            _activeDocument = _application.Documents.Open(path);
            ReadDocument();
        }

        private void ReadDocument()
        {
            if (ActiveDocument == null)
                return;

            TB.Text = ActiveDocument.Content.Text;
        }

        private void CloseDocument()
        {
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
