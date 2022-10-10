using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
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

        public MainWindow()
        {
            InitializeComponent();

            var app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = app.Documents.Open("C:\\Users\\Ильназ\\Desktop\\Test.docx");
            // Loop through all words in the document.
            int count = document.Words.Count;
            string text = "";
            for (int i = 1; i <= count; i++)
            {
                // Write the word.
                string s = document.Words[i].Text;
                text += s;
            }
            // Close word.
            app.Quit();
            TB.Text = text;
        }

        private void TB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            var prevSelectionStart = TB.SelectionStart;
            TB.Text = TB.Text.Insert(TB.SelectionStart, "\n");
            TB.SelectionStart = prevSelectionStart + 1;
        }
    }
}
