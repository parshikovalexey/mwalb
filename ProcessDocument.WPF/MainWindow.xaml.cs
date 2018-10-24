using CommonLibrary;
using Microsoft.Win32;
using ProcessDocumentCore;
using StandardsLibrary;
using System.Windows;
using ProcessDocumentCore.Processing;

namespace ProcessDocument.WPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string FilePath {
            get => FilePathTextBox.Text;
            set => FilePathTextBox.Text = value;
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Все документы Word (*.doc;*.docx)|*.doc;*.docx";
            if (openFileDialog.ShowDialog() == true)
                FilePath = openFileDialog.FileName;
        }


        private void FormatDocumet_Click(object sender, RoutedEventArgs e)
        {
            _ = new Execute(FilePath, new Gost1(), new ProcessingOpenXml(), ResultDocument);
        }

        private void ResultDocument(ResultExecute resultexecute)
        {
            ResultExecuteTextBox.Text = resultexecute.ToString();
        }
    }
}
