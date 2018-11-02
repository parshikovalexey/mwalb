using CommonLibrary;
using Microsoft.Win32;
using ProcessDocumentCore;
using ProcessDocumentCore.Processing;
using StandardsLibrary;
using System.Windows;

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

        private string GostPath {
            get => GostFilePathTextBox.Text;
            set => GostFilePathTextBox.Text = value;
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
            _ = new Execute(FilePath, new Gost1(GostPath), new ProcessingOpenXml(), ResultDocument);
        }

        private void ResultDocument(ResultExecute resultexecute)
        {
            ResultExecuteTextBox.Text = resultexecute.ToString();
        }

        private void OpenDocument_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(FilePath);
        }
        private void LoadGostButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON (*.json)|*.json";
            if (openFileDialog.ShowDialog() == true)
                GostPath = openFileDialog.FileName;

        }
    }
}
