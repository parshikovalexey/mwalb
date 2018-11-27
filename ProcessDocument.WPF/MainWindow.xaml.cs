using System.Collections.Generic;
using System.ComponentModel;
using CommonLibrary;
using Microsoft.Win32;
using ProcessDocumentCore;
using ProcessDocumentCore.Processing;
using StandardsLibrary;
using System.Windows;
using StandardsLibrary.Simple;

namespace ProcessDocument.WPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

        GostService _gostRepository = new GostService();
        private List<SimpleHeaderGost> _gosts;
        private SimpleHeaderGost _selectGost;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            Gosts = new List<SimpleHeaderGost>();
            _gostRepository.LoadGostFromFile();
            loadGosts();
        }

        private void loadGosts()
        {
            Gosts = _gostRepository.GetGostList();
        }

        public List<SimpleHeaderGost> Gosts
        {
            get => _gosts;
            set { _gosts = value; NotifyPropertyChanged("Gosts"); }
        }

        public SimpleHeaderGost SelectGost
        {
            get => _selectGost;
            set
            {
                _selectGost = value;
                NotifyPropertyChanged("SelectGost");
            }
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

            if (CheckInput()) return;
          var selectedGost =   _gostRepository.GetGostModel(SelectGost.GuidGost);
            var d = new GostGenericRepository<GostModel>(selectedGost);
            Standards gost = TestGostCheck.IsChecked != null && !(bool) TestGostCheck.IsChecked
                ? (Standards) new Gost1(GostPath)
                : new GostTest();
            _ = new Execute(FilePath, gost, new ProcessingOpenXml(), ResultDocument);
        }

        private bool CheckInput()
        {
            if (string.IsNullOrEmpty(GostPath) && TestGostCheck.IsChecked == false)
            {
                MessageBox.Show("Гост не выбран!");
                return true;
            }

            if (string.IsNullOrEmpty(FilePath))
            {
                MessageBox.Show("Файл документа не выбран!");
                return true;
            }

            return false;
        }

        private void ResultDocument(ResultExecute resultexecute)
        {
            returnStatusExecute = resultexecute;
            if (resultexecute.StatusExecute == ResultExecute.Status.Success)
            {
                if (returnStatusExecute.Callbacks is string filePath)
                {
                    System.Diagnostics.Process.Start(filePath);
                }
            }
            else
            {
                MessageBox.Show(resultexecute.ErrorMsg);
            }
            //ResultExecuteTextBox.Text = resultexecute.ToString();
           
        }

        private ResultExecute returnStatusExecute { get; set; }

        private void OpenDocument_Click(object sender, RoutedEventArgs e)
        {
            if (returnStatusExecute.Callbacks is string filePath)
            {
                System.Diagnostics.Process.Start(filePath);
            }
           
        }
        private void LoadGostButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON (*.json)|*.json";
            if (openFileDialog.ShowDialog() == true)
                GostPath = openFileDialog.FileName;

        }

        private void ToggleButton_OnCheckedTest(object sender, RoutedEventArgs e)
        {
            StackPanelLoadGost.IsEnabled = TestGostCheck.IsChecked == null || !(bool) TestGostCheck.IsChecked;
        }
    }
}
