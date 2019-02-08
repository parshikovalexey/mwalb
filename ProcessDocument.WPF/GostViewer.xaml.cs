using StandardsLibrary;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using StandardsLibrary.Simple;

namespace ProcessDocument.WPF
{
    /// <summary>
    /// Логика взаимодействия для GostViewer.xaml
    /// </summary>
    public partial class GostViewer : MetroWindow, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }



        GostModel _currentGostModel = new GostModel();
        private List<GostModelViewer> _view;

        public List<GostModelViewer> View
        {
            get => _view;
            set { _view = value; NotifyPropertyChanged("View"); }
        }

        public GostViewer(GostModel gostModel)
        {
            InitializeComponent();
            DataContext = this;
            _currentGostModel = gostModel;
            View = GetModelForView();
            NotifyPropertyChanged("View");
        }

        private List<GostModelViewer> GetModelForView()
        {
            var result = new List<GostModelViewer>();

            result.Add(new GostModelViewer(){ Header="Стиль основного текста", Model=_currentGostModel.GlobalText, ListDescription = new List<SimpleGostModelViewer>()
            {
                new SimpleGostModelViewer()
                {
                    Description = "Test",
                    Value = "123"
                }
            }});
            result.Add(new GostModelViewer(){ Header= "Стиль изображений", Model=_currentGostModel.Headline});
            result.Add(new GostModelViewer(){ Header= "Стиль заголовков", Model=_currentGostModel.Image});



            return result;
        }
    }

    public class GostModelViewer
    {
        public string Header { get; set; }
       public List<SimpleGostModelViewer> ListDescription = new List<SimpleGostModelViewer>();
        public SimpleStyle Model { get; set; }
    }

    public class SimpleGostModelViewer
    {
        public string Description { get; set; }
        public string Value { get; set; }
    }
}
