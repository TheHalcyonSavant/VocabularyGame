using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows;
using System.Linq;

namespace VocabularyGame
{
    public partial class RecordsWindow : Window
    {
        private string _recordsFn;
        private MainWindow _mainWindow;

        private static Properties.Settings _s = Properties.Settings.Default;

        public bool isSaving = false;
        public ObservableCollection<Record> ocRecordsList = new ObservableCollection<Record>();

#region Constructors & Window Events

        public RecordsWindow(MainWindow owner)
        {
            InitializeComponent();
            Owner = _mainWindow = owner;
        }

        private void Window_Activated(object sender, System.EventArgs e)
        {
            dataGrid.ItemsSource = ocRecordsList;
            if (!isSaving) return;
            gridNewRecord.Visibility = Visibility.Visible;
            lblRecordPoints.Content = (Owner as MainWindow).points;
            txtName.Focus();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Hide();
            isSaving = false;
            e.Cancel = true;
        }

#endregion

        private void btnInsert_Click(object sender, RoutedEventArgs e)
        {
            ocRecordsList.Add(new Record(txtName.Text, (int)lblRecordPoints.Content));
            ocRecordsList = new ObservableCollection<Record>(ocRecordsList.OrderByDescending(x => x.Score));
            dataGrid.ItemsSource = ocRecordsList;
            using (FileStream fs = new FileStream(_recordsFn, FileMode.Create))
                _mainWindow.binFormatter.Serialize(fs, ocRecordsList);

            isSaving = false;
            gridNewRecord.Visibility = Visibility.Collapsed;
            Close();
        }

        public void deserializeOC(string fileName)
        {
            _recordsFn = fileName;
            if (File.Exists(_recordsFn))
                using (FileStream fs = new FileStream(_recordsFn, FileMode.Open))
                    ocRecordsList = _mainWindow.binFormatter.Deserialize(fs) as ObservableCollection<Record>;
        }
    }
}