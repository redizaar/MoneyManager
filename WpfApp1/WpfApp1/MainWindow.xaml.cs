using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
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
using System.IO;

namespace WpfApp1
{
    
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            DataContext = this;
            InitializeComponent();

            startUpReadIn();

            banksComboBox.Visibility = System.Windows.Visibility.Hidden;
            FileBrowser.Visibility = System.Windows.Visibility.Hidden;
            FolderAddressLabel.Visibility = System.Windows.Visibility.Hidden;
            HelpChooseLabel.Visibility = System.Windows.Visibility.Hidden;
            if (LatestImportDate_Label.Content.Equals("Label"))
            {
                LatestImportDate_Label.Content = "You haven't imported yet!";
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void startUpReadIn()
        {
            //reading in the already saved transactions
            new SavedTransactions();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            banksComboBox.Visibility= System.Windows.Visibility.Visible;
            HelpChooseLabel.Visibility = System.Windows.Visibility.Visible;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String refreshDate=DateTime.Now.ToString("yyyy-MM-dd");
            LatestImportDate_Label.Content = refreshDate;
            if(banksComboBox.SelectedItem!=null)
            {
                FileBrowser.Visibility = System.Windows.Visibility.Visible;
            }
        }
        private void FileBrowser_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xlsm)|*.xlsm";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                FolderAddressLabel.Content = dlg.FileName;
            }
            getTransactions(banksComboBox.Text,FolderAddressLabel.Content.ToString());
        }
        private void getTransactions(string bankName,string folderAddress)
        {
            new ImportReadIn(bankName, folderAddress);
        }
    }
}
