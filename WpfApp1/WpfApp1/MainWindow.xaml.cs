using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
namespace WpfApp1
{
    public partial class MainWindow : INotifyPropertyChanged
    {
        private ButtonCommands btnCommand;

        public MainWindow()
        {
            DataContext = this;
            InitializeComponent();

            
            startUpReadIn();
            banksComboBox.Visibility = System.Windows.Visibility.Hidden;
            FileBrowser.Visibility = System.Windows.Visibility.Hidden;
            FolderAddressLabel.Visibility = System.Windows.Visibility.Hidden;
            HelpChooseLabel.Visibility = System.Windows.Visibility.Hidden;
            LatestImportDate_Label.Visibility = System.Windows.Visibility.Hidden;

            if (LatestImportDate_Label.Content.Equals("Label"))
            {
                LatestImportDate_Label.Content = "You haven't imported yet!";
            }
        }

        public ButtonCommands ImportPushed
        {
            get
            {
                //if(btnCommand==null)
                //{
                    btnCommand = new ButtonCommands(ImportButton.Content.ToString(),this);
                //}

                return btnCommand;
            }
        }
        public ButtonCommands OpenFilePushed
        {
            get
            {
                btnCommand = new ButtonCommands(FileBrowser.Content.ToString(), this);

                return btnCommand;
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
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String refreshDate=DateTime.Now.ToString("yyyy-MM-dd");
            LatestImportDate_Label.Content = refreshDate;
            if(banksComboBox.SelectedItem!=null)
            {
                FileBrowser.Visibility = System.Windows.Visibility.Visible;
            }
        }
        public void getTransactions(string bankName,string folderAddress)
        {
            new ImportReadIn(bankName, folderAddress);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Content = new TransactionMain();
        }
    }
    public class ButtonCommands : ICommand
    {
        private string buttonContent;
        private MainWindow mainWindow;
        public ButtonCommands(string buttonContent,MainWindow mainWindow)
        {
            this.buttonContent = buttonContent;
            this.mainWindow = mainWindow;

            this.mainWindow.PropertyChanged += new PropertyChangedEventHandler(test_PropertyChanged);
        }
        private void test_PropertyChanged(object sender,PropertyChangedEventArgs e)
        {
            if(CanExecuteChanged!=null)
            {
                CanExecuteChanged(this, EventArgs.Empty);
            }
        }
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            //todo
            return true;
        }

        public void Execute(object parameter)
        {
           if(buttonContent.Equals("Import"))
            {
                mainWindow.banksComboBox.Visibility = System.Windows.Visibility.Visible;
                mainWindow.HelpChooseLabel.Visibility = System.Windows.Visibility.Visible;
            }
           else if(buttonContent.Equals("Open File"))
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.DefaultExt = ".xls";
                dlg.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xlsm)|*.xlsm";
                Nullable<bool> result = dlg.ShowDialog();
                if (result == true)
                {
                    mainWindow.FolderAddressLabel.Content = dlg.FileName;
                }
                mainWindow.getTransactions(mainWindow.banksComboBox.Text, mainWindow.FolderAddressLabel.Content.ToString());
                mainWindow.LatestImportDate_Label.Visibility = System.Windows.Visibility.Visible;
            }
        }
    }
}
