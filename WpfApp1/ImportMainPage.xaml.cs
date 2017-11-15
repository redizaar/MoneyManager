using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for ImportMainPage.xaml
    /// </summary>
    public partial class ImportMainPage : Page
    {
        private ButtonCommands btnCommand;
        private MainWindow mainWindow;
        private User currentUser;
        private static ImportMainPage instance;
        private ImportMainPage(MainWindow mainWindow)
        {
            DataContext = this;

            InitializeComponent();
            this.mainWindow = mainWindow;
            this.currentUser = mainWindow.getCurrentUser();
            if (currentUser.getAccountNumber().Equals(mainWindow.getAccounNumber()))
            {
                getUserStatistics(currentUser);
            }
            else
            {
                getUserStatistics(mainWindow.getAccounNumber());
            }
            FolderAddressLabel.Visibility = System.Windows.Visibility.Hidden;
        }

        public void getUserStatistics(string accountNumber)
        {
            int numberOfTransactions = 0;
            string lastImportDate = "";
            string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
            DateTime todayDate = Convert.ToDateTime(todaysDate);
            usernameLabel.Content = "Test";
            foreach (var transactions in SavedTransactions.getSavedTransactions())
            {
                if (transactions.getAccountNumber().Equals(accountNumber))
                {
                    numberOfTransactions++;
                    lastImportDate = transactions.getWriteDate();//always overwrites it --- todo (more logic needed lulz)
                }
            }
            if (lastImportDate.Length > 12)
            {
                lastImportDateLabel.Content = lastImportDate.Substring(0, 12);
            }
            else
            {
                lastImportDateLabel.Content = lastImportDate;
            }
            noTransactionsLabel.Content = numberOfTransactions;
            if (lastImportDate.Length > 0)
            {
                DateTime importDate = Convert.ToDateTime(lastImportDate);
                int diffDays = (todayDate - importDate).Days;
                if (diffDays >= 30)
                {
                    urgencyLabel.Content = "Very urgent!";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(217, 30, 24));
                }
                else
                {
                    urgencyLabel.Content = "Not urgent";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(46, 204, 113));
                }
            }
            else
            {
                urgencyLabel.Content = "You haven't imported yet!";
                lastImportDateLabel.Content = "You haven't imported yet!";
            }
        }

        private void getUserStatistics(User currentUser)
        {
            int numberOfTransactions = 0;
            string latestImportDate = "";
            string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
            DateTime todayDate = Convert.ToDateTime(todaysDate);
            usernameLabel.Content = currentUser.getUsername();
            foreach (var transactions in SavedTransactions.getSavedTransactions())
            {
                if(transactions.getAccountNumber().Equals(currentUser.getAccountNumber()))
                {
                    numberOfTransactions++;
                    latestImportDate = transactions.getWriteDate();//always overwrites it --- todo (more logic needed lulz)
                    string importDatestring = transactions.getWriteDate();
                }
            }
            if (latestImportDate.Length > 12)
            {
                lastImportDateLabel.Content = latestImportDate.Substring(0, 12);
            }
            else
            {
                lastImportDateLabel.Content = latestImportDate;
            }
            noTransactionsLabel.Content = numberOfTransactions;
            DateTime importDate;
            if (latestImportDate.Length > 0)
            {
                importDate = Convert.ToDateTime(latestImportDate);
                float diffTicks = (todayDate - importDate).Days;
                if (diffTicks >= 30)
                {
                    urgencyLabel.Content = "Very urgent!";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(217, 30, 24));
                }
                else
                {
                    urgencyLabel.Content = "Not urgent";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(46, 204, 113));
                }
            }
            else
            {
                urgencyLabel.Content = "You haven't imported yet!";
                lastImportDateLabel.Content = "You haven't imported yet!";
            }

        }
        private void getTransactions(string bankName, string folderAddress)
        {
            new ImportReadIn(bankName, folderAddress, mainWindow);
        }

        public ButtonCommands OpenFilePushed
        {
            get
            {
                btnCommand = new ButtonCommands(FileBrowser.Content.ToString(),this);
                return btnCommand;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public static ImportMainPage getInstance(MainWindow mainWindow)
        {
            if(instance==null)
            {
                instance = new ImportMainPage(mainWindow);
            }
            return instance;
        }
        public class ButtonCommands : ICommand
        {
            private string buttonContent;
            private ImportMainPage importPage;
            public ButtonCommands(string buttonContent,ImportMainPage importPage)
            {
                this.buttonContent = buttonContent;
                this.importPage = importPage;

                this.importPage.PropertyChanged += new PropertyChangedEventHandler(test_PropertyChanged);
            }
            private void test_PropertyChanged(object sender, PropertyChangedEventArgs e)
            {
                if (CanExecuteChanged != null)
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
                if (buttonContent.Equals("Import Transactions"))
                {
                    Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                    dlg.DefaultExt = ".xls";
                    dlg.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xlsm)|*.xlsm";
                    Nullable<bool> result = dlg.ShowDialog();
                    if (result == true)
                    {
                        importPage.FolderAddressLabel.Content = dlg.FileName;
                    }
                    importPage.getTransactions("All", importPage.FolderAddressLabel.Content.ToString());
                }
            }
        }
    }
}
