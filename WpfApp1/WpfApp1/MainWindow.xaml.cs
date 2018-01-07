using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private ButtonCommands btnCommand;
        private List<Transaction> tableAttributes=null;
        Boolean newImport = false;
        public User currentUser=new User();
        private string accountNumber= "11773470-00817789";
        public MainWindow()
        {
            DataContext = this;
            InitializeComponent();
            LoginFrame.Content = new Login_Page(this);
            tableMenuTop.Visibility = System.Windows.Visibility.Hidden; //importmenu is default
            startUpReadIn();
            currentUser.setUsername("Patrik01");
            currentUser.setAccountNumber("11773470-00817789");
            currentUserLabel.Content = currentUser.getUsername(); //notification label
        }

        public void setTableAttribues(List<Transaction> impoertedTransactions,String accountNumber)
        {
            this.tableAttributes = impoertedTransactions;
            this.accountNumber = accountNumber;
        }
        public void setTableAttribues(List<Transaction> impoertedTransactions,Boolean newImport)
        {
            this.tableAttributes = impoertedTransactions;
            newImport = true;
        }
        public Boolean getNewImport()
        {
            return newImport;
        }
        public String getAccounNumber()
        {
            return accountNumber;
        }
        public List<Transaction> getTableAttributes()
        {
            return tableAttributes;
        }
        public User getCurrentUser()
        {
            return currentUser;
        }
        public void setAccountNumber(string _accountNumber)
        {
            accountNumber = _accountNumber;
        }
        public ButtonCommands ImportPushed
        {
            get
            {
                btnCommand = new ButtonCommands(ImportButton.Content.ToString(), this);
                return btnCommand;
            }
        }
        public ButtonCommands TablePushed
        {
            get
            {
                btnCommand = new ButtonCommands(TableButton.Content.ToString(), this);
                return btnCommand;
            }
        }
        public ButtonCommands ExitPushed
        {
            get
            {
                btnCommand = new ButtonCommands(ExitButton.Content.ToString(), this);
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
            SavedTransactions.getInstance();
        }
        public void getTransactions(string bankName,string folderAddress)
        {
            new ImportReadIn(bankName, folderAddress,this);
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
                mainWindow.MainFrame.Content = ImportMainPage.getInstance(mainWindow);
                mainWindow.tableMenuTop.Visibility = System.Windows.Visibility.Hidden;
                mainWindow.importMenuTop.Visibility = System.Windows.Visibility.Visible;
                mainWindow.importDock.Background = new SolidColorBrush(Color.FromRgb(198, 61, 15));
                mainWindow.tableDock.Background = new SolidColorBrush(Color.FromRgb(217, 133, 59));
            }
           else if(buttonContent.Equals("Database"))
           {
                mainWindow.MainFrame.Content=TransactionMain.getInstance(mainWindow, mainWindow.getTableAttributes(), mainWindow.getAccounNumber());
                mainWindow.importMenuTop.Visibility = System.Windows.Visibility.Hidden;
                mainWindow.tableMenuTop.Visibility = System.Windows.Visibility.Visible;
                mainWindow.tableDock.Background = new SolidColorBrush(Color.FromRgb(198, 61, 15));
                mainWindow.importDock.Background = new SolidColorBrush(Color.FromRgb(217, 133, 59));
           }
           else if(buttonContent.Equals("Exit"))
           {
                mainWindow.Close();
           }
        }
    }
}
