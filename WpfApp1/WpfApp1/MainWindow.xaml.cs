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
        private string accountNumber="";
        public MainWindow()
        {
            DataContext = this;
            InitializeComponent();

            
            startUpReadIn();
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
            new SavedTransactions();
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
                mainWindow.MainFrame.Content = new ImportMainPage(mainWindow);
            }
           else if(buttonContent.Equals("Table"))
            {
                mainWindow.MainFrame.Content=new TransactionMain(mainWindow, mainWindow.getTableAttributes(), mainWindow.getAccounNumber());
            }
           else if(buttonContent.Equals("Exit"))
            {
                mainWindow.Close();
            }
        }
    }
}
