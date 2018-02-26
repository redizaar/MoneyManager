using System;
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

namespace WpfApp1
{
    public partial class TransactionMain : Page
    {
        private List<Transaction> tableAttribues;
        public List<string> categoryName { get; set; }
        private static TransactionMain instance;

        private TransactionMain(MainWindow mainWindow,List<Transaction> tableAttribues, string accountNumber)
        {
            this.DataContext = this;
            InitializeComponent();
            if (TransactionTableXAML != null)
            {
                TransactionTableXAML.Items.Clear();
            }
            if (tableAttribues != null && tableAttribues!=this.tableAttribues)
            {
                foreach (var transaction in tableAttribues)
                {
                    if (transaction.getWriteDate() != null && transaction.getWriteDate().Length>=12)
                    {
                        transaction.setWriteDate(transaction.getWriteDate().Substring(0, 12));
                    }
                    else
                    {
                        transaction.setWriteDate(DateTime.Now.ToString("yyyy/MM/dd"));
                    }
                }
            }
            else if(tableAttribues==null)
            {
                foreach (var transaction in SavedTransactions.getSavedTransactions())
                {
                    if (transaction.getWriteDate() != null && transaction.getWriteDate().Length != 12)
                    {
                        transaction.setWriteDate(transaction.getWriteDate().Substring(0, 12));
                    }
                    else
                    {
                        transaction.setWriteDate(DateTime.Now.ToString("yyyy/MM/dd"));
                    }
                }
            }
            if ((this.tableAttribues != tableAttribues) && (tableAttribues != null))
            {
                this.tableAttribues = tableAttribues;
                if (accountNumber.Equals(""))
                {
                    addAtribuesToTable(); //we have imported and saved files in this case
                                         //the accountNumber is already matching
                }
                else
                {
                    addAtribuesToTable(accountNumber);
                }
            }
            else
            {
                addAtribuesToTable(accountNumber);
            }
        }
        private void addAtribuesToTable(String accountNumber)
        {
            if (accountNumber.Equals("empty"))//only imported files
            {
                foreach (var attribute in tableAttribues)
                {
                    TransactionTableXAML.Items.Add(attribute);
                }
            }
            else
            {
                foreach (var attribute in SavedTransactions.getSavedTransactions())
                {
                    if (attribute.getAccountNumber().Equals(accountNumber))//only saved files
                    {
                        TransactionTableXAML.Items.Add(attribute);
                    }
                }
            }
        }
        private void addAtribuesToTable()
        {
            foreach (var attribute in tableAttribues)
            {
                TransactionTableXAML.Items.Add(attribute);
            }
        }
        public static TransactionMain getInstance(MainWindow mainWindow,List<Transaction> attributes,string accountnumber)
        {
            if(instance==null)
            {
                instance = new TransactionMain(mainWindow,attributes,accountnumber);
            }
            return instance;
        }
    }
}
