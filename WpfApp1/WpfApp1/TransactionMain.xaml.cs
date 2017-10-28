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
        public TransactionMain(List<Transaction> tableAttribues, String accountNumber)
        {
            InitializeComponent();
            if (this.tableAttribues != tableAttribues && tableAttribues != null)
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
        }
        private void addAtribuesToTable(String accountNumber)
        {
            if (accountNumber.Equals("empty"))//only imported files
            {
                foreach (var attribute in tableAttribues)
                {
                    if (attribute.getWriteDate().Equals(null))
                    {
                        attribute.setWriteDate(DateTime.Now.ToString("M/d/yyyy"));
                    }
                    TransactionTableXAML.Items.Add(attribute);
                }
            }
            else
            {
                foreach (var attribute in tableAttribues)
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
                if (attribute.getWriteDate().Equals(null))
                {
                    attribute.setWriteDate(DateTime.Now.ToString("M/d/yyyy"));
                }
                TransactionTableXAML.Items.Add(attribute);
            }
        }
    }
}
