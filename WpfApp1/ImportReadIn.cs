using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace WpfApp1
{
    public class ImportReadIn
    {
        private List<string> path;
        private string bankName = "";
        private string currentFileName;
        private MainWindow mainWindow;
        List<Transaction> transactions;

        _Application excel = new _Excel.Application();
        Workbook ReadWorkbook;
        Worksheet ReadWorksheet;
        public ImportReadIn(string bankName, List<string> path,MainWindow mainWindow,bool userSpecified)
        {
            this.path = path;
            this.bankName = bankName;
            this.mainWindow = mainWindow;
            if (path[0] != "FolderAdress")//a path wasn't choosen
            {
                for (int i = 0; i < path.Count; i++)
                {
                    string [] splittedFileName=path[i].Split('\\');
                    int lastSplitIndex = splittedFileName.Length-1;
                    currentFileName = splittedFileName[lastSplitIndex];
                    ReadWorkbook = excel.Workbooks.Open(path[i]);
                    ReadWorksheet = ReadWorkbook.Worksheets[1];
                    if (bankName.Equals("All"))
                    {
                        if (!userSpecified)
                        {
                            TemplateReadIn templateBank = new TemplateReadIn(this, ReadWorkbook, ReadWorksheet, mainWindow, false);
                            //so far we got the Starting Row(of the transactions),Number of Columns, account number
                            templateBank.readOutTransactionColumns(templateBank.getStartingRow(), templateBank.getNumberOfColumns());
                        }
                        else //userSpecified==true
                        {
                            TemplateReadIn templateBank = new TemplateReadIn(this, ReadWorkbook, ReadWorksheet, mainWindow, true);
                            string startingRow = SpecifiedImport.getInstance(null,mainWindow).transactionsRowTextBox.Text.ToString();
                            string dateColumn = SpecifiedImport.getInstance(null,mainWindow).dateColumnTextBox.Text.ToString();
                            string commentColumn = SpecifiedImport.getInstance(null,mainWindow).commentColumnTextBox.Text.ToString();
                            string accountNumberCB = SpecifiedImport.getInstance(null,mainWindow).accountNumberCB.SelectedItem.ToString();
                            string transactionPriceCB = SpecifiedImport.getInstance(null,mainWindow).priceColumnCB.SelectedItem.ToString();
                            string balanceCB = SpecifiedImport.getInstance(null,mainWindow).balanceColumnCB.SelectedItem.ToString();
                            string balanceComboBocString = SpecifiedImport.getInstance(null, mainWindow).balanceColumnTextBox.Text.ToString();
                            templateBank.readOutUserspecifiedTransactions(startingRow, dateColumn, commentColumn, accountNumberCB, transactionPriceCB, balanceCB, balanceComboBocString);
                        }
                    }
                }
                excel.Application.Quit();
                excel.Quit();
            }
        }
        ~ImportReadIn()
        {
            excel.Application.Quit();
            excel.Quit();
        }
        public void addTransactions(List<Transaction> newTransactions)
        {
            this.transactions = newTransactions;
            writeOutTransactions();
        }
        public void writeOutTransactions()
        {
            new ExportTransactions(transactions,mainWindow,currentFileName);
        }
    }
}
