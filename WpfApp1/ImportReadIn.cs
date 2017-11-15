using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace WpfApp1
{
    class ImportReadIn
    {
        private string path = "";
        private string bankName = "";
        private MainWindow mainWindow;
        List<Transaction> transactions;

        _Application excel = new _Excel.Application();
        Workbook ReadWorkbook;
        Worksheet ReadWorksheet;
        public ImportReadIn(string bankName, string path,MainWindow mainWindow)
        {
            this.path = path;
            this.bankName = bankName;
            this.mainWindow = mainWindow;
            //this.importMainPage = importMainPage;

            if (path != "FolderAdress")//a path wasn't choosen
            {
                ReadWorkbook = excel.Workbooks.Open(path);
                ReadWorksheet = ReadWorkbook.Worksheets[1];
                if (bankName.Equals("OTP"))
                {
                    new ReadInOTP(this, ReadWorkbook, ReadWorksheet);
                }
                else if (bankName.Equals("FHB"))
                {
                    new ReadInFHB(this, ReadWorkbook, ReadWorksheet);
                }
                else if (bankName.Equals("K&H"))
                {
                    new ReadInKandH(this, ReadWorkbook, ReadWorksheet);
                }
                else if (bankName.Equals("All"))
                {
                    TemplateReadIn TemplateBank = new TemplateReadIn(this, ReadWorkbook, ReadWorksheet);
                    //so far we got the Starting Row(of the transactions),Number of Columns, account number
                    TemplateBank.readOutTransactionColumns(TemplateBank.getStartingRow(), TemplateBank.getNumberOfColumns());
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
            new ExportTransactions(transactions,mainWindow);
        }
    }
}
