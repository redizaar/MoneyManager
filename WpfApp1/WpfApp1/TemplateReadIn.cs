using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace WpfApp1
{
    class TemplateReadIn
    {
        private Worksheet TransactionSheet;

        private List<Transaction> transactions;
        private ImportReadIn bankHanlder = null;
        private int startingRow;
        private int nofColumns;
        private int pastTransactionPrice;//in case of missing Balance column..
        private bool isFirstTransaction;//in case of missing Balance column..
        private string accountNumber;
        private bool multipleColumn;

        public TemplateReadIn(ImportReadIn importReadin, Workbook workbook, Worksheet worksheet)
        {
            worksheet = workbook.Worksheets[1];
            this.bankHanlder = importReadin;
            transactions = new List<Transaction>();
            this.TransactionSheet = worksheet;
            this.multipleColumn = false;
            this.isFirstTransaction = false;

            getTransactionRows();
        }
        private void getTransactionRows()
        {
            this.accountNumber = "";
            Regex accoutNumberRegex1 = new Regex(@"^Számlaszám$");
            Regex accountNumberRegex2 = new Regex(@"^Könyvelési számla$"); 
            int blank_row=0;
            int blank_cells = 0;
            int i = 1;

            int maxColumns=1;
            int transactionsStartRow = 1;
            while (blank_row<4)
            {
                int column = 1;
                if(TransactionSheet.Cells[i,column].Value!=null)
                {
                    if (this.accountNumber.Equals(""))
                    {
                        if (column == 1)
                        {
                            string cellValue = TransactionSheet.Cells[i, column].Value.ToString();
                            if (accoutNumberRegex1.IsMatch(cellValue))
                            {
                                string accountNumberValue = TransactionSheet.Cells[i, column + 1].Value.ToString();//the cell next to it
                                setAccountNumber(accountNumberValue);
                            }
                        }
                    }
                    blank_cells=0;
                    while(blank_cells<3)
                    {
                        if(TransactionSheet.Cells[i, column].Value != null)
                        {
                            column++;
                            blank_cells = 0;
                        }
                        else
                        {
                            column++;
                            blank_cells++;
                        }
                    }
                    blank_row=0;
                }
                else
                {
                    blank_row++;
                }
                if(column>maxColumns)
                {
                    maxColumns = column;
                    transactionsStartRow = i;
                    if(this.accountNumber.Equals(""))
                    {
                        for(int j=1;j<column;j++)
                        {
                            if(TransactionSheet.Cells[i, j].Value!=null)
                            {
                                string cellValue = TransactionSheet.Cells[i, j].Value.ToString();
                                if(accoutNumberRegex1.IsMatch(cellValue) || accountNumberRegex2.IsMatch(cellValue))
                                {
                                    string accountNumberValue = TransactionSheet.Cells[i+1, j].Value.ToString();//the cell below it
                                    setAccountNumber(accountNumberValue);
                                }
                            }
                        }
                    }
                }
                i++;
            }
            setStartingRow(transactionsStartRow);
            setNofColumns(maxColumns-blank_cells);
        }

        public void readOutTransactionColumns(int row, int maxColumn)
        {
            int dateColumn=getDateColumn(row,maxColumn);
            string pricecolumnType = isMultiplePriceColumn(row,maxColumn);
            int singlepriceColumn = -1;
            try
            {
                singlepriceColumn=int.Parse(pricecolumnType);
            }
            catch(Exception e)
            {

            }
            if(singlepriceColumn==-1)
            {
                this.multipleColumn = true;
            }
            int balaceColumn=getAccountBalanceColumn(row,maxColumn);
            if(balaceColumn==-1)
            {

            }
            readOutTransactions(row,maxColumn,dateColumn,singlepriceColumn,balaceColumn);
        }

        private void readOutTransactions(int row, int maxColumn,int dateColumn, int singlepriceColumn, int balaceColumn)
        {
            if(row==1)
            {
                row++;
            }
            else
            {
                Regex dateRegex1 = new Regex(@"^20\d{2}.\d{2}.\d{2}");
                Regex dateRegex2 = new Regex(@"^20\d{2}-\d{2}-\d{2}");
                Regex dateRegex3 = new Regex(@"^20\d{2}.\s\d{2}.\s\d{2}");
                bool titleRow = true;
                for (int j = 1; j < maxColumn; j++)
                {
                    if (TransactionSheet.Cells[row, j].Value != null)
                    { 
                        string inputdata = TransactionSheet.Cells[row, j].Value.ToString();
                        if ((dateRegex1.IsMatch(inputdata) || dateRegex2.IsMatch(inputdata) || dateRegex3.IsMatch(inputdata)))
                        {
                            titleRow = false;
                            break;
                        }
                    }
                }
                if(titleRow)
                {
                    row++;
                }
            }
            if(singlepriceColumn!=-1)//single column
            {
                int blank_counter = 0;
                List<Transaction> transaction=new List<Transaction>();
                while(blank_counter<2)
                {
                    if (balaceColumn != -1)//have balance column
                    {
                        if (TransactionSheet.Cells[row, dateColumn].Value != null && TransactionSheet.Cells[row, singlepriceColumn].Value != null)
                        {
                            blank_counter = 0;

                            string transactionDate = TransactionSheet.Cells[row, dateColumn].Value.ToString();
                            string accountNumber = getAccountNumber();
                            string transactionPriceString = TransactionSheet.Cells[row, singlepriceColumn].Value.ToString();
                            string transactionBalanceString = TransactionSheet.Cells[row, balaceColumn].Value.ToString();

                            int transactionPrice = 0;
                            int transactionBalance = 0;
                            try
                            {
                                transactionPrice = int.Parse(transactionPriceString);
                                transactionBalance = int.Parse(transactionBalanceString);
                            }
                            catch (Exception e)
                            {

                            }
                            transaction.Add(new Transaction(transactionBalance, transactionDate, transactionPrice, transactionBalance + transactionPrice, accountNumber));
                        }
                        else
                        {
                            blank_counter++;
                        }
                    }
                    else//don't have balance column
                    {
                        if (TransactionSheet.Cells[row, dateColumn].Value != null && TransactionSheet.Cells[row, singlepriceColumn].Value != null)
                        {
                            blank_counter = 0;

                            string transactionDate = TransactionSheet.Cells[row, dateColumn].Value.ToString();
                            string accountNumber = getAccountNumber();
                            string transactionPriceString = TransactionSheet.Cells[row, singlepriceColumn].Value.ToString();
                            int transactionPrice = 0;
                            try
                            {
                                transactionPrice = int.Parse(transactionPriceString);
                            }
                            catch (Exception e)
                            {

                            }
                            if (this.getIsFirstTransaction())//we pretend that the balance is 0
                            {
                                transaction.Add(new Transaction(transactionPrice, transactionDate, transactionPrice, 0, accountNumber));
                                this.setPastTransactionPrice(transactionPrice);
                                this.setIsFirstTransaction(false);
                            }
                            else
                            {
                                transaction.Add(new Transaction(this.getPastTransactionPrice() + transactionPrice, transactionDate, transactionPrice, this.getPastTransactionPrice(), accountNumber));
                                this.setPastTransactionPrice(this.getPastTransactionPrice()+transactionPrice);
                            }
                        }
                        else
                        {
                            blank_counter++;
                        }
                    }
                    row++;
                }
                foreach(var valami in transaction)
                {
                    Console.WriteLine("datum: "+valami.getTransactionDate() + " szamlaszam: " + valami.getAccountNumber() + " osszeg: " + valami.getTransactionPrice() + " egyenleg: " + valami.getBalance_rn());
                }
            }
            else//multiple price columns
            {

            }
        }

        private int getAccountBalanceColumn(int row, int maxColumn)
        {
            Regex balanceRegex1 = new Regex(@"^Egyenleg$");
            Regex balanceRegex2 = new Regex(@"könyvelt egyenleg$");

            if (row != 1)
            {
                for (int i = row - 1; i <= row + 2; i++)
                {
                    for (int j = 1; j < maxColumn; j++)
                    {
                        if (TransactionSheet.Cells[i, j].Value != null)
                        {
                            string inputData = TransactionSheet.Cells[i, j].Value.ToString();
                            if (balanceRegex1.IsMatch(inputData) || balanceRegex2.IsMatch(inputData))
                            {
                                return j;
                            }
                        }
                    }
                }
            }
            else//colum titles first row
            {
                for (int j = 1; j < maxColumn; j++)
                {
                    if (TransactionSheet.Cells[row, j].Value != null)
                    {
                        string inputData = TransactionSheet.Cells[row, j].Value.ToString();
                        if (balanceRegex1.IsMatch(inputData) || balanceRegex2.IsMatch(inputData))
                        {
                            return j;
                        }
                    }
                }
            }
            return -1;
        }

        private string isMultiplePriceColumn(int row, int maxColumn)
        {
            Regex priceRegex1 = new Regex(@"Összeg");
            Regex priceRegex2 = new Regex(@"összeg");
            Regex priceRegex3 = new Regex(@"Terhelés$");
            Regex priceRegex4 = new Regex(@"Jóváírás$");
            if (row != 1)
            {
                for (int i = row-1; i <= row+2; i++)
                {
                    for (int j = 1; j < maxColumn; j++)
                    {
                        if (TransactionSheet.Cells[i, j].Value != null)
                        {
                            string inputData = TransactionSheet.Cells[i, j].Value.ToString();
                            if (priceRegex1.IsMatch(inputData) || priceRegex2.IsMatch(inputData))
                            {
                                return j.ToString();
                            }
                            else if (priceRegex3.IsMatch(inputData) || priceRegex4.IsMatch(inputData))
                            {
                                return "multiple";
                            }
                         }
                     }
                }
            }
            else//colum titles first row
            {
                for (int j = 1; j < maxColumn; j++)
                {
                    if (TransactionSheet.Cells[row, j].Value != null)
                    {
                        string inputData = TransactionSheet.Cells[row, j].Value.ToString();
                        if (priceRegex1.IsMatch(inputData) || priceRegex2.IsMatch(inputData))
                        {
                            return j.ToString();
                        }
                        else if (priceRegex3.IsMatch(inputData) || priceRegex4.IsMatch(inputData))
                        {
                            return "multiple";
                        }
                    }
                }
            }
            return null;
        }

        private int getDateColumn(int row, int maxColumn)
        {
            Regex dateRegex1 = new Regex(@"^20\d{2}.\d{2}.\d{2}");
            Regex dateRegex2 = new Regex(@"^20\d{2}-\d{2}-\d{2}");
            Regex dateRegex3 = new Regex(@"^20\d{2}.\s\d{2}.\s\d{2}");
            if (row != 1)
            {
                for (int i = row; i < row + 3; i++)
                {
                    for (int column = 1; column < maxColumn; column++)
                    {
                        if (TransactionSheet.Cells[i, column].Value != null)
                        {
                            string inputData = TransactionSheet.Cells[i, column].Value.ToString();
                            if (dateRegex1.IsMatch(inputData) || dateRegex2.IsMatch(inputData) || dateRegex3.IsMatch(inputData))
                            {
                                return column;
                            }
                        }
                    }
                }
            }
            else//colum titles first row
            {
                for (int i = row + 1; i < row + 3; i++)
                {
                    for (int column = 1; column < maxColumn; column++)
                    {
                        if (TransactionSheet.Cells[i, column].Value != null)
                        {
                            string inputData = TransactionSheet.Cells[i, column].Value.ToString();
                            if (dateRegex1.IsMatch(inputData) || dateRegex2.IsMatch(inputData) || dateRegex3.IsMatch(inputData))
                            {
                                return column;
                            }
                        }
                    }
                }
            }
            return -1;
        }
        
        private void setStartingRow(int value)
        {
            startingRow = value;
        }
        private void setNofColumns(int value)
        {
            nofColumns = value;
        }
        private void setAccountNumber(string value)
        {
            accountNumber = value;
        }
        private void setPastTransactionPrice(int value)
        {
            pastTransactionPrice = value;
        }
        private void setIsFirstTransaction(bool value)
        {
            isFirstTransaction = value;
        }

        public int getStartingRow()
        {
            return startingRow;
        }
        public int getNumberOfColumns()
        {
            return nofColumns;
        }
        public string getAccountNumber()
        {
            return accountNumber;
        }
        public int getPastTransactionPrice()
        {
            return pastTransactionPrice;
        }
        public bool getIsFirstTransaction()
        {
            return isFirstTransaction;
        }
    }
}
