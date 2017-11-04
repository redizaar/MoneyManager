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
        private string accountNumber;
        public TemplateReadIn(ImportReadIn importReadin, Workbook workbook, Worksheet worksheet)
        {
            worksheet = workbook.Worksheets[1];
            this.bankHanlder = importReadin;
            transactions = new List<Transaction>();
            this.TransactionSheet = worksheet;
            getNofTransactions();
        }
        private void getNofTransactions()
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
        public void getTransactionDate(int row, int maxColumn)
        {
            Regex dateRegex1 = new Regex(@"^20\d{2}.\d{2}.\d{2}");
            Regex dateRegex2 = new Regex(@"^20\d{2}-\d{2}-\d{2}");
            Regex dateRegex3 = new Regex(@"^20\d{2}.\s\d{2}.\s\d{2}");
            for (int column = 1; column < maxColumn; column++)
            {
                if (TransactionSheet.Cells[row+1, column].Value != null)
                {
                    string inputData = TransactionSheet.Cells[row+1, column].Value.ToString();
                    if (dateRegex1.IsMatch(inputData) || dateRegex2.IsMatch(inputData) || dateRegex3.IsMatch(inputData))
                    {
                        Console.WriteLine(inputData + " egy datum");
                    }
                }
            }
            Console.WriteLine(getAccountNumber());
        }
        public void getTransactionPrices(int row, int maxColumn)
        {
            int blank_counter = 0;
            for (int j = 1; j < maxColumn; j++)
            {

            }
            /*
            while (blank_counter>2)
            {
                for (int j = 1; j < maxColumn; j++)
                {

                }
                    if (TransactionSheet.Cells[row , j].Value != null)
                    {
                        blank_counter = 0;

                    }
                    else
                    {
                        blank_counter++;
                    }
                row++;
            }
            */
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
    }
}
