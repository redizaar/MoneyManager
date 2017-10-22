using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    class SavedTransactions
    {
        _Application excel = new _Excel.Application();
        Workbook ReadWorkbook;
        Worksheet ReadWorksheet;
        static List<Transaction> savedTransactions;

        public SavedTransactions()
        {
            savedTransactions = new List<Transaction>();
            ReadWorkbook = excel.Workbooks.Open(@"C:\Users\Tocki\Desktop\Kimutatas.xlsx");
            ReadWorksheet = ReadWorkbook.Worksheets[1];
            int i = 2;
            while(ReadWorksheet.Cells[i,1].Value!=null)
            {
                string writeoutDate="";
                string transactionDate="";
                string balanceString = "";
                int balance = 0;
                string transactionPriceString = "";
                int transactionPrice = 0;
                string accountNumber = "";

                writeoutDate = ReadWorksheet.Cells[i, 1].Value.ToString();
                transactionDate = ReadWorksheet.Cells[i, 2].Value.ToString();
                balanceString = ReadWorksheet.Cells[i, 3].Value.ToString();
                balance = int.Parse(balanceString);
                if(ReadWorksheet.Cells[i,7].Value!=null)
                {
                    transactionPriceString = ReadWorksheet.Cells[i, 7].Value.ToString();
                    transactionPrice = int.Parse(transactionPriceString);
                }
                else if(ReadWorksheet.Cells[i,9].Value!=null)
                {
                    transactionPriceString = ReadWorksheet.Cells[i, 9].Value.ToString();
                    transactionPrice = int.Parse(transactionPriceString);
                }
                accountNumber = ReadWorksheet.Cells[i, 14].Value.ToString();

                savedTransactions.Add(new Transaction(writeoutDate,transactionDate,balance,transactionPrice,accountNumber));
                i++;
            }
            Console.WriteLine(savedTransactions.Count);
        }
        public static List<Transaction> getSavedTransactions()
        {
            if (savedTransactions != null)
            {
                return savedTransactions;
            }
            else
            {
                return null;
            }
        }
        ~SavedTransactions()
        {
            excel.Application.Quit();
            excel.Quit();
        }
    }
}
