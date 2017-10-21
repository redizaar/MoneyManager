using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    class ExportTransactions
    {
        Workbook WriteWorkbook;
        Worksheet WriteWorksheet;
        _Application excel = new _Excel.Application();
        public ExportTransactions(List<Transaction> transactions)
        {

            List<Transaction> neededTransactions=newTransactions(transactions);
            WriteWorkbook = excel.Workbooks.Open(@"C:\Users\Tocki\Desktop\Kimutatas.xlsx");
            WriteWorksheet = WriteWorkbook.Worksheets[1];
            if (neededTransactions != null)
            {
                string todaysDate = DateTime.Now.ToString("yyyy-MM-dd"); ;
                int row_number = 1;
                while (WriteWorksheet.Cells[row_number, 1].Value != null)
                {
                    row_number++; // get the current last row
                }
                foreach (var transctn in neededTransactions)
                {

                    WriteWorksheet.Cells[row_number, 1].Value = todaysDate;
                    WriteWorksheet.Cells[row_number, 2].Value = transctn.getTransactionDate();
                    WriteWorksheet.Cells[row_number, 3].Value = transctn.getBalance_rn();
                    WriteWorksheet.Cells[row_number, 7].Value = transctn.getTransactionPrice();
                    if (transctn.getTransactionPrice() < 0)
                    {
                        WriteWorksheet.Cells[row_number, 9].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 11].Value = transctn.getBalance_rn() - transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 15].Value = "havi";
                    }
                    else
                    {
                        WriteWorksheet.Cells[row_number, 8].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 10].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 11].Value = transctn.getBalance_rn() - transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 15].Value = "havi";
                    }
                    WriteWorksheet.Cells[row_number, 14].Value = transctn.getAccountNumber();
                    row_number++;
                    Range line = (Range)WriteWorksheet.Rows[row_number];
                    line.Insert();
                    Console.WriteLine(row_number + " sor beszurva");
                }
                WriteWorkbook.SaveAs(@"C:\Users\Tocki\Desktop\Kimutatas.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                WriteWorkbook.Close();
            }
            else
            {
                return;
            }
        }
        private List<Transaction> newTransactions(List<Transaction> importedTransactions)
        {
            List<Transaction> savedTransactions = SavedTransactions.getSavedTransactions();
            List<Transaction> neededTransactions=new List<Transaction>();
            string accountNumber = importedTransactions[0].getAccountNumber();//account number is the same for all
            if (savedTransactions.Count != 0)
            {
                List<Transaction> tempTransactions = new List<Transaction>();
                foreach (var saved in savedTransactions)
                {
                   if(saved.getAccountNumber().Equals(accountNumber))
                    {
                        tempTransactions.Add(saved);
                    }
                }
                foreach (var saved in tempTransactions)
                {
                    foreach (var imported in importedTransactions)
                    {
                        if(!(saved.getTransactionDate().Equals(imported.getTransactionDate()) && 
                                saved.getTransactionPrice().Equals(imported.getTransactionPrice()) && 
                                saved.getBalance_rn()==imported.getBalance_rn()))
                        {
                            neededTransactions.Add(imported);
                        }
                        
                    }
                }
                return neededTransactions;
            }
            else
            {
                return importedTransactions;
            }
        }
        ~ExportTransactions()
        {
            excel.Application.Quit();
            excel.Quit();
        }
    }
}
