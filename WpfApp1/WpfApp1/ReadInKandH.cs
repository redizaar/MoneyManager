using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    class ReadInKandH
    {
        private List<Transaction> transactions;
        private ImportReadIn bankHanlder = null;
        public ReadInKandH(ImportReadIn importReadin, Workbook workbook, Worksheet worksheet)
        {
            worksheet = workbook.Worksheets[1];
            this.bankHanlder = importReadin;
            transactions = new List<Transaction>();
            string transactionDate = "";
            string osszegString = "";
            //string egyenlegString = "";
            int osszeg = 0;
            int currentEgyenleg = 0;
            string accountNumber = worksheet.Cells[2,4].Value.ToString();

            int tempIndex = 2;
            while (worksheet.Cells[tempIndex, 1].Value != null)
            {
                tempIndex++;
            }
            int i = tempIndex;

            while (i!=2)
            {
                transactionDate = worksheet.Cells[i, 1].Value.ToString();
                osszegString = worksheet.Cells[i, 8].Value.ToString();
                osszeg = int.Parse(osszegString);
                currentEgyenleg += osszeg;
                transactions.Add(new Transaction(currentEgyenleg, transactionDate, osszeg, "old read IN OTP", accountNumber));
                i--;
            }
            bankHanlder.addTransactions(transactions);
        }
    }
}
