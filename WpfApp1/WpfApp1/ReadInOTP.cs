using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    class ReadInOTP
    {
        private List<Transaction> transactions;
        private ImportReadIn bankHanlder=null;
        public ReadInOTP(ImportReadIn importReadin,Workbook workbook, Worksheet worksheet)
        {
            worksheet = workbook.Worksheets[1];
            this.bankHanlder = importReadin;
            transactions = new List<Transaction>();
            int i = 1;
            int egyenleg_rn=0;
            string todaysDate = DateTime.Now.ToString("yyyy-MM-dd"); ;
            string transactionDate = "";
            int osszeg = 0;
            int new_egyenleg = 0;
            bool need_values = true;
            string osszeg_string = "";
            string new_balance_string = "";
            i = 15;
            while (worksheet.Cells[i, 1].Value != null)
            {
                //egyenleg += osszeg;
                while (need_values)
                {
                    int j = 3;
                    transactionDate = worksheet.Cells[i, j].Value.ToString();
                    j = j + 2;
                    osszeg_string = worksheet.Cells[i, j].Value.ToString();
                    j++;
                    if(i>15)
                    {
                        egyenleg_rn = new_egyenleg;
                    }
                    new_balance_string = worksheet.Cells[i, j].Value.ToString();
                    if (i==15)
                    {
                        egyenleg_rn = int.Parse(new_balance_string);
                    }
                    need_values = false;
                    osszeg = int.Parse(osszeg_string);
                    new_egyenleg = int.Parse(new_balance_string);
                }
                i++;
                need_values = true;
                transactions.Add(new Transaction(egyenleg_rn, transactionDate, osszeg, new_egyenleg));
            }
            bankHanlder.addTransactions(transactions);
        }
        public List<Transaction> getTransactions()
        {
            return transactions;
        }
    }
}
