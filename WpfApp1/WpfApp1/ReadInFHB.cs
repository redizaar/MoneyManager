using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    class ReadInFHB
    {
        private List<Transaction> transactions;
        private ImportReadIn bankHanlder = null;
        public ReadInFHB(ImportReadIn importReadin, Workbook workbook, Worksheet worksheet)
        {
            worksheet = workbook.Worksheets[1];
            this.bankHanlder = importReadin;
            transactions = new List<Transaction>();

            int i = 20;
            string transactionDate = "";
            string osszegString = "";
            string egyenlegString = "";
            string accountNumberExtra = worksheet.Cells[8, 2].Value.ToString();
            string accountNumber = accountNumberExtra.Substring(0, 25); //substraction the HUF word
            int osszeg=0;
            int currentEgyenleg = 0;
            while ((worksheet.Cells[i, 1].Value != null) || (worksheet.Cells[i+1,1].Value!=null))//interesing FHB file..
            {
                if (worksheet.Cells[i, 1].Value != null)
                {
                    transactionDate = worksheet.Cells[i, 1].Value.ToString();
                    if (worksheet.Cells[i, 9].Value != null) //cost
                    {
                        osszegString = worksheet.Cells[i, 9].Value.ToString();
                        osszeg = int.Parse(osszegString);
                    }
                    else if (worksheet.Cells[i, 11].Value != null)//income
                    {
                        osszegString = worksheet.Cells[i, 11].Value.ToString();
                        osszeg = int.Parse(osszegString)*(-1);
                    }
                    if (worksheet.Cells[i, 13].Value == null)//in case if the Egyenleg cell is null in the first transaction (interesting FHB file)
                    {
                        int tempIndex = i + 1; //don't scan the current cell because we already know it's null
                        while (worksheet.Cells[tempIndex, 13].Value == null)
                        {
                            tempIndex++;
                        }
                        string oldEgyenlegString = "";
                        oldEgyenlegString = worksheet.Cells[tempIndex, 13].Value.ToString();
                        int oldEgyenlegInt = int.Parse(oldEgyenlegString);
                        //adding or substracting other transactions -- to get the real Egyenleg
                        while (tempIndex != i - 1)
                        {
                            string tempOsszegString = "";
                            int tempOsszegInt = 0;
                            if (worksheet.Cells[tempIndex, 9].Value != null)//cost
                            {
                                tempOsszegString = worksheet.Cells[tempIndex, 9].Value.ToString();
                                tempOsszegInt = int.Parse(tempOsszegString);
                            }
                            else if (worksheet.Cells[tempIndex, 11].Value != null)//income
                            {
                                tempOsszegString = worksheet.Cells[tempIndex, 11].Value.ToString();
                                tempOsszegInt = int.Parse(tempOsszegString) * (-1);
                            }
                            oldEgyenlegInt += tempOsszegInt;
                            tempIndex--;//going back up
                        }
                        currentEgyenleg = oldEgyenlegInt;
                    }
                    else
                    {
                        if (worksheet.Cells[i, 13].value != null)
                        {
                            egyenlegString = worksheet.Cells[i, 13].Value.ToString();
                            currentEgyenleg = int.Parse(egyenlegString);
                        }
                        else
                        {
                            int tempEgyenleg = 0;
                            if (worksheet.Cells[i, 9].Value != null)
                            {
                                egyenlegString = worksheet.Cells[i, 9].Value.ToString();
                                tempEgyenleg = int.Parse(egyenlegString) * (-1);
                                currentEgyenleg += tempEgyenleg;
                            }
                            else if (worksheet.Cells[i, 11].Value != null)
                            {
                                egyenlegString = worksheet.Cells[i, 11].Value.ToString();
                                tempEgyenleg = int.Parse(egyenlegString);
                                currentEgyenleg += tempEgyenleg;
                            }
                        }
                    }
                    Console.WriteLine(currentEgyenleg);
                    transactions.Add(new Transaction(currentEgyenleg, transactionDate, osszeg, currentEgyenleg += osszeg,accountNumber));
                }
                i++;
            }
            bankHanlder.addTransactions(transactions);
        }
        public List<Transaction> getTransactions()
        {
            return transactions;
        }
    }
}
