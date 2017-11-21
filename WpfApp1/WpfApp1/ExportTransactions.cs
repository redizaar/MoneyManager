﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using System.IO;

namespace WpfApp1
{
    class ExportTransactions
    {
        Workbook WriteWorkbook;
        Worksheet WriteWorksheet;
        _Application excel = new _Excel.Application();
        private MainWindow mainWindow;
        private string importerAccountNumber;
        public ExportTransactions(List<Transaction> transactions,MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
                                                    //BUT FIRST - check if the transaction is already exported or not 
            List<Transaction> neededTransactions = newTransactions(transactions);
            SavedTransactions.addToSavedTransactions(neededTransactions);//adding the freshyl imported transactions to the saved 
            WriteWorkbook = excel.Workbooks.Open(@"C:\Users\Tocki\Desktop\Kimutatas.xlsx");
            WriteWorksheet = WriteWorkbook.Worksheets[1];
            if (neededTransactions != null)
            {
                string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
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
                        WriteWorksheet.Cells[row_number, 11].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 15].Value = "havi";
                    }
                    else
                    {
                        WriteWorksheet.Cells[row_number, 8].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 10].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 11].Value = transctn.getTransactionPrice();
                        WriteWorksheet.Cells[row_number, 15].Value = "havi";
                    }
                    WriteWorksheet.Cells[row_number, 14].Value = transctn.getTransactionDescription();
                    WriteWorksheet.Cells[row_number, 16].Value = transctn.getAccountNumber();
                    row_number++;
                    Range line = (Range)WriteWorksheet.Rows[row_number];
                    line.Insert();
                }
                try
                {
                    excel.DisplayAlerts = false;
                    WriteWorkbook.SaveAs(@"C:\Users\Tocki\Desktop\Kimutatas.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                        Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges,
                        Type.Missing, Type.Missing);
                    ImportMainPage.getInstance(mainWindow).getUserStatistics(importerAccountNumber);
                }
                catch(Exception e)
                {

                }
                excel.Application.Quit();
                excel.Quit();
            }
            else
            {
                return;
            }
        }
        private List<Transaction> newTransactions(List<Transaction> importedTransactions) //check if the transaction is already exported or not
        {
            List<Transaction> savedTransactions = SavedTransactions.getSavedTransactions();
            List<Transaction> neededTransactions=new List<Transaction>();
            importerAccountNumber = importedTransactions[0].getAccountNumber();//account number is the same for all
            mainWindow.setAccountNumber(importerAccountNumber);
            if (savedTransactions.Count != 0)//if the export file is not empty we scan it
            {
                List<Transaction> tempTransactions = new List<Transaction>();
                foreach (var saved in savedTransactions)
                {
                   //egy külön listába tesszük azokat az elemeket a már elmentet tranzakciókból ahol a bankszámlaszám
                   //megegyezik az importálandó bankszámlaszámmal
                   if(saved.getAccountNumber().Equals(importerAccountNumber))
                    {
                        tempTransactions.Add(saved);
                    }
                }
                if (tempTransactions.Count != 0)//ha van olyan már elmentett tranzakció aminek az  a bankszámlaszáma mint amit importálni akarunk
                {
                    int explicitImported=0;
                    //StreamWriter logFile =new System.IO.StreamWriter("C:\\Users\\Tocki\\Desktop\\transactionsLog.txt", append:true);
                    foreach (var imported in importedTransactions)
                    {
                        bool redundant = false;
                        foreach (var saved in tempTransactions)
                        {
                            if (saved.getTransactionDate().Equals(imported.getTransactionDate()) &&
                                    saved.getTransactionPrice().Equals(imported.getTransactionPrice()) &&
                                    saved.getBalance_rn() == imported.getBalance_rn())
                            {
                                redundant = true;
                                if (ImportMainPage.getInstance(mainWindow).alwaysAsk.Equals(true))
                                {
                                    if (MessageBox.Show("This transaction is most likely to be in your Databse already\n Transaction date: " + imported.getTransactionDate() + "\nTransaction price: " + imported.getTransactionPrice()
                                        + "\nPossibly imported on: " + saved.getWriteDate().Substring(0,12)+"\nWould you like to import it?",
                                     "Redundant Transactions",
                                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                                    {
                                        neededTransactions.Add(imported);
                                        explicitImported++;
                                        //logFile.WriteLine("AccountNumber: " + imported.getAccountNumber() +
                                         //   "\n ImportDate: " + imported.getTransactionDate() +
                                          //  "\n TransactionPrice: " 
                                          //  + imported.getTransactionPrice()+"\n*");
                                    }
                                }
                                break;
                            }
                        }
                        if (redundant == false)
                        {
                            neededTransactions.Add(imported);;
                        }
                    }
                    if(neededTransactions.Count==0)
                    {
                        mainWindow.setTableAttribues(savedTransactions, importerAccountNumber);
                        //only pass the saved transactions because we didn't add new
                        //and the accountNumber so we can select it by user
                    }
                    else
                    {
                        //we pass both the saved and the new transcations
                        List<Transaction> savedAndImported=new List<Transaction>();

                        //tempTrancations containts saved Transactions where the accountnumber matches with the imported Transactions
                        foreach (var attribue in tempTransactions)
                        {
                            savedAndImported.Add(attribue);
                        }
                        foreach (var attribue in neededTransactions)
                        {
                            savedAndImported.Add(attribue);
                        }
                        mainWindow.setTableAttribues(savedAndImported,true);
                    }
                    if (MessageBox.Show("You have imported "+neededTransactions.Count+" new transaction(s)!\n" +
                        "("+(tempTransactions.Count-explicitImported)+" was already imported)", "OK",
                         MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                    {
                        return neededTransactions;
                    }
                    return neededTransactions;
                }
                else //nincs olyan elmentett tranzakció aminek az lenne a bankszámlaszáma mint amit importálni akarunk
                {
                    mainWindow.setTableAttribues(importedTransactions,"empty");
                    if (MessageBox.Show("You have imported " + importedTransactions.Count + " new transaction(s)!\n", "OK",
                         MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                    {
                        return importedTransactions;
                    }
                    return importedTransactions;
                }
            }
            else // még nincs elmentett tranzakció
            {
                mainWindow.setTableAttribues(importedTransactions,"empty");
                if (MessageBox.Show("You have imported " + importedTransactions.Count + " new transaction(s)!\n", "OK",
                         MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                {
                    return importedTransactions;
                }
                return importedTransactions;
            }
        }
        public string geImporterAccountNumber()
        {
            return importerAccountNumber;
        }
        public void setimporterAccountNumber(string value)
        {
            importerAccountNumber = value;
        }
        ~ExportTransactions()
        {
            excel.Application.Quit();
            excel.Quit();
        }
    }
}
