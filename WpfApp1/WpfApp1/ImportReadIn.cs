﻿using System;
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
        List<Transaction> transactions;

        _Application excel = new _Excel.Application();
        Workbook ReadWorkbook,WriteWorkbook;
        Worksheet ReadWorksheet,WriteWorksheet;
        public ImportReadIn(string bankName,string path)
        {
            this.path = path;
            this.bankName = bankName;
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
            new ExportTransactions(transactions);
        }
    }
}
