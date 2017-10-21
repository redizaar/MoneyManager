using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    class Transaction
    {
        private int balance_rn;
        private int balance_past;
        private string transactionDate;
        private int transactionPrice;
        private string accountNumber;

        private string writeDate;
        public Transaction(int balance_rn,string date,int price,int balance_past,string accountNumber)
        {
            this.balance_rn = balance_rn;
            this.transactionDate = date;
            this.transactionPrice = price;
            this.balance_past = balance_past;
            this.accountNumber = accountNumber;
        }
        public Transaction(string writeDate,string transactionDate,int balance,int price,string accountNumber)
        {
            this.writeDate = writeDate;
            this.transactionDate = transactionDate;
            this.balance_rn = balance;
            this.transactionPrice = price;
            this.accountNumber = accountNumber;
        }
        public string getAccountNumber()
        {
            return accountNumber;
        }
        public int getBalance_rn()
        {
            return balance_rn;
        }
        public int getBalance_past()
        {
            return balance_past;
        }
        public string getTransactionDate()
        {
            return transactionDate;
        }
        public int getTransactionPrice()
        {
            return transactionPrice;
        }
    }
}
