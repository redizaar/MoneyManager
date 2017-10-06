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
        private string date;
        private int price;
        public Transaction(int balance_rn,string date,int price,int balance_past)
        {
            this.balance_rn = balance_rn;
            this.date = date;
            this.price = price;
            this.balance_past = balance_past;
        }
        public int getBalance_rn()
        {
            return balance_rn;
        }
        public int getBalance_past()
        {
            return balance_past;
        }
        public string getDate()
        {
            return date;
        }
        public int getPrice()
        {
            return price;
        }
    }
}
