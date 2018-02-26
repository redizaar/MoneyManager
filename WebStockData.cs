﻿using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;

namespace WpfApp1
{
    public class WebStockData
    {
        private List<Stock> stocks;
        private List<string> dates;
        private List<double> prices;
        public WebStockData()
        {
            stocks = new List<Stock>();
            //todo
            //cant call this class more than one whitin a minute
            //because the ip address will get blocked
        }
        public void getCSVDataFromGoogle(string ticker,string day,string month,string year)
        {
            dates = new List<string>();
            prices = new List<double>();
            string csv;
            using (var web = new WebClient())
            {
                var url = $"https://finance.google.com/finance/historical?q="+ticker+"&startdate="+day+"-"+month+"-"+year+"&output=csv";
                //$"https://finance.google.com/finance/historical?q=AAPL&startdate=01-Jan-2016&output=csv";
                csv = web.DownloadString(url);
            }
            string[] lines = csv.Split(',');
            int j = 0;
            string regex = "[0-9]{2}-[a-zA-Z]{3}-[0-9]{2}";
            string regex2 = "[0-9]-[a-zA-Z]{3}-[0-9]{2}";
            for (int i = 0; i < lines.Length; i++)
            {
                if ((Regex.IsMatch(lines[i], regex)) || (Regex.IsMatch(lines[i], regex2)))
                {
                    string[] date = lines[i].Split('\n');
                    dates.Add(date[1]);
                }
                if (i > 4 && j == 4)
                {
                    prices.Add(double.Parse(lines[i].Replace('.',',')));
                    j = 0;
                }
                else if (i > 4)
                    j++;
            }
            ThreadStart threadStart = delegate
            {
                writeStocksToSQL(ticker, dates ,prices);
            };
            Thread sqlThread = new Thread(threadStart);
            sqlThread.IsBackground = true;
            sqlThread.Start();
            sqlThread.Join();
        }
        public void GetDataFromWeb()
        {
            Stock stock;
            const string tickers = "AAPL,GOOG,GOOGL,YHOO,TSLA,INTC,AMZN,BIDU,ORCL,MSFT,ORCL,ATVI,NVDA,GME,LNKD,NFLX";

            string json;

            using (var web = new WebClient())
            {
                var url = $"https://finance.google.com/finance?q=AAPL&output=json";
                json = web.DownloadString(url);
            }

            //Google adds a comment before the json for some unknown reason, so we need to remove it
            json = json.Replace("//", "");

            var v = JArray.Parse(json);

            foreach (var i in v)
            {
                var ticker = i.SelectToken("t");
                var price = (float)i.SelectToken("l");
                //var lastTradeTime = i.SelectToken("ltt");
                //var change = i.SelectToken("c");
                //var changePercentage = i.SelectToken("cp");
                stock = new Stock(ticker.ToString(), float.Parse(price.ToString()));
                stocks.Add(stock);
                //Console.WriteLine($"{ticker} : {price}");
            }
        }
        private void writeStocksToSQL(string ticker,List<string> new_dates,List<double> new_prices)
        {
            for (int i = 0; i < new_dates.Count; i++)
                Console.WriteLine(new_dates[i]);
            //elől vannak a friss dátumok, árak
            //atatbázusba nyilván fordítva
            //https://stackoverflow.com/questions/41161104/error-converting-data-type-varchar-to-float-c-sharp-webservice
            string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
            SqlConnection sqlConn = new SqlConnection("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=StockData;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            sqlConn.Open();
            string loginQuery = "Select * From [Stock_WebData] where Name = '"+ticker+"'";
            SqlDataAdapter sda = new SqlDataAdapter(loginQuery, sqlConn);
            DataTable dtb = new DataTable();
            sda.Fill(dtb);
            if (dtb.Rows.Count == 0)
            {
                SqlCommand sqlCommand = new SqlCommand("insertStockDataNew4", sqlConn);//SQLQuery 4
                sqlCommand.CommandType = CommandType.StoredProcedure;
                for (int i = new_dates.Count-1; i > 0; i--)
                {
                    sqlCommand.Parameters.Clear();
                    sqlCommand.Parameters.AddWithValue("@name", ticker);
                    sqlCommand.Parameters.AddWithValue("@price", new_prices[i]);
                    sqlCommand.Parameters.AddWithValue("@date", new_dates[i]);
                    sqlCommand.ExecuteNonQuery();
                }
            }
            else
            {
                bool storedinSql;
                List<int> notStoredIndexes = new List<int>();
                for (int i = 0; i < new_dates.Count; i++)
                {
                    storedinSql = false;
                    foreach (DataRow row in dtb.Rows)
                    {
                        string dateFromSql = row["Date"].ToString();
                        if(new_dates[i]==dateFromSql)
                        {
                            storedinSql = true;
                            break;
                        }
                        //DateTime dt1 = DateTime.ParseExact(dateFromSql, "dd-MMM-yy", System.Globalization.CultureInfo.InvariantCulture);
                        //converts a string to a date fromat for example : 27-feb-18
                    }
                    if (!storedinSql)
                        notStoredIndexes.Add(i);
                }
                if (notStoredIndexes.Count > 0)
                {
                    SqlCommand sqlCommand = new SqlCommand("insertStockDataNew4", sqlConn);//SQLQuery 4
                    sqlCommand.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < notStoredIndexes.Count; i++)
                    {
                        sqlCommand.Parameters.Clear();
                        sqlCommand.Parameters.AddWithValue("@name", ticker);
                        sqlCommand.Parameters.AddWithValue("@price", new_prices[notStoredIndexes[i]]);
                        sqlCommand.Parameters.AddWithValue("@date", new_dates[notStoredIndexes[i]]);
                        sqlCommand.ExecuteNonQuery();
                    }
                }
            }
        }
        public bool writeStocksToSQL()
        {
            //https://stackoverflow.com/questions/41161104/error-converting-data-type-varchar-to-float-c-sharp-webservice
            //string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
            var src = DateTime.Now;
            var exactDate = new DateTime(src.Year, src.Month, src.Day, src.Hour, src.Minute, 0);
            SqlConnection sqlConn = new SqlConnection("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=StockData;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            sqlConn.Open();
            SqlCommand sqlCommand = new SqlCommand("insertStockDataNew4", sqlConn);//SQLQuery 4
            sqlCommand.CommandType = CommandType.StoredProcedure;
            for (int i = 0; i < stocks.Count; i++)
            {
                sqlCommand.Parameters.AddWithValue("@name", stocks[i].getStockName());
                sqlCommand.Parameters.AddWithValue("@price", stocks[i].getStockPrice());
                sqlCommand.Parameters.AddWithValue("@date", exactDate.ToString());
                sqlCommand.ExecuteNonQuery();
            }
            return true;
        }
        public List<double> getPrices()
        {
            return prices;
        }
        public List<string> getDates()
        {
            return dates;
        }
    }
}
