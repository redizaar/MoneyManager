using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    public class TransactionCategory
    {
        public string mainCategoryName { get; set; }
        public string categoryName { get; set; }
        public bool isMainCategory { get; set; }
        public bool hasBudget { get; set; }
        public int maxBudgetLimit { get; set; }

        public bool addCategory(string categoryName,bool isMainCategory,string mainCategoryName)
        {
            return true;
        }
    }
}
