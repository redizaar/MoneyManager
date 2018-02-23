using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WPFCustomMessageBox;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for ImportMainPage.xaml
    /// </summary>
    public partial class ImportMainPage : System.Windows.Controls.Page
    {
        private ButtonCommands btnCommand;
        private MainWindow mainWindow;
        private User currentUser;
        private static ImportMainPage instance;
        public bool alwaysAsk
        {
            get
            {
                if(alwaysAskCB.IsChecked.Equals(true))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                if(value)
                {
                    neverAskCB.SetCurrentValue(RadioButton.IsCheckedProperty, false);
                }
            }
        }
        public bool neverAsk
        {
            get
            {
                if (neverAskCB.IsChecked.Equals(true))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                if (value)
                {
                    alwaysAskCB.SetCurrentValue(RadioButton.IsCheckedProperty, false);
                }
            }
        }
        private ImportMainPage(MainWindow mainWindow)
        {
            DataContext = this;
            InitializeComponent();
            neverAskCB.IsChecked = true;
            descriptionComboBox.Visibility = System.Windows.Visibility.Hidden;

            this.mainWindow = mainWindow;
            this.currentUser = mainWindow.getCurrentUser();
            if (currentUser.getAccountNumber().Equals(mainWindow.getAccounNumber()))
            {
                getUserStatistics(currentUser);
            }
            else
            {
                getUserStatistics(mainWindow.getAccounNumber());
            }
            FolderAddressLabel.Visibility = System.Windows.Visibility.Hidden;
        }
        public void getUserStatistics(string accountNumber)
        {
            urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
            int totalIncome = 0;
            int totalSpendings = 0;
            int numberOfTransactions = 0;
            string lastImportDate = "";
            string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
            DateTime todayDate = Convert.ToDateTime(todaysDate);
            usernameLabel.Content = "Test";
            foreach (var transactions in SavedTransactions.getSavedTransactions())
            {
                if (transactions.getAccountNumber().Equals(accountNumber))
                {
                    numberOfTransactions++;
                    lastImportDate = transactions.getWriteDate();//always overwrites it --- todo (more logic needed lulz)
                    if(transactions.getTransactionPrice()>0)
                    {
                        totalIncome += transactions.getTransactionPrice();
                    }
                    else
                    {
                        totalSpendings += transactions.getTransactionPrice();
                    }
                }
            }
            if (lastImportDate == null)
            {
                lastImportDate= DateTime.Now.ToString("yyyy-MM-dd");
            }
            if (lastImportDate.Length > 12)
            {
                lastImportDateLabel.Content = lastImportDate.Substring(0, 12);
            }
            else
            {
                lastImportDateLabel.Content = lastImportDate;
            }
            noTransactionsLabel.Content = numberOfTransactions;
            if (lastImportDate.Length > 0)
            {
                DateTime importDate = Convert.ToDateTime(lastImportDate);
                int diffDays = (todayDate - importDate).Days;
                if (diffDays >= 30)
                {
                    urgencyLabel.Content = "Recommended!";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(217, 30, 24));
                }
                else
                {
                    urgencyLabel.Content = "Not urgent";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(46, 204, 113));
                }
            }
            else
            {
                urgencyLabel.Content = "You haven't imported yet!";
                lastImportDateLabel.Content = "You haven't imported yet!";
            }
        }

        private void getUserStatistics(User currentUser)
        {
            int numberOfTransactions = 0;
            int totalIncome = 0;
            int totalSpendings = 0;
            string latestImportDate = "";
            string todaysDate = DateTime.Now.ToString("yyyy-MM-dd");
            DateTime todayDate = Convert.ToDateTime(todaysDate);
            usernameLabel.Content = currentUser.getUsername();
            foreach (var transactions in SavedTransactions.getSavedTransactions())
            {
                if(transactions.getAccountNumber().Equals(currentUser.getAccountNumber()))
                {
                    numberOfTransactions++;
                    latestImportDate = transactions.getWriteDate();//always overwrites it --- todo (more logic needed lulz)
                    if (transactions.getTransactionPrice() > 0)
                    {
                        totalIncome += transactions.getTransactionPrice();
                    }
                    else
                    {
                        totalSpendings += transactions.getTransactionPrice();
                    }
                }
            }
            if (latestImportDate.Length > 12)
            {
                lastImportDateLabel.Content = latestImportDate.Substring(0, 12);
            }
            else
            {
                lastImportDateLabel.Content = latestImportDate;
            }
            noTransactionsLabel.Content = numberOfTransactions;
            DateTime importDate;
            if (latestImportDate.Length > 0)
            {
                importDate = Convert.ToDateTime(latestImportDate);
                float diffTicks = (todayDate - importDate).Days;
                if (diffTicks >= 30)
                {
                    urgencyLabel.Content = "Recommended!";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(217, 30, 24));
                    mainWindow.exclamImage.Visibility = System.Windows.Visibility.Visible;
                }
                else
                {
                    urgencyLabel.Content = "Not urgent";
                    urgencyLabel.Foreground = new SolidColorBrush(Color.FromRgb(46, 204, 113));
                    mainWindow.exclamImage.Visibility = System.Windows.Visibility.Hidden;
                }
            }
            else
            {
                urgencyLabel.Content = "You haven't imported yet!";
                lastImportDateLabel.Content = "You haven't imported yet!";
            }
        }
        private void getTransactions(string bankName, List<string> folderAddress)
        {
            new ImportReadIn(bankName, folderAddress, mainWindow,false);
        }

        public ButtonCommands OpenFilePushed
        {
            get
            {
                btnCommand = new ButtonCommands(FileBrowser.Content.ToString(),this);
                return btnCommand;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public static ImportMainPage getInstance(MainWindow mainWindow)
        {
            if(instance==null)
            {
                instance = new ImportMainPage(mainWindow);
            }
            return instance;
        }
        public class ButtonCommands : ICommand
        {
            private string buttonContent;
            private ImportMainPage importPage;
            public ButtonCommands(string buttonContent,ImportMainPage importPage)
            {
                this.buttonContent = buttonContent;
                this.importPage = importPage;

                this.importPage.PropertyChanged += new PropertyChangedEventHandler(test_PropertyChanged);
            }
            private void test_PropertyChanged(object sender, PropertyChangedEventArgs e)
            {
                if (CanExecuteChanged != null)
                {
                    CanExecuteChanged(this, EventArgs.Empty);
                }
            }
            public event EventHandler CanExecuteChanged;

            public bool CanExecute(object parameter)
            {
                //todo
                return true;
            }
            public static int ExcelColumnNameToNumber(string columnName)
            {
                if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

                columnName = columnName.ToUpperInvariant();

                int sum = 0;

                for (int i = 0; i < columnName.Length; i++)
                {
                    sum *= 26;
                    sum += (columnName[i] - 'A' + 1);
                }

                return sum;
            }
            private void check_if_csv(int fileIndex, Microsoft.Win32.OpenFileDialog dialog)
            {
                string[] fileName = dialog.FileNames.ToList()[fileIndex].Split('\\');
                int lastPartIndex = fileName.Length - 1;
                Regex csvPattern = new Regex(@".csv$");
                if (csvPattern.IsMatch(fileName[lastPartIndex]))
                {
                    string newExcelPath = dialog.FileNames[fileIndex].Substring(0, dialog.FileNames[fileIndex].Length - 4);
                    string csv = dialog.FileNames[fileIndex];
                    string xls = newExcelPath + ".xls";

                    List<List<string>> allWords = new List<List<string>>();
                    IEnumerable<String> all_lines = System.IO.File.ReadLines(dialog.FileNames[fileIndex], Encoding.Default);
                    foreach (var lines in all_lines)
                    {
                        List<string> words = lines.Split(';').ToList();
                        allWords.Add(words);
                    }
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = app.Workbooks.Open(dialog.FileNames[fileIndex], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Worksheet sheet = wb.Worksheets[1];
                    int row = 1;
                    foreach (List<string> lines in allWords)
                    {
                        int column = 1;
                        for (int itr = 0; itr < lines.Count; itr++)
                        {
                            sheet.Cells[row, column].Value = lines[itr];
                            column++;
                        }
                        row++;
                    }
                    wb.SaveAs(newExcelPath, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wb.Close();
                    app.Quit();
                    dialog.FileNames[fileIndex] = newExcelPath; //overwriting the old string
                }
            }
            public System.Data.DataTable checkIfFileInformationStored(string filePath)
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excel.Workbooks.Open(filePath);
                Worksheet worksheet = workbook.Worksheets[1];
                SqlConnection sqlConn = new SqlConnection("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=ImportFileData;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
                sqlConn.Open();
                string getEveryRow = "Select * From [StoredColumns]";
                SqlDataAdapter sda = new SqlDataAdapter(getEveryRow, sqlConn);
                System.Data.DataTable dtb = new System.Data.DataTable();
                sda.Fill(dtb);
                if (dtb.Rows.Count > 0)
                {
                    foreach (DataRow row in dtb.Rows)
                    {
                        int trasactionsRow = int.Parse(row["TransStartRow"].ToString());
                        string accountNumberPosString = row["AccountNumberPos"].ToString();
                        string dateColumnString = row["DateColumn"].ToString();
                        string priceColumnString = row["PriceColumn"].ToString();
                        string balanceColumnString = row["BalanceColumn"].ToString();
                        string commentColumnString = row["CommentColumn"].ToString();

                        int dateColumn;
                        try
                        {
                            dateColumn=int.Parse(dateColumnString);
                        }
                        catch(Exception e)
                        {
                            dateColumn=ExcelColumnNameToNumber(dateColumnString);
                        }
                        int balanceColumn=-1;
                        if(dateColumnString!="None")
                        {
                            try
                            {
                                balanceColumn = int.Parse(balanceColumnString);
                            }
                            catch (Exception e)
                            {
                                balanceColumn = ExcelColumnNameToNumber(balanceColumnString);
                            }
                        }
                        List<int> accountNumberPos = new List<int>(); 
                        // if it has 2 elements its in a cell
                        // if it has 1 element it is a column
                        if(accountNumberPosString!="Sheet name")
                        {
                            int tempValue1=0;
                            long size = sizeof(char) * accountNumberPosString.Length;
                            //todo
                            if(size>1)//its a cell 
                            {
                                int tempValue2 = 0;
                                try
                                {
                                    tempValue1 = int.Parse(accountNumberPosString[1].ToString());
                                }
                                catch(Exception e)
                                {
                                    tempValue1 = ExcelColumnNameToNumber(accountNumberPosString[1].ToString());
                                }
                                try
                                {
                                    tempValue2 = int.Parse(accountNumberPosString[0].ToString());
                                }
                                catch (Exception e)
                                {
                                    tempValue2 = ExcelColumnNameToNumber(accountNumberPosString[0].ToString());
                                }
                                accountNumberPos.Add(tempValue1);
                                accountNumberPos.Add(tempValue2);
                            }
                            else if(size == 1)
                            {
                                try
                                {
                                    tempValue1 = int.Parse(accountNumberPosString);
                                }
                                catch (Exception e)
                                {
                                    balanceColumn = ExcelColumnNameToNumber(accountNumberPosString);
                                }
                                accountNumberPos.Add(tempValue1);
                            }
                        }
                        else
                        {
                            accountNumberPos = null;
                        }
                        List<int> commentColumns = new List<int>();
                        string[] commentColumnsSplitted = commentColumnString.Split(',');
                        for(int i=0;i<commentColumnsSplitted.Length;i++)
                        {
                            int tempValue;
                            try
                            {
                                tempValue = int.Parse(commentColumnsSplitted[i]);
                            }
                            catch(Exception e )
                            {
                                tempValue = ExcelColumnNameToNumber(commentColumnsSplitted[i]);
                            }
                            commentColumns.Add(tempValue);
                        }
                        List<int> priceColumn = new List<int>();
                        string[] priceColumnsSplitted = priceColumnString.Split(',');
                        if(priceColumnsSplitted.Length>1)
                        {
                            for (int i = 0; i < priceColumnsSplitted.Length; i++)
                            {
                                int tempValue;
                                try
                                {
                                    tempValue = int.Parse(priceColumnsSplitted[i]);
                                }
                                catch (Exception e)
                                {
                                    tempValue = ExcelColumnNameToNumber(priceColumnsSplitted[i]);
                                }
                                priceColumn.Add(tempValue);
                            }
                        }
                        else
                        {
                            int tempValue;
                            try
                            {
                                tempValue = int.Parse(priceColumnsSplitted[0]);
                            }
                            catch (Exception e)
                            {
                                tempValue = ExcelColumnNameToNumber(priceColumnsSplitted[0]);
                            }
                            priceColumn.Add(tempValue);
                        }
                    }
                    return dtb;
                }
                return null;
            }
            public void Execute(object parameter)
            {
                if (buttonContent.Equals("Import Transactions"))
                {
                    MessageBoxResult messageBoxResult = CustomMessageBox.ShowYesNo(
                        "\tPlease choose an import type!",
                        "Import type alert!",
                        "Automatized",
                        "User specified");
                    if (messageBoxResult == MessageBoxResult.Yes || messageBoxResult==MessageBoxResult.No)
                    {
                        Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                        dlg.DefaultExt = ".xls,.csv";
                        dlg.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xlsm)|*.xlsm|CSV Files (*.csv)|*.csv";
                        dlg.Multiselect = true;
                        Nullable<bool> result = dlg.ShowDialog();
                        if (result == true)
                        {
                            for (int i = 0; i < dlg.FileNames.ToList().Count; i++)
                            {
                                check_if_csv(i,dlg);
                            }
                            //importPage.FolderAddressLabel.Content = dlg.FileName.;
                            if (messageBoxResult == MessageBoxResult.Yes)
                            {
                                importPage.getTransactions("All", dlg.FileNames.ToList());
                            }
                            else if (messageBoxResult == MessageBoxResult.No)
                            {
                                string[] fileName = dlg.FileNames.ToList()[0].Split('\\');
                                int lastPartIndex = fileName.Length - 1; // to see which file the user immporting first
                                SpecifiedImport.getInstance(dlg.FileNames.ToList(), importPage.mainWindow).setCurrentFileLabel(fileName[lastPartIndex]);
                                //fájl felismerés
                                checkIfFileInformationStored(dlg.FileNames.ToList()[0]);
                                importPage.mainWindow.MainFrame.Content = SpecifiedImport.getInstance(dlg.FileNames.ToList(), importPage.mainWindow);
                            }
                        }
                    }
                }
            }
        }
    }
}
