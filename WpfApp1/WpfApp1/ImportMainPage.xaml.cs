using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for ImportMainPage.xaml
    /// </summary>
    public partial class ImportMainPage : Page
    {
        private ButtonCommands btnCommand;
        public ImportMainPage()
        {
            InitializeComponent();

            FolderAddressLabel.Visibility = System.Windows.Visibility.Hidden;
        }


        private void getTransactions(string bankName, string folderAddress)
        {
            //new ImportReadIn(bankName, folderAddress, this);
        }

        public ButtonCommands OpenFilePushed
        {
            get
            {
                btnCommand = new ButtonCommands(FileBrowser.Content.ToString(), this);

                return btnCommand;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public class ButtonCommands : ICommand
        {
            private string buttonContent;
            private ImportMainPage importPage;
            public ButtonCommands(string buttonContent, ImportMainPage importPage)
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

            public void Execute(object parameter)
            {
                if (buttonContent.Equals("Import Transactions"))
                {
                    Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                    dlg.DefaultExt = ".xls";
                    dlg.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xlsm)|*.xlsm";
                    Nullable<bool> result = dlg.ShowDialog();
                    if (result == true)
                    {
                        importPage.FolderAddressLabel.Content = dlg.FileName;
                    }
                    importPage.getTransactions(importPage.banksComboBox.Text, importPage.FolderAddressLabel.Content.ToString());
                }
            }
        }
    }
}
