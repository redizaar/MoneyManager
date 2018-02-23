using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
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
    /// Interaction logic for Register_Page.xaml
    /// </summary>
    public partial class Register_Page : Page
    {
        private MainWindow mainWindow;
        public Register_Page(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
            InitializeComponent();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SqlConnection sqlConn = new SqlConnection("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=LoginDB;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            sqlConn.Open();
            SqlCommand sqlCommand = new SqlCommand("registrationQuery", sqlConn);
            sqlCommand.CommandType = CommandType.StoredProcedure;
            bool usernameCorrect=checkUsername();
            bool passwordCorrect=checkPassword();
            if (usernameCorrect && passwordCorrect)
            {
                sqlCommand.Parameters.AddWithValue("@username", registerUsernameTextbox.Text.ToString());
                sqlCommand.Parameters.AddWithValue("@password", registerPasswordTextbox.Password.ToString());
                sqlCommand.Parameters.AddWithValue("@failedlogins", 0);
                sqlCommand.ExecuteNonQuery();
                if (MessageBox.Show("You can log in now!", "Successfull registartion!",
                         MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                {
                    mainWindow.LoginFrame.Content = new Login_Page(mainWindow);
                }
            }

        }

        private bool checkPassword()
        {
            //todo
            return true;
        }

        private bool checkUsername()
        {
            //todo
            return true;
        }
    }
}
