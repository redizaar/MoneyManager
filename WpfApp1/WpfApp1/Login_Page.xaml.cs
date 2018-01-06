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
    /// Interaction logic for Login_Page.xaml
    /// </summary>
    public partial class Login_Page : Page
    {
        public Login_Page()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection sqlConn = new SqlConnection("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=LoginDB;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            string loginQuery = "Select * from dbo Where username = '"+usernameTextbox.Text.ToString()+"' and password = '"+passwordTextbox.Password.ToString()+"'";
            SqlDataAdapter sda = new SqlDataAdapter(loginQuery,sqlConn);
            DataTable dtb = new DataTable();
            sda.Fill(dtb);
            if(dtb.Rows.Count==1)
            {
                this.Visibility = System.Windows.Visibility.Hidden;
            }
        }
    }
}
