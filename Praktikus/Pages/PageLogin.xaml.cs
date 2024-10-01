using Praktikus.Misc;
using System;
using System.Collections.Generic;
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

namespace Praktikus.Pages
{
    public partial class PageLogin : Page
    {
        public PageLogin()
        {
            InitializeComponent();
        }
        private void LoginBTN_Click(object sender, RoutedEventArgs e)
        {
            bool IsAdmin = false; bool norigh = true;
            foreach (var item in Connect.context.Users.ToList())
            {
                if (item.Login == LoginBox.Text && item.Password == PassBox.Password)
                {
                    if (item.Role == "ADMIN")
                    {
                        IsAdmin = true;
                    }
                    norigh = false;
                    Nav.MainFrame.Navigate(new MainMenu(IsAdmin));
                }  
            }
            if (norigh)
            {
                MessageBox.Show("Не правильный логин или пароль");
            }
        }
    }
}
