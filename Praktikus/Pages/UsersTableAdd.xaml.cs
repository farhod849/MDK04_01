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
    /// <summary>
    /// Логика взаимодействия для UsersTableAdd.xaml
    /// </summary>
    public partial class UsersTableAdd : Page
    {
        Users user;
        bool checkNew;
        public UsersTableAdd(Users c)
        {
            InitializeComponent();
            DG5.ItemsSource = Connect.context.Users.ToList();
            if (c == null)
            {
                c = new Users();
                checkNew = true;
            }
            else
                checkNew = false;
            DataContext = user = c;
        }


        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            if (checkNew)
            {
                Connect.context.Users.Add(user);
            }
            try
            {
                Connect.context.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Connect.context.SaveChanges();
            DG5.ItemsSource = Connect.context.Users.ToList();
        }
    }
}
