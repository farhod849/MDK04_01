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
    /// Логика взаимодействия для OrdersTableAdd.xaml
    /// </summary>
    public partial class OrdersTableAdd : Page
    {
        Orders order;
        bool checkNew;
        public OrdersTableAdd(Orders c)
        {
            InitializeComponent();
            DG2.ItemsSource = Connect.context.Orders.ToList();
            if (c == null)
            {
                c = new Orders();
                checkNew = true;
            }
            else
                checkNew = false;
            DataContext = order = c;
        }


        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            if (checkNew)
            {
                Connect.context.Orders.Add(order);
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
            DG2.ItemsSource = Connect.context.Orders.ToList();
        }
    }
}
