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
    /// Логика взаимодействия для ManagersTableAdd.xaml
    /// </summary>
        public partial class ManagersTableAdd : Page
        {
            Managers manager;
            bool checkNew;
            public ManagersTableAdd(Managers c)
            {
                InitializeComponent();
                DG1.ItemsSource = Connect.context.Managers.ToList();
                if (c == null)
                {
                    c = new Managers();
                    checkNew = true;
                }
                else
                    checkNew = false;
                DataContext = manager = c;
            }

            private void BackButton_Click(object sender, RoutedEventArgs e)
            {
                Nav.MainFrame.GoBack();
            }

            private void AddButton_Click(object sender, RoutedEventArgs e)
            {
                if (checkNew)
                {
                    Connect.context.Managers.Add(manager);
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
                DG1.ItemsSource = Connect.context.Managers.ToList();
            }
        }
}
