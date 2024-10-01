using Microsoft.Win32;
using Praktikus.Misc;
using Praktikus.Windows;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
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
    public partial class MainMenu : Page
    {
      public bool IsAdmin;

        public MainMenu(bool check)
        {
            InitializeComponent();
            
          IsAdmin = check;
            TableSalaryCalc.Visibility = Visibility.Hidden;
            TableOrders.Visibility = Visibility.Hidden;
            BackUpBD.Visibility = Visibility.Hidden;
            TableUser.Visibility = Visibility.Hidden;
          if (IsAdmin)
          {
              TableSalaryCalc.Visibility = Visibility.Visible;
              TableOrders.Visibility = Visibility.Visible;
              BackUpBD.Visibility = Visibility.Visible;
              TableUser.Visibility = Visibility.Visible;

            }
         
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void TableManagers_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new ManagersTable(IsAdmin));
        }

        private void TableOrders_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new OrdersTable(IsAdmin));
        }

        private void TableSalaryCalc_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new SalaryCalcTable(IsAdmin));
        }

        private void TableWorkGroup_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new WokrGroupTable(IsAdmin));
        }

        private void LogOutBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void BackUpBD_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog();
            dialog.Filter = "Резервная копия(*.bak)|*.bak|Все файлы(*.*)|*.*";
            bool? result = dialog.ShowDialog();
            if (result == true)
                Connect.context.Database.ExecuteSqlCommand(TransactionalBehavior.DoNotEnsureTransaction,
                    $@"BACKUP DATABASE [{Directory.GetCurrentDirectory()}\Sabirov_BD.mdf] TO  " +
                $@"DISK = N'{dialog.FileName}' WITH NOFORMAT, NOINIT,  " +
                $@"NAME = N'{dialog.FileName}', SKIP, NOREWIND, NOUNLOAD,  STATS = 10");
            MessageBox.Show("Резервная копия создана");
        }

        private void TableUser_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new UsersTable());
        }
    }
}
