using Praktikus.Misc;
using System;
using System.Collections.Generic;
using System.Drawing;
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
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Page = System.Windows.Controls.Page;

namespace Praktikus.Pages
{
    /// <summary>
    /// Логика взаимодействия для UsersTable.xaml
    /// </summary>
    public partial class UsersTable : Page
    {
        public UsersTable()
        {
            InitializeComponent();
            DG5.ItemsSource = Connect.context.Users.ToList();
        }

        private void BackBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void AddBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new UsersTableAdd(null));
        }

        private void DelBTN_Click(object sender, RoutedEventArgs e)
        {
            var delUsers = DG5.SelectedItems.Cast<Users>().ToList();
            if (MessageBox.Show($"Удалить {delUsers.Count} записей?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.Users.RemoveRange(delUsers);
            try
            {
                Connect.context.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            DG5.ItemsSource = Connect.context.Managers.ToList();
        }

        private void UpdateBTN_Click(object sender, RoutedEventArgs e)
        {
            DG5.ItemsSource = Connect.context.Users.ToList();
        }

        private void SearchBTN_Click(object sender, RoutedEventArgs e)
        {
            var poisk = Connect.context.Users.Where(x => x.Login.ToString().StartsWith(SearchBox.Text)).ToList();
            DG5.ItemsSource = poisk;
        }

        private void ReportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application app = new Excel.Application()
            {
                Visible = true,
                SheetsInNewWorkbook = 1
            };
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing); app.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1); sheet.Name = "Учётная таблица";
            sheet.Cells[1, 1] = "ID Пользователя"; sheet.Cells[1, 2] = "Логин";
            sheet.Cells[1, 3] = "Пароль"; sheet.Cells[1, 4] = "Роль";
            var currentRow = 2;
            var prod = Connect.context.Users.ToList();
            foreach (var item in prod)
            {
                sheet.Cells[currentRow, 1] = item.ID_User;
                sheet.Cells[currentRow, 2] = item.Login;
                sheet.Cells[currentRow, 3] = item.Password;
                sheet.Cells[currentRow, 4] = item.Role;

                currentRow++;
            }
            Excel.Range range2 = sheet.get_Range("A1", "F20"); range2.Cells.Font.Name = "Cascadia mono";
            range2.Cells.Font.Size = 10; range2.Font.Bold = false;
            range2.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
            range2.Borders.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
        }

        private void Change_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
