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
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Page = System.Windows.Controls.Page;
using Button = System.Windows.Controls.Button;
using System.Drawing;
using System.Security.Cryptography;

namespace Praktikus.Pages
{
    /// <summary>
    /// Логика взаимодействия для SalaryCalcTable.xaml
    /// </summary>
    public partial class SalaryCalcTable : Page
    {
        public bool IsAdmin;
        public SalaryCalcTable(bool check)
        {
            InitializeComponent();
            DG3.ItemsSource = Connect.context.SalaryCalc.ToList();
            IsAdmin = check;
            AddBTN.Visibility = Visibility.Hidden;
            DelBTN.Visibility = Visibility.Hidden;
            ReportExcel.Visibility = Visibility.Hidden;
            ReportPDF.Visibility = Visibility.Hidden;
            SalaryCalcBTN.Visibility = Visibility.Hidden;
            if (IsAdmin)
            {
                AddBTN.Visibility = Visibility.Visible;
                DelBTN.Visibility = Visibility.Visible;
                ReportExcel.Visibility = Visibility.Visible;
                ReportPDF.Visibility = Visibility.Visible;
                SalaryCalcBTN.Visibility = Visibility.Visible;
            }
        }

        private void BackBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void Change_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new SalaryCalcTableAdd((sender as Button).DataContext as SalaryCalc));
        }

        private void AddBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new SalaryCalcTableAdd(null));
        }

        private void DelBTN_Click(object sender, RoutedEventArgs e)
        {
            var delSalary = DG3.SelectedItems.Cast<SalaryCalc>().ToList();
            if (MessageBox.Show($"Удалить {delSalary.Count} записей?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.SalaryCalc.RemoveRange(delSalary);
            try
            {
                Connect.context.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            DG3.ItemsSource = Connect.context.SalaryCalc.ToList();
        }

        private void SearchBTN_Click(object sender, RoutedEventArgs e)
        {

            var poisk = Connect.context.SalaryCalc.Where(x => x.ID_Ticket.ToString().StartsWith(SearchBox.Text)).ToList();
            DG3.ItemsSource = poisk;
        }

        private void UpdateBTN_Click(object sender, RoutedEventArgs e)
        {
            DG3.ItemsSource = Connect.context.SalaryCalc.ToList();
            MessageBox.Show("Таблица обновлена!", "Обновление", MessageBoxButton.OK, MessageBoxImage.Information);
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
            sheet.Cells[1, 1] = "ID Документа"; sheet.Cells[1, 2] = "ID Заказа";
            sheet.Cells[1, 3] = "ID Руководителя"; sheet.Cells[1, 4] = "ЗП для руководителя";
            sheet.Cells[1, 5] = "ID Рабочей группы"; sheet.Cells[1, 6] = "ЗП для рабочей группы";
            sheet.Cells[1, 7] = "ЗП каждого работника группы";
            var currentRow = 2;
            var prod = Connect.context.SalaryCalc.ToList();
            var work = Connect.context.WorkGroup.ToList();
            var ord = Connect.context.Orders.ToList();

            double a = 0;
            double b = 0;
            double c = 0;
            foreach (var item in prod)
            {
                foreach (var item2 in work)
                {
                    foreach (var item3 in ord)
                    {
                        sheet.Cells[currentRow, 1] = item.ID_Ticket;
                        sheet.Cells[currentRow, 2] = item.ID_Of_Order;
                        sheet.Cells[currentRow, 3] = item.ID_Of_Manager;
                        sheet.Cells[currentRow, 4] = (item.Orders.TotalPriceOrder * 10) / 100;
                        sheet.Cells[currentRow, 5] = item.ID_Of_Workgroup;
                        sheet.Cells[currentRow, 6] = item.Orders.TotalPriceOrder - ((item.Orders.TotalPriceOrder * 10) / 100);
                        sheet.Cells[currentRow, 7] = (item.Orders.TotalPriceOrder - ((item.Orders.TotalPriceOrder * 10) /100)) / item.WorkGroup.Number_Of_Workers;
                    }
                }
                currentRow++;
            }
            
            Excel.Range range2 = sheet.get_Range("A1", "I20"); range2.Cells.Font.Name = "Cascadia mono";
            range2.Cells.Font.Size = 10; range2.Font.Bold = false;
            range2.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
            range2.Borders.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
        }

        private void ReportPDF_Click(object sender, RoutedEventArgs e)
        {

        }

        private void SalaryCalcBTN_Click(object sender, RoutedEventArgs e)
        {
            var prod = Connect.context.SalaryCalc.ToList();
            var work = Connect.context.WorkGroup.ToList();
            var ord = Connect.context.Orders.ToList();

            int a = 0;
            int b = 0;
            int c = 0;
            foreach (var item in prod)
            {
                foreach (var item2 in work)
                {
                    foreach (var item3 in ord)
                    {
                        a = (int)((item.Orders.TotalPriceOrder * 10) / 100);
                        b = (int)(item.Orders.TotalPriceOrder - ((item.Orders.TotalPriceOrder * 10) / 100));
                        c = (int)((item.Orders.TotalPriceOrder - ((item.Orders.TotalPriceOrder * 10) / 100)) / item.WorkGroup.Number_Of_Workers);
                    }
                }
                
            }

            MessageBox.Show("ЗП руководителя: " + a + " ЗП на всю рабочую группу: " + b + " Каждый работник получит по: " + c);
        }
    }
}
