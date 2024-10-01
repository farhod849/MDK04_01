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
using iTextSharp.text;
using iTextSharp.text.pdf;
using Font = iTextSharp.text.Font;
using PdfPCell = iTextSharp.text.pdf.PdfPCell;
using System.IO;

namespace Praktikus.Pages
{
    /// <summary>
    /// Логика взаимодействия для OrdersTable.xaml
    /// </summary>
    public partial class OrdersTable : Page
    {
        public bool IsAdmin;
        public OrdersTable(bool check)
        {
            InitializeComponent();
            DG2.ItemsSource = Connect.context.Orders.ToList();
            IsAdmin = check;
            AddBTN.Visibility = Visibility.Hidden;
            DelBTN.Visibility = Visibility.Hidden;
            ReportExcel.Visibility = Visibility.Hidden;
            ReportPDF.Visibility = Visibility.Hidden;
            if (IsAdmin)
            {
                AddBTN.Visibility = Visibility.Visible;
                DelBTN.Visibility = Visibility.Visible;
                ReportExcel.Visibility = Visibility.Visible;
                ReportPDF.Visibility = Visibility.Visible;
            }
        }

        private void BackBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void Change_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new OrdersTableAdd((sender as Button).DataContext as Orders));
        }

        private void AddBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new OrdersTableAdd(null));
        }

        private void DelBTN_Click(object sender, RoutedEventArgs e)
        {
            var delOrders = DG2.SelectedItems.Cast<Orders>().ToList();
            if (MessageBox.Show($"Удалить {delOrders.Count} записей?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.Orders.RemoveRange(delOrders);
            try
            {
                Connect.context.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            DG2.ItemsSource = Connect.context.Orders.ToList();
        }

        private void SearchBTN_Click(object sender, RoutedEventArgs e)
        {

            var poisk = Connect.context.Orders.Where(x => x.Order_Name.ToString().StartsWith(SearchBox.Text)).ToList();
            DG2.ItemsSource = poisk;
        }

        private void UpdateBTN_Click(object sender, RoutedEventArgs e)
        {
            DG2.ItemsSource = Connect.context.Orders.ToList();
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
            sheet.Cells[1, 1] = "ID заказа"; sheet.Cells[1, 2] = "Статус заказа";
            sheet.Cells[1, 3] = "Название заказа"; sheet.Cells[1, 4] = "Тип заказа";
            sheet.Cells[1, 5] = "Текущая рабочая группа"; sheet.Cells[1, 6] = "Руководитель заказа";
            sheet.Cells[1, 7] = "Итоговая стоимость заказа";
            var currentRow = 2;
            var prod = Connect.context.Orders.ToList();
            foreach (var item in prod)
            {
                sheet.Cells[currentRow, 1] = item.ID_Order;
                sheet.Cells[currentRow, 2] = item.Order_Status;
                sheet.Cells[currentRow, 3] = item.Order_Name;
                sheet.Cells[currentRow, 4] = item.Order_Type;
                sheet.Cells[currentRow, 5] = item.Current_WorkGroup;
                sheet.Cells[currentRow, 6] = item.Order_Manager;
                sheet.Cells[currentRow, 7] = item.TotalPriceOrder;
                currentRow++;
            }
            Excel.Range range2 = sheet.get_Range("A1", "I20"); range2.Cells.Font.Name = "Cascadia mono";
            range2.Cells.Font.Size = 10; range2.Font.Bold = false;
            range2.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
            range2.Borders.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
        }

        private void ReportPDF_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Document doc = new Document();
                PdfWriter.GetInstance(doc, new FileStream("Отчет по таблице Заказы.pdf", FileMode.Create));
                doc.Open();
                BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\Arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                Font font = new Font(baseFont, Font.DEFAULTSIZE, Font.NORMAL);
                PdfPTable table = new PdfPTable(1);
                PdfPTable table2 = new PdfPTable(1);
                PdfPTable table3 = new PdfPTable(1);
                PdfPTable table4 = new PdfPTable(1);
                PdfPTable table5 = new PdfPTable(1);
                PdfPTable table6 = new PdfPTable(1);
                PdfPTable table7 = new PdfPTable(1);
               
                PdfPCell cell = new PdfPCell(new Phrase("Ведомость заказов", font));
                cell.Colspan = 7;
                cell.HorizontalAlignment = 1;
                cell.Border = 0;
                table.AddCell(cell);
                table2.AddCell(cell);
                table3.AddCell(cell);
                table4.AddCell(cell);
                table5.AddCell(cell);
                table6.AddCell(cell);
                table7.AddCell(cell);
               

                table.AddCell(new PdfPCell(new Phrase(new Phrase("ID Заказа", font))));
                table2.AddCell(new PdfPCell(new Phrase(new Phrase("Статус заказа", font))));
                table3.AddCell(new PdfPCell(new Phrase(new Phrase("Название заказа", font))));
                table4.AddCell(new PdfPCell(new Phrase(new Phrase("Тип заказа", font))));
                table5.AddCell(new PdfPCell(new Phrase(new Phrase("Текущая рабочая группа", font))));
                table6.AddCell(new PdfPCell(new Phrase(new Phrase("Руководитель заказа", font))));
                table7.AddCell(new PdfPCell(new Phrase(new Phrase("Итоговая цена заказа", font))));
                var a = 0;
                for (int i = 0; i < Connect.context.Orders.ToList().Count; i++)
                {
                    var itemOrder = Connect.context.Orders.ToList()[i];
                    table.AddCell(new Phrase(itemOrder.ID_Order.ToString(), font));
                    table2.AddCell(new Phrase(itemOrder.Order_Status.ToString(), font));
                    table3.AddCell(new Phrase(itemOrder.Order_Name.ToString(), font));
                    table4.AddCell(new Phrase(itemOrder.Order_Type.ToString(), font));
                    table5.AddCell(new Phrase(itemOrder.Current_WorkGroup.ToString(), font));
                    table6.AddCell(new Phrase(itemOrder.Order_Manager.ToString(), font));
                    table7.AddCell(new Phrase(itemOrder.TotalPriceOrder.ToString(), font));
                   
                   
                }
                doc.Add(table);
                doc.Add(table2);
                doc.Add(table3);
                doc.Add(table4);
                doc.Add(table5);
                doc.Add(table6);
                doc.Add(table7);
                doc.Close();
                MessageBox.Show("Pdf-документ сохранен");
            }
            catch
            {
                MessageBox.Show("Pdf-документ не сохранен", "Ошибка!");
            }
        }
    }
}
