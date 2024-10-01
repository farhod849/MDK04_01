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
    /// Логика взаимодействия для WokrGroupTable.xaml
    /// </summary>
    public partial class WokrGroupTable : Page
    {
        public bool IsAdmin;
        public WokrGroupTable(bool check)
        {
            InitializeComponent();
            DG4.ItemsSource = Connect.context.WorkGroup.ToList();
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
            Nav.MainFrame.Navigate(new WokrGroupTableAdd((sender as Button).DataContext as WorkGroup));
        }

        private void AddBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new WokrGroupTableAdd(null));
        }

        private void DelBTN_Click(object sender, RoutedEventArgs e)
        {
            var delWorkGroup = DG4.SelectedItems.Cast<WorkGroup>().ToList();
            if (MessageBox.Show($"Удалить {delWorkGroup.Count} записей?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.WorkGroup.RemoveRange(delWorkGroup);
            try
            {
                Connect.context.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            DG4.ItemsSource = Connect.context.WorkGroup.ToList();
        }

        private void SearchBTN_Click(object sender, RoutedEventArgs e)
        {

            var poisk = Connect.context.WorkGroup.Where(x => x.Name_WorkGroup.ToString().StartsWith(SearchBox.Text)).ToList();
            DG4.ItemsSource = poisk;
        }

        private void UpdateBTN_Click(object sender, RoutedEventArgs e)
        {
            DG4.ItemsSource = Connect.context.WorkGroup.ToList();
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
            sheet.Cells[1, 1] = "ID Группы"; sheet.Cells[1, 2] = "Название группы";
            sheet.Cells[1, 3] = "Кол-во работников";
            var currentRow = 2;
            var prod = Connect.context.WorkGroup.ToList();
            foreach (var item in prod)
            {
                sheet.Cells[currentRow, 1] = item.ID_WorkGroup;
                sheet.Cells[currentRow, 2] = item.Name_WorkGroup;
                sheet.Cells[currentRow, 3] = item.Number_Of_Workers;
                currentRow++;
            }
            Excel.Range range2 = sheet.get_Range("A1", "F20"); range2.Cells.Font.Name = "Cascadia mono";
            range2.Cells.Font.Size = 10; range2.Font.Bold = false;
            range2.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
            range2.Borders.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
        }

        private void ReportPDF_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Document doc = new Document();
                PdfWriter.GetInstance(doc, new FileStream("Отчет по таблице Рабочие группы.pdf", FileMode.Create));
                doc.Open();
                BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                Font font = new Font(baseFont, Font.DEFAULTSIZE, Font.NORMAL);
                PdfPTable table = new PdfPTable(1);
                PdfPTable table2 = new PdfPTable(1);
                PdfPTable table3 = new PdfPTable(1);
                
                PdfPCell cell = new PdfPCell(new Phrase("Ведомость по рабочим группам", font));
                cell.Colspan = 5;
                cell.HorizontalAlignment = 1;
                cell.Border = 0;
                table.AddCell(cell);
                table2.AddCell(cell);
                table3.AddCell(cell);
               

                table.AddCell(new PdfPCell(new Phrase(new Phrase("ID рабочей группы", font))));
                table2.AddCell(new PdfPCell(new Phrase(new Phrase("Название рабочей группы", font))));
                table3.AddCell(new PdfPCell(new Phrase(new Phrase("Кол-во работников", font))));
              
                var a = 0;
                for (int i = 0; i < Connect.context.WorkGroup.ToList().Count; i++)
                {
                    var itemWorkGroup = Connect.context.WorkGroup.ToList()[i];
                    table.AddCell(new Phrase(itemWorkGroup.ID_WorkGroup.ToString(), font));
                    table2.AddCell(new Phrase(itemWorkGroup.Name_WorkGroup.ToString(), font));
                    table3.AddCell(new Phrase(itemWorkGroup.Number_Of_Workers.ToString(), font));
                   

                }
                doc.Add(table);
                doc.Add(table2);
                doc.Add(table3);
               
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
