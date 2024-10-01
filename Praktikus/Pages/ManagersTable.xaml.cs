using Praktikus.Misc;
using Praktikus.Windows;
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
using System.Windows.Media.Animation;

namespace Praktikus.Pages
{
    /// <summary>
    /// Логика взаимодействия для ManagersTables.xaml
    /// </summary>
    public partial class ManagersTable : Page
    {
        public bool IsAdmin;
        public ManagersTable(bool check)
        {
            InitializeComponent();
            DG1.ItemsSource = Connect.context.Managers.ToList();
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

        private void DG2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Change_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new ManagersTableAdd((sender as Button).DataContext as Managers));
        }

        private void BackBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }

        private void AddBTN_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new ManagersTableAdd(null));
        }

        private void DelBTN_Click(object sender, RoutedEventArgs e)
        {
            var delManagers = DG1.SelectedItems.Cast<Managers>().ToList();
            if (MessageBox.Show($"Удалить {delManagers.Count} записей?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Connect.context.Managers.RemoveRange(delManagers);
            try
            {
                Connect.context.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            DG1.ItemsSource = Connect.context.Managers.ToList();
        }

        private void SearchBTN_Click(object sender, RoutedEventArgs e)
        {
            var poisk = Connect.context.Managers.Where(x => x.FIO.ToString().StartsWith(SearchBox.Text)).ToList();
            DG1.ItemsSource = poisk;
        }

        private void UpdateBTN_Click(object sender, RoutedEventArgs e)
        {
            DG1.ItemsSource = Connect.context.Managers.ToList();
            MessageBox.Show("Таблица обновлена!","Обновление", MessageBoxButton.OK, MessageBoxImage.Information);
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
            sheet.Cells[1, 1] = "ID Руководителя"; sheet.Cells[1, 2] = "ФИО";
            sheet.Cells[1, 3] = "Номер телефона"; sheet.Cells[1, 4] = "Электронная почта";
            var currentRow = 2;
            var prod = Connect.context.Managers.ToList();
            foreach (var item in prod)
            {
                sheet.Cells[currentRow, 1] = item.ID_Manager;
                sheet.Cells[currentRow, 2] = item.FIO;
                sheet.Cells[currentRow, 3] = item.Phone_Number;
                sheet.Cells[currentRow, 4] = item.Email_Adress;

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
                PdfWriter.GetInstance(doc, new FileStream("Отчет по таблице Руководители.pdf", FileMode.Create));
                doc.Open();
                BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                Font font = new Font(baseFont, Font.DEFAULTSIZE, Font.NORMAL);
                PdfPTable table = new PdfPTable(1);
                PdfPTable table2 = new PdfPTable(1);
                PdfPTable table3 = new PdfPTable(1);
                PdfPTable table4 = new PdfPTable(1);
                
                PdfPCell cell = new PdfPCell(new Phrase("Ведомость по руководителям", font));
                cell.Colspan = 5;
                cell.HorizontalAlignment = 1;
                cell.Border = 0;
                table.AddCell(cell);
                table2.AddCell(cell);
                table3.AddCell(cell);
                table4.AddCell(cell);
               

                table.AddCell(new PdfPCell(new Phrase(new Phrase("ID Руководителя", font))));
                table2.AddCell(new PdfPCell(new Phrase(new Phrase("ФИО", font))));
                table3.AddCell(new PdfPCell(new Phrase(new Phrase("Номер телефона", font))));
                table4.AddCell(new PdfPCell(new Phrase(new Phrase("Электронная почта", font))));
                var a = 0;
                for (int i = 0; i < Connect.context.Managers.ToList().Count; i++)
                {
                    var itemManager = Connect.context.Managers.ToList()[i];
                    table.AddCell(new Phrase(itemManager.ID_Manager.ToString(), font));
                    table2.AddCell(new Phrase(itemManager.FIO.ToString(), font));
                    table3.AddCell(new Phrase(itemManager.Phone_Number.ToString(), font));
                    table4.AddCell(new Phrase(itemManager.Email_Adress.ToString(), font));
                    

                }
                doc.Add(table);
                doc.Add(table2);
                doc.Add(table3);
                doc.Add(table4);
             
                doc.Close();
                MessageBox.Show("Pdf-документ сохранен");
            }
            catch
            {
                MessageBox.Show("Pdf-документ не сохранен", "Ошибка!");
            }
        }

        private void FilterBTN_Click(object sender, RoutedEventArgs e)
        {
            if (ChooseBox.SelectedIndex == 0)
            {
                var poisk = Connect.context.Managers.Where(x => x.Phone_Number.ToString().StartsWith(FilterBox.Text)).ToList();
                DG1.ItemsSource = poisk;
            }
            if (ChooseBox.SelectedIndex == 1)
            {
                var poisk = Connect.context.Managers.Where(x => x.ID_Manager.ToString().StartsWith(FilterBox.Text)).ToList();
                DG1.ItemsSource = poisk;
            }
            if (ChooseBox.SelectedIndex == 2)
            {
                var poisk = Connect.context.Managers.Where(x => x.FIO.ToString().StartsWith(FilterBox.Text)).ToList();
                DG1.ItemsSource = poisk;
            }
            if (ChooseBox.SelectedIndex == 3)
            {
                var poisk = Connect.context.Managers.Where(x => x.Email_Adress.ToString().StartsWith(FilterBox.Text)).ToList();
                DG1.ItemsSource = poisk;
            }
        }
    }
}
