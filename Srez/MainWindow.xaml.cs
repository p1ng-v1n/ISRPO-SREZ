using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
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
using System.Text.Json;
using static Srez.Data;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Srez
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        List<Sale> sales = new List<Sale>();
        
        private  async void BtnGetData_Click(object sender, RoutedEventArgs e)
        {
            using (HttpClient httpClient = new HttpClient { BaseAddress = new Uri(Properties.Settings.Default.BaseAddress) })
            {
                
                var content = new StringContent("", Encoding.UTF8, "application/json");
                HttpResponseMessage httpResponseMessage = await httpClient.PostAsync($"/api/Sale?dateStart={Convert.ToDateTime(DpDateStart.SelectedDate).ToString("yyyy-MM-dd")}&dateEnd={Convert.ToDateTime( DpDateEnd.SelectedDate).ToString("yyyy-MM-dd")}",content);
                string data = httpResponseMessage.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                sales = JsonSerializer.Deserialize<List<Sale>>(data);
                DgSale.ItemsSource = sales;
                

            }
        }

   

       

        private void BtnCheque_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                var selectedClient = DgSale.SelectedItem as Sale;
                if (selectedClient != null)
                {
                    Word._Application wApp = new Word.Application();
                    Word._Document wDoc = wApp.Documents.Add();
                    wApp.Visible = false;
                    wDoc.Activate();
                    var ClientParagraph = wDoc.Content.Paragraphs.Add();
                    ClientParagraph.Range.Text = $"Фамилия:\t{selectedClient.client.lastName}\n" +
                        $"Имя:\t{selectedClient.client.firstName}\n" +
                        $"Отчество:\t{selectedClient.client.patronymic}\n" + $"Дата продажи\t{selectedClient.dateSale.ToString("yyyy-MM-dd")}";
                    Word.Table wTable = wDoc.Tables.Add((Word.Range)ClientParagraph.Range,
                        selectedClient.telephones.Length + 1, 6, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
                    wTable.Cell(1, 1).Range.Text = "Артикул";
                    wTable.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 2).Range.Text = "Наименование";
                    wTable.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 3).Range.Text = "Категория";
                    wTable.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 4).Range.Text = "Количество";
                    wTable.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 5).Range.Text = "Стоимость";
                    wTable.Cell(1, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 6).Range.Text = "Cумма";
                    wTable.Cell(1, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int countRow = 2;
                    foreach (var item in selectedClient.telephones)
                    {
                        wTable.Cell(countRow, 1).Range.Text = item.articul.ToString();
                        wTable.Cell(countRow, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 2).Range.Text = item.nameTelephone.ToString();
                        wTable.Cell(countRow, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 3).Range.Text = item.category.ToString();
                        wTable.Cell(countRow, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 4).Range.Text = item.count.ToString();
                        wTable.Cell(countRow, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 5).Range.Text = item.cost.ToString();
                        wTable.Cell(countRow, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        double fullpice = item.cost * item.count;
                        wTable.Cell(countRow, 6).Range.Text = fullpice.ToString();
                        wTable.Cell(countRow, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        countRow++;
                    }
                    wDoc.SaveAs2($@"{Environment.CurrentDirectory}\1.docx");
                    wDoc.SaveAs2($@"{Environment.CurrentDirectory}\1.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                    wDoc.Close();
                    MessageBox.Show("Чек сформирован");
                }
               
            }
            catch (Exception ex)
            {

                MessageBox.Show($"Ошибка, выберите клиента");
            }
        }

        private void BtnChequeExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedClient = DgSale.SelectedItem as Sale;
                if (selectedClient != null)
                {
                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelApp.SheetsInNewWorkbook = 2;
                    Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
                    excelApp.DisplayAlerts = false;
                    Excel.Worksheet sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
                    sheet.Name = "Чек";
                    sheet.Columns[2].ColumnWidth = 15;
                    sheet.Cells[1, 1] = "Артикул";
                    sheet.Cells[1, 2] = "Наименование";
                    sheet.Cells[1, 3] = "Категория";
                    sheet.Cells[1, 4] = "Количество";
                    sheet.Cells[1, 5] = "Цена";
                    sheet.Cells[1, 6] = "Манафактура";
                    int countrow = 2;
                    foreach (var item in selectedClient.telephones)
                    {
                        sheet.Cells[countrow, 1] = item.articul;
                        sheet.Cells[countrow, 2] = item.nameTelephone;
                        sheet.Cells[countrow, 3] = item.category;
                        sheet.Cells[countrow, 4] = item.count;
                        sheet.Cells[countrow, 5] = item.cost;
                        sheet.Cells[countrow, 6] = item.manufacturer;
                        sheet.Cells[countrow, 7].Formula = $"=D{countrow}*E{countrow}";
                        countrow++;
                    }
                    sheet.Cells[1, 8] = "Итого";
                    sheet.Cells[2, 8].Formula = $"=SUM(G2:G{countrow - 1} )";
                    sheet.SaveAs($@"{ Environment.CurrentDirectory}\1.xlsx");
                    MessageBox.Show("Файл сохранен!");
                    excelApp.Quit();
                }
                  
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка!");
            }
        }

        private void BtnReportWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Word._Application wApp = new Word.Application();
                Word._Document wDoc = wApp.Documents.Add();
                wApp.Visible = false;
                wDoc.Activate();
                foreach (var item in sales)
                {
                    float cost = 0;
                    var ClientParagraph = wDoc.Content.Paragraphs.Add();
                    ClientParagraph.Range.Text = $"Фамилия:\t{item.client.lastName}\n" +
                        $"Имя:\t{item.client.firstName}\n" +
                        $"Отчество:\t{item.client.patronymic}\n";
                    Word.Table wTable = wDoc.Tables.Add((Word.Range)ClientParagraph.Range,
                        item.telephones.Length + 1, 6, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
                    wTable.Cell(1, 1).Range.Text = "Артикул";
                    wTable.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 2).Range.Text = "Наименование";
                    wTable.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 3).Range.Text = "Категория";
                    wTable.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 4).Range.Text = "Количество";
                    wTable.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 5).Range.Text = "Цена";
                    wTable.Cell(1, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 6).Range.Text = "Производитель";
                    wTable.Cell(1, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int countRow = 2;
                    foreach (var telephone in item.telephones)
                    {
                        wTable.Cell(countRow, 1).Range.Text = telephone.articul.ToString();
                        wTable.Cell(countRow, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 2).Range.Text = telephone.nameTelephone.ToString();
                        wTable.Cell(countRow, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 3).Range.Text = telephone.category.ToString();
                        wTable.Cell(countRow, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 4).Range.Text = telephone.count.ToString();
                        wTable.Cell(countRow, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 5).Range.Text = telephone.cost.ToString();
                        wTable.Cell(countRow, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 6).Range.Text = telephone.manufacturer.ToString();
                        wTable.Cell(countRow, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cost += telephone.cost * telephone.count;
                        countRow++;
                    }
                    var CostParagraph = wDoc.Content.Paragraphs.Add();
                    CostParagraph.Range.Text = $"Стоимость:\t{cost}\n";
                }
                wDoc.SaveAs2($@"{Environment.CurrentDirectory}\2.docx");
                wDoc.SaveAs2($@"{Environment.CurrentDirectory}\2.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                MessageBox.Show("Файл сформирован");
                wDoc.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show($"Ошибка!");
            }

        }

        private void CbiBrand_Selected(object sender, RoutedEventArgs e)
        {

            List<string> name = new List<string>();
            List<double> countsale = new List<double>();
            int count = 0;
            foreach (var item in sales)
            {

                foreach (var telephone in item.telephones)
                {
                    if (!name.Contains(telephone.manufacturer))
                    {
                        name.Add(telephone.manufacturer);
                        countsale.Add(telephone.count * telephone.cost);
                    }
                    else
                    {
                        var x = countsale.ElementAt(name.IndexOf(telephone.manufacturer));
                        x += telephone.count * telephone.cost;
                    }
                }

            }
            var pie = PieChart.Plot.AddPie(countsale.ToArray());
            pie.SliceLabels = name.ToArray();
            pie.ShowPercentages = true;
            pie.ShowValues = true;
            pie.ShowLabels = true;
            PieChart.Plot.Legend();
            PieChart.Refresh();
            GridChartLine.Visibility = Visibility.Collapsed;
            GridChartPie.Visibility = Visibility.Visible;

        }

        private void BtnReportExcel_Click(object sender, RoutedEventArgs e)
        {


            DateTime? dateStart = DpDateStart.SelectedDate;
            if (dateStart == null)
            {
                return;
            }
            DateTime? dateEnd = DpDateEnd.SelectedDate;
            if (dateEnd == null)
            {
                return;
            }
            if (sales.Count == 0 || sales == null)
            {
                return;
            }
            Excel.Application excelApp = new Excel.Application();
            try
            {
                excelApp.Visible = false;
                excelApp.SheetsInNewWorkbook = 2;
                Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
                excelApp.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
                sheet.Name = "Чек";
                sheet.Columns[2].ColumnWidth = 15;
                sheet.Cells[1, 1] = $"Отчет по продажам за период от {dateStart?.ToString("d")} до {dateEnd?.ToString("d")}";
                sheet.Cells[3, 1] = "Дата продажи";
                sheet.Cells[3, 2] = "Клиент";
                sheet.Cells[3, 3] = "Артикул";
                sheet.Cells[3, 4] = "Количество";
                sheet.Cells[3, 5] = "Цена";
                sheet.Cells[3, 6] = "Сумма";
                int countrow = 4;
                foreach (var item in sales)
                {
                    sheet.Cells[countrow, 1] = item.dateSale;
                    sheet.Cells[countrow, 2] = item.client.fullname;
                    foreach (var item2 in item.telephones)
                    {
                        
                        
                        
                        sheet.Cells[countrow, 3] = item2.articul;
                        sheet.Cells[countrow, 4] = item2.count;
                        sheet.Cells[countrow, 5] = item2.cost;
                        sheet.Cells[countrow, 6] = item2.cost * item2.count;
                        countrow++;
                    }
                }
                sheet.Cells[countrow, 4] = "Итого";
                sheet.Cells[countrow, 5].Formula = $"=SUM(F3:F{countrow - 1} )";
                sheet.SaveAs($@"{ Environment.CurrentDirectory}\12.xlsx");
                MessageBox.Show("Файл сохранен!");
                excelApp.Quit();
            }
            catch
            {
                excelApp.Quit();
            }
        }

        private void CbiSales_Selected(object sender, RoutedEventArgs e)
        {
            int count = (DpDateEnd.SelectedDate.Value.Date - DpDateStart.SelectedDate.Value.Date).Days+1;
            double[] countDate = new double[count];
            DateTime[] dates = new DateTime[count];
            dates[0] = DpDateStart.SelectedDate.Value.Date;

            for (int i = 0; i < count; i++)
            {
                dates[i] = dates[0].AddDays(i);
                countDate[i] = sales.Where(c=>c.dateSale==dates[i]).Sum(x => x.telephones.Sum(c => c.count * c.cost));
            }
            double[] xs = dates.Select(x => x.ToOADate()).ToArray();
            LineChart.Plot.AddScatter(xs, countDate);
            GridChartLine.Visibility = Visibility.Visible;
            GridChartPie.Visibility = Visibility.Collapsed;
            LineChart.Plot.XAxis.DateTimeFormat(true);
    
            LineChart.Refresh();
        }
    }
}

