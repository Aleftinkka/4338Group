using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using Excel = Microsoft.Office.Interop.Excel;

namespace _4338Project
{
    /// <summary>
    /// Логика взаимодействия для Kuznetsova_4338.xaml
    /// </summary>
    public partial class Kuznetsova_4338 : System.Windows.Window
    {
        public ObservableCollection<isrpo_2> ServicesCollection { get; set; }

        public Kuznetsova_4338()
        {
            InitializeComponent();
            //    LoadServicesFromDatabase();
        }
        //private void LoadServicesFromDatabase()
        //{
        //    using (var context = new isrpo_2Entities1())
        //    {
        //        ServicesCollection = new ObservableCollection<isrpo_2Entities1>(
        //            context.isrpo_2.Select(s => new isrpo_2Entities1
        //            {
        //                Код_сотрудника = s.Код_сотрудника,
        //                Должность = s.Должность,
        //                Логин = s.Логин
        //            }).ToList()
        //        );

        //        isrpo_2.ItemsSource = ServicesCollection;
        //    }
        //}

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Прикиньте, меня зовут Альбина, мне 19 годиков!",
                  "Внимание");
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();


            using (isrpo_2Entities1 isrpo_2Entities = new isrpo_2Entities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    isrpo_2Entities.isrpo_2.Add(new isrpo_2()
                    {
                        Код_сотрудника = list[i, 0],
                        ФИО = list[i, 2],
                        Должность = list[i, 1],
                        Логин = list[i, 3],
                        Пароль = list[i, 4],
                        Последний_вход = list[i, 5],
                        Тип_входа = list[i, 6]
                    });
                }
                isrpo_2Entities.SaveChanges();
            }

        }
        List<isrpo_2> logins;
        List<isrpo_2> types;

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                DefaultExt = "*.xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Выберите место для сохранения файла"
            };

            if (sfd.ShowDialog() != true)
                return;

            using (var context = new isrpo_2Entities1())
            {
                types = context.isrpo_2.ToList().GroupBy(s => s.Тип_входа).Select(y => y.First()).ToList();
                logins = context.isrpo_2.ToList().OrderBy(s => s.Логин).ToList();

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = types.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < types.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet1 = app.Worksheets.Item[i + 1];

                    worksheet1.Name = Convert.ToString(types[i].Тип_входа);
                    worksheet1.Cells[1][2] = "Логин";
                    worksheet1.Cells[2][2] = "Должность";
                    worksheet1.Cells[3][2] = "Код сотрудника";
                    startRowIndex++;

                    foreach (var user in context.isrpo_2)
                    {

                        Excel.Range headerRange = worksheet1.Range[worksheet1.Cells[1][1], worksheet1.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = types[i].Тип_входа;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;

                        if (types[i].Тип_входа == user.Тип_входа)
                        {
                            worksheet1.Cells[1][startRowIndex] = user.Логин;
                            worksheet1.Cells[2][startRowIndex] = user.Должность;
                            worksheet1.Cells[3][startRowIndex] = user.Код_сотрудника;
                            startRowIndex++;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                app.Visible = true;
            }
        }
    }
}


