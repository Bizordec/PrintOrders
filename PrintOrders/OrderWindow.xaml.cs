using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
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
using System.Windows.Shapes;
using System.Windows.Threading;

namespace PrintOrdersGUI
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class OrderInfo : Window
    {
        private static OrderInfo instance;
        private int fileCount = 0;
        private int totalPages = 0;
        private string orderId = "";

        public OrderInfo(string orderId, int filePages)
        {
            InitializeComponent();
            orderLabel.Content += orderId;
            if (filePages == 0)
            {
                pagesLabel.Text += "Н/Д";
                totalLabel.Content += "Н/Д";
            }
            else
            {
                pagesLabel.Text += filePages.ToString() + " стр.";
                totalLabel.Content += filePages.ToString() + " стр.";
            }
            totalPages = filePages;
            fileCount++;
        }

        public static OrderInfo GetInstance(string orderId, int filePages)
        {
            if (instance == null)
                instance = new OrderInfo(orderId, filePages);
            return instance;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // объект для печати
            PrintDocument printDocument = new PrintDocument
            {
                DocumentName = "orderInfo" + DateTime.Now.ToLongTimeString()
            };
            string pages = totalPages.ToString();
            //string pages = jobs.Last().NumberOfPages.ToString();
            // обработчик события печати
            printDocument.PrintPage += new PrintPageEventHandler(PrintPageHandler);
            printDocument.Print();

            // обработчик события печати
            void PrintPageHandler(object s, PrintPageEventArgs j)
            {
                DirectoryInfo dir = Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\PrintOrders\\");
                string path = dir.FullName + "\\orders_info_" + DateTime.Now.ToShortDateString() + ".txt";
                FileInfo fi1 = new FileInfo(path);
                string row = DateTime.Now.ToShortDateString() + "\t" +
                    DateTime.Now.ToShortTimeString() + "\t" +
                    orderId + "\t\t" +
                    pages;
                if (!fi1.Exists)
                {
                    using (StreamWriter sw = fi1.CreateText())
                    {
                        sw.WriteLine("Дата\t\tВремя\tНомер заказа\tКол-во страниц");
                        sw.WriteLine(row);
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(row);
                    }
                }

                string result = "";
                // задаем текст для печати
                result += "Дата: " + DateTime.Now.ToString() + "\n";
                result += "Номер заказа: " + orderId + "\n";
                result += "Количество страниц: " + pages;
                // печать строки result
                j.Graphics.DrawString(result, new Font("Arial", 14), System.Drawing.Brushes.Black, 0, 0);
                j.HasMorePages = false;
            }
            instance = null;
            Close();
        }

        public void AddFile(int filePages)
        {
            fileCount++;
            totalPages += filePages;
            pagesLabel.Text += "\r\n" + fileCount.ToString() + "-й файл: " + filePages + " стр.";
            totalLabel.Content = "Итого: " + totalPages + " стр.";
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }
    }
}
