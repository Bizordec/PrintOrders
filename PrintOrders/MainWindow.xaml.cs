using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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
using System.Printing;
using System.Timers;
using Microsoft.Win32;
using System.Windows.Threading;
using System.IO;
using System.Threading;
using System.Drawing.Printing;
using System.Drawing;
using System.Net;
using System.Net.Sockets;
using System.Data;
using Monitors;


namespace PrintOrdersGUI
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PrintQueueMonitor pqm = null;
        private int counter = 0;
        private List<int> pausedJobs = null;
        private bool createdOrder = false;
        private int fileCount = 1;
        private int totalPages = 0;
        private string orderId = "";

        public MainWindow()
        {
            InitializeComponent();
            Hide();
            MidnightNotifier.DayChanged += (s, e) => { counter = 0; };
            pausedJobs = new List<int>();
            PrinterSettings settings = new PrinterSettings();
            pqm = new PrintQueueMonitor(settings.PrinterName);
            pqm.OnJobStatusChange += new PrintJobStatusChanged(pqm_OnJobStatusChange);
        }
        
        void pqm_OnJobStatusChange(object Sender, PrintJobChangeEventArgs e)
        {     
            Application.Current.Dispatcher.Invoke(() => BlockButtons());
            if (e.JobStatus.HasFlag(PrintSpool.JOBSTATUS.JOB_STATUS_DELETING))
            {
                if (pausedJobs.Contains(e.JobID))
                {
                    pausedJobs.Remove(e.JobID);
                }
                if (createdOrder && pausedJobs.Count() == 0)
                {
                    createdOrder = false;
                    Application.Current.Dispatcher.Invoke(() => Hide());
                }
                return;
            }            
            if (e.JobStatus != 0 || e.JobName == "orderInfo" || e.JobInfo == null)
                return;
            if (pausedJobs.Contains(e.JobID))
            {
                pausedJobs.Remove(e.JobID);
                return;
            }                
            Application.Current.Dispatcher.Invoke(() => UnblockButtons());
            e.JobInfo.Pause();
            pausedJobs.Add(e.JobID);
            if (!IsVisible && !createdOrder)
            {
                createdOrder = true;

                string pages = e.JobInfo.NumberOfPages.ToString();

                IPHostEntry hostEntry = Dns.GetHostEntry(Dns.GetHostName());
                IPAddress[] addr = hostEntry.AddressList;
                var ip = addr.Where(x => x.AddressFamily == AddressFamily.InterNetwork).FirstOrDefault();
                int lastpart = int.Parse(ip.ToString().Split('.')[3]);
                Random rng = new Random();
                string alphabet = "АБВГДЕЖИКЛМНПРСТ";
                char ipLetter = alphabet[(lastpart + alphabet.Length) % alphabet.Length];
                char randLetter = alphabet[rng.Next(alphabet.Length)];
                string letters = ipLetter.ToString() + randLetter.ToString();
                orderId = letters + "-" + lastpart.ToString() + counter++.ToString();
                
                Application.Current.Dispatcher.Invoke(() => CreateOrder(new Job(pages, orderId, e.JobID)));
            }
            else
            {
                string pages = e.JobInfo.NumberOfPages.ToString();
                Application.Current.Dispatcher.Invoke(() => AddFile(new Job(pages, e.JobID)));
            }      
        }

        private void CreateOrder(Job j)
        {
            pagesLabel.Text = "";
            fileCount = 1;
            totalPages = 0;
            orderLabel.Content = "Номер: " + j.OrderId;
            AddFile(j);
        }

        private void AddFile(Job j)
        {
            LocalPrintServer printServer = new LocalPrintServer();
            PrintQueue pq = printServer.DefaultPrintQueue;
            if (pq.GetJob(j.JobId).Name == "orderInfo")
                return;
            totalPages += int.Parse(j.FilePages);
            pagesLabel.Text += "\r\n" + fileCount++.ToString() + "-й файл: " + j.FilePages + " стр.";
            totalLabel.Content = "Итого: " + totalPages + " стр.";
            if (!IsVisible)
            {
                Show();
            }            
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            LocalPrintServer printServer = new LocalPrintServer();
            PrintQueue pq = printServer.DefaultPrintQueue;
            foreach (int jobId in pausedJobs)
            {
                pq.GetJob(jobId).Resume();
            }

            // объект для печати
            PrintDocument printDocument = new PrintDocument
            {
                DocumentName = "orderInfo"
            };
            string pages = totalPages.ToString();
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
            }                      
                        
            Hide();
            createdOrder = false;
        }

        private void BlockButtons()
        {
            okButton.IsEnabled = false;
            addButton.IsEnabled = false;
        }

        private void UnblockButtons()
        {
            okButton.IsEnabled = true;
            addButton.IsEnabled = true;
        }

        static class MidnightNotifier
        {
            private static readonly System.Timers.Timer timer;
            public static event EventHandler<EventArgs> DayChanged;

            static MidnightNotifier()
            {
                timer = new System.Timers.Timer(GetSleepTime());
                timer.Elapsed += (s, e) =>
                {
                    OnDayChanged();
                    timer.Interval = GetSleepTime();
                };
                timer.Start();

                SystemEvents.TimeChanged += OnSystemTimeChanged;
            }

            private static double GetSleepTime()
            {
                var midnightTonight = DateTime.Today.AddDays(1);
                var differenceInMilliseconds = (midnightTonight - DateTime.Now).TotalMilliseconds;
                return differenceInMilliseconds;
            }

            private static void OnDayChanged()
            {
                DayChanged?.Invoke(null, null);
            }

            private static void OnSystemTimeChanged(object sender, EventArgs e)
            {
                timer.Interval = GetSleepTime();
            }
        }

        class Job
        {
            public string FilePages { get; private set; } = "";
            public string OrderId { get; private set; } = "";
            public int JobId { get; private set; } = 0;
            public Job(string pages, string orderId, int jobId)
            {
                FilePages = pages;
                OrderId = orderId;
                JobId = jobId;
            }
            public Job(string pages, int jobId)
            {
                FilePages = pages;
                JobId = jobId;
            }
        }
    }
}
