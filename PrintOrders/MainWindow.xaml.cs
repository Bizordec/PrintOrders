﻿using System;
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
using System.Diagnostics;
using System.Management;

namespace PrintOrdersGUI
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PrintQueueMonitor pqm = null;
        private int counter = 0;
        private SortedDictionary<int, JobPagesInfo> pausedJobs = null;
        private bool orderIsCreated = false;
        private int totalPages = 0;
        private string orderId = "";

        public MainWindow()
        {
            Mutex PJmutex = new Mutex(true, Process.GetCurrentProcess().ProcessName, out bool first);
            if (!first)
            {
                MessageBox.Show("PrintOrders уже запущен!", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Warning);
                Close();
            }
            InitializeComponent();
            Hide();
            MidnightNotifier.DayChanged += (s, e) => { counter = 0; };
            pausedJobs = new SortedDictionary<int, JobPagesInfo>();
            PrinterSettings settings = new PrinterSettings();
            pqm = new PrintQueueMonitor(settings.PrinterName);
            pqm.OnJobStatusChange += new PrintJobStatusChanged(pqm_OnJobStatusChange);
        }
        
        void pqm_OnJobStatusChange(object Sender, PrintJobChangeEventArgs e)
        {     
            Application.Current.Dispatcher.Invoke(() => BlockButtons());
            if (e.JobStatus.HasFlag(PrintSpool.JOBSTATUS.JOB_STATUS_DELETING))
            {
                if(pausedJobs.ContainsKey(e.JobID))
                {
                    totalPages -= pausedJobs[e.JobID].TotalPages;
                    pausedJobs.Remove(e.JobID);
                    Application.Current.Dispatcher.Invoke(() => UpdateFilesList());                    
                }
                Application.Current.Dispatcher.Invoke(() => UnblockButtons());
                PrintJobInfoCollection pjic = LocalPrintServer.GetDefaultPrintQueue().GetPrintJobInfoCollection();
                if (orderIsCreated && (!pausedJobs.Any() || !pjic.Any()))
                {
                    orderIsCreated = false;
                    pausedJobs.Clear();
                    totalPages = 0;
                    Application.Current.Dispatcher.Invoke(() => Hide());
                    return;
                }
                return;
            }
            Application.Current.Dispatcher.Invoke(() => UnblockButtons());
            if (e.JobStatus != 0 || e.JobName == "orderInfo" || e.JobInfo == null)
                return;
            if (pausedJobs.ContainsKey(e.JobID))
            {
                pausedJobs.Remove(e.JobID);
                return;
            }            
            e.JobInfo.Pause();
            
            JobPagesInfo job = new JobPagesInfo(e.JobInfo.NumberOfPages, e.JobCopies);
            pausedJobs.Add(e.JobID, job);
            totalPages += job.TotalPages;
            if (!orderIsCreated)
            {
                orderIsCreated = true;                
                Application.Current.Dispatcher.Invoke(() => CreateOrder());
            }
            else
            {
                Application.Current.Dispatcher.Invoke(() => UpdateFilesList());
            }      
        }

        private void CreateOrder()
        {
            pagesLabel.Text = "";

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
            orderLabel.Content = "Номер: " + orderId;
            UpdateFilesList();
        }

        private void UpdateFilesList()
        {
            int fileCount = 1;
            pagesLabel.Text = "";
            foreach (KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                pagesLabel.Text += "\r\n" + fileCount++ + "-й файл: ";
                pagesLabel.Text += job.Value.Pages != 0 ? job.Value.Pages + " стр.\t(копий: " + job.Value.Copies + ")" : "Н/Д";
            }
            totalLabel.Content = "Итого: " + totalPages + " стр.";
            if (!IsVisible)
                Show();
            if(WindowState == WindowState.Minimized)
                WindowState = WindowState.Normal;
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            PrintQueue pq = LocalPrintServer.GetDefaultPrintQueue();
            foreach (KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                pq.GetJob(job.Key).Cancel();
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PrintQueue pq = LocalPrintServer.GetDefaultPrintQueue();
            foreach(KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                pq.GetJob(job.Key).Resume();
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
            totalPages = 0;
            orderIsCreated = false;
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

        public class JobPagesInfo
        {
            public int Pages { get; private set; }
            public int Copies { get; private set; }
            public int TotalPages { get; private set; }
            public JobPagesInfo(int pages, short copies)
            {
                Pages = pages;
                Copies = copies;
                TotalPages = pages * copies;
            }
        }
    }
}
