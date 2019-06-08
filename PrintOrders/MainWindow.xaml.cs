using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Printing;
using Microsoft.Win32;
using System.Windows.Threading;
using System.IO;
using System.Threading;
using System.Drawing.Printing;
using System.Drawing;
using System.Net;
using System.Net.Sockets;
using System.Data;
using System.Diagnostics;
using PQM;
using PrintSpool;

namespace PrintOrdersGUI
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static Mutex m = null;
        private Dispatcher dispatcher = null;
        private PrintQueue pq = null;
        private PrintQueueMonitor pqm = null;
        private int IdCounter
        {
            get => (int)rk.GetValue("idCounter", 0);
            set => rk.SetValue("idCounter", value);
        }
        private SortedDictionary<int, JobPagesInfo> pausedJobs = null;
        private bool orderIsCreated = false;
        private int totalPages = 0;
        private string orderId = "";
        private System.Windows.Forms.NotifyIcon notifyIcon = null;
        private RegistryKey rk = null;

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

        public MainWindow()
        {
            m = new Mutex(true, Process.GetCurrentProcess().ProcessName, out bool first);
            if (!first)
            {
                MessageBox.Show("PrintOrders.exe уже запущен!", "Внимание!", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                Close();
            }

            InitializeComponent();
            Hide();

            notifyIcon = new System.Windows.Forms.NotifyIcon
            {
                Icon = PrintOrders.Properties.Resources.printer,
                Visible = true,
            };
            System.Windows.Forms.ContextMenu niContextMenu = new System.Windows.Forms.ContextMenu();
            niContextMenu.MenuItems.Add("Выход", new EventHandler(Exit));
            notifyIcon.ContextMenu = niContextMenu;

            rk = Registry.CurrentUser.OpenSubKey(@"Software\PrintOrders", true) ?? 
                Registry.CurrentUser.CreateSubKey(@"Software\PrintOrders");
            try
            {
                if (DateTime.Now.ToShortDateString() != rk.GetValue("dateOnShutdown").ToString())
                    IdCounter = 0;
            }
            catch
            {
                rk.SetValue("dateOnShutdown", DateTime.Now.ToShortDateString());
            }
            
            dispatcher = Application.Current.Dispatcher;
            MidnightNotifier.DayChanged += (s, e) => 
            {
                IdCounter = 0;
                rk.SetValue("dateOnShutdown", DateTime.Now.ToShortDateString());
            };
            pausedJobs = new SortedDictionary<int, JobPagesInfo>();
            rk.SetValue("dateOnShutdown", DateTime.Now.ToShortDateString());

            pq = LocalPrintServer.GetDefaultPrintQueue();
            notifyIcon.Text = "PrintOrders\nРаботает (" + pq.Name + ")";
            pqm = new PrintQueueMonitor(pq.Name);
            pqm.OnJobStatusChange += new PrintJobStatusChanged(Pqm_OnJobStatusChange);
        }
        
        void Pqm_OnJobStatusChange(object Sender, PrintJobChangeEventArgs e)
        {
            if (e.JobName == "orderInfo")
                return;            
            dispatcher.BeginInvoke(new Action(delegate { BlockButtons(); }));

            if ((e.JobStatus & JOBSTATUS.JOB_STATUS_DELETING) == JOBSTATUS.JOB_STATUS_DELETING)
            {
                if (pausedJobs.ContainsKey(e.JobID))
                {
                    totalPages -= pausedJobs[e.JobID].TotalPages;
                    pausedJobs.Remove(e.JobID);
                    if (orderIsCreated)
                        dispatcher.BeginInvoke(new Action(delegate { UpdateFilesList(); }));
                }
                dispatcher.BeginInvoke(new Action(delegate { UnblockButtons(); }));
                PrintJobInfoCollection pjic = LocalPrintServer.GetDefaultPrintQueue().GetPrintJobInfoCollection();
                if (orderIsCreated && (!pausedJobs.Any() || !pjic.Any()))
                {
                    orderIsCreated = false;
                    pausedJobs.Clear();
                    totalPages = 0;
                    dispatcher.BeginInvoke(new Action(delegate { Hide(); }));
                }
                return;
            }
            if((e.JobStatus & JOBSTATUS.JOB_STATUS_PAUSED) == JOBSTATUS.JOB_STATUS_PAUSED || 
                e.JobInfo == null)
            {
                dispatcher.BeginInvoke(new Action(delegate { UnblockButtons(); }));
                return;
            }
            if ((e.JobStatus & JOBSTATUS.JOB_STATUS_SPOOLING) != JOBSTATUS.JOB_STATUS_SPOOLING && 
                !pausedJobs.ContainsKey(e.JobID))
            {
                e.JobInfo.Pause();
                JobPagesInfo job = new JobPagesInfo(e.JobInfo.NumberOfPages, e.JobCopies);
                pausedJobs.Add(e.JobID, job);
                totalPages += job.TotalPages;
                if (!orderIsCreated)
                {
                    orderIsCreated = true;
                    dispatcher.BeginInvoke(new Action(delegate { CreateOrder(); }));
                }
                else
                {
                    dispatcher.BeginInvoke(new Action(delegate { UpdateFilesList(); }));
                }
            }
            dispatcher.BeginInvoke(new Action(delegate { UnblockButtons(); }));
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
            orderId = letters + "-" + lastpart.ToString() + IdCounter++.ToString();
            orderLabel.Content = "Номер: " + orderId;
            UpdateFilesList();
        }

        private void UpdateFilesList()
        {
            int fileCount = 1;
            pagesLabel.Text = "";
            foreach (KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                JobPagesInfo jobInfo = job.Value;
                pagesLabel.Text += "\r\n" + fileCount++ + "-й файл: ";
                pagesLabel.Text += jobInfo.Pages != 0 ? jobInfo.Pages + " стр.\t(копий: " + jobInfo.Copies + ")" : "Н/Д";
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
            foreach (KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                pq.GetJob(job.Key).Cancel();
            }
            pausedJobs.Clear();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            foreach(KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                pq.GetJob(job.Key).Resume();
            }
            
            PrintDocument printDocument = new PrintDocument
            {
                DocumentName = "orderInfo"
            };
            string pages = totalPages.ToString();

            printDocument.PrintPage += new PrintPageEventHandler(PrintPageHandler);
            printDocument.Print();
            
            void PrintPageHandler(object s, PrintPageEventArgs j)
            {
                string infoDirPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\PrintOrders\\";
                if(!Directory.Exists(infoDirPath))
                    Directory.CreateDirectory(infoDirPath);

                string infoFilePath = infoDirPath + "\\orders_info_" + DateTime.Now.ToShortDateString() + ".txt";
                FileInfo fi1 = new FileInfo(infoFilePath);
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
                    using (StreamWriter sw = File.AppendText(infoFilePath))
                    {
                        sw.WriteLine(row);
                    }
                }

                string result = "";
                result += "Дата: " + DateTime.Now.ToString() + "\n";
                result += "Номер заказа: " + orderId + "\n";
                result += "Количество страниц: " + pages;
                j.Graphics.DrawString(result, new Font("Arial", 14), Brushes.Black, 30, 30);
            }
            Hide();
            orderIsCreated = false;
        }

        private void Exit(object sender, EventArgs e)
        {
            Close();       
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            notifyIcon.Visible = false;
            rk.SetValue("dateOnShutdown", DateTime.Now.ToShortDateString());
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            Activate();
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
    }
}
