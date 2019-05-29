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
        PrintQueueMonitor pqm = null;
        int counter = 0;
        List<int> pausedJobs = null;
        private OrderInfo orderWindow = null;
        private bool created = false;

        public MainWindow()
        {
            InitializeComponent();
            MidnightNotifier.DayChanged += (s, e) => { counter = 0; };
            pausedJobs = new List<int>();
            PrinterSettings settings = new PrinterSettings();
            pqm = new PrintQueueMonitor(settings.PrinterName);
            pqm.OnJobStatusChange += new PrintJobStatusChanged(pqm_OnJobStatusChange);            
        }

        void pqm_OnJobStatusChange(object Sender, PrintJobChangeEventArgs e)
        {
            if (e.JobStatus != 0)
                return;
            e.JobInfo.Pause();
            pausedJobs.Add(e.JobID);
            if (e.JobName.Contains("orderInfo"))
            {
                LocalPrintServer printServer = new LocalPrintServer();
                PrintQueue pq = printServer.DefaultPrintQueue;
                foreach (int jobId in pausedJobs)
                {
                    pq.GetJob(jobId).Resume();
                }
                pausedJobs.Clear();
                //orderWindow = null;
            }
            else if (orderWindow == null && !e.JobName.Contains("orderInfo"))
            {
                Job job = new Job(e.JobInfo.NumberOfPages);
                Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() => CreateOrder(job)));
                created = true;
            }
            else if (created && !e.JobName.Contains("orderInfo"))
            {
                Job job = new Job(e.JobInfo.NumberOfPages);
                Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, 
                    new Action(() => AddFile(job)));
            }
        }

        private void CreateOrder(Job job)
        {
            Thread th = Thread.CurrentThread;
            IPHostEntry hostEntry = Dns.GetHostEntry(Dns.GetHostName());
            IPAddress[] addr = hostEntry.AddressList;
            var ip = addr.Where(x => x.AddressFamily == AddressFamily.InterNetwork).FirstOrDefault();
            int lastpart = int.Parse(ip.ToString().Split('.')[3]);
            Random rng = new Random();
            string alphabet = "АБВГДЕЖИКЛМНПРСТ";
            char ipLetter = alphabet[(lastpart + alphabet.Length) % alphabet.Length];
            char randLetter = alphabet[rng.Next(alphabet.Length)];
            string letters = ipLetter.ToString() + randLetter.ToString();
            string orderId = letters + "-" + lastpart.ToString() + counter++.ToString();

            //orderWindow = OrderInfo.GetInstance(orderId, job.FilePages);
            orderWindow = new OrderInfo(orderId, job.FilePages);
            orderWindow.Show();
        }

        private void AddFile(Job job)
        {
            if(orderWindow != null)
            {
                orderWindow.AddFile(job.FilePages);
                if(!orderWindow.IsVisible)
                {
                    orderWindow.Show();
                }
            }
            
        }

        private void BlockButtons()
        {
            orderWindow.okButton.IsEnabled = false;
            orderWindow.addButton.IsEnabled = false;
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
            public int FilePages { get; private set; }
            public Job(int pages)
            {
                FilePages = pages;
            }
        }
    }
}
