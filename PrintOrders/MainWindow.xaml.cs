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
using System.Management;

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
        private string printerName = "";
        private bool currentPrinterStatus = false;
        private PrintQueueMonitor pqm = null;
        private int IdCounter
        {
            get => (int)rk.GetValue("idCounter", 0);
            set => rk.SetValue("idCounter", value);
        }
        private SortedDictionary<int, JobPagesInfo> pausedJobs = null;
        private List<int> deletingJobs = null;
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
            // Запрещаем запускаться программе, если уже есть запущенный процесс.
            m = new Mutex(true, Process.GetCurrentProcess().ProcessName, out bool first);
            if (!first)
            {
                MessageBox.Show("PrintOrders.exe уже запущен!", "Внимание!", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                Close();
            }

            InitializeComponent();

            // Окно появляется только при созданном заказе, поэтому изначально оно скрыто.
            Hide();

            // Добавляем иконку в область уведомлений с возможностью выхода из программы.
            notifyIcon = new System.Windows.Forms.NotifyIcon
            {
                Icon = PrintOrders.Properties.Resources.printer,
                Visible = true,
            };
            System.Windows.Forms.ContextMenu niContextMenu = new System.Windows.Forms.ContextMenu();
            niContextMenu.MenuItems.Add("Выход", new EventHandler(Exit));
            notifyIcon.ContextMenu = niContextMenu;

            /*
             * Получаем из реестра подраздел HKEY_CURRENT_USER\Software\PrintOrders.
             * Если его не существует, то создаем новый.
             */
            rk = Registry.CurrentUser.OpenSubKey(@"Software\PrintOrders", true) ?? 
                Registry.CurrentUser.CreateSubKey(@"Software\PrintOrders");
            try
            {
                /*
                 * Если полученный из реестра день не будет совпадать с текущим днем, 
                 * то счетчик для номера заказа сбрасывается.
                 * Это может произойти, если включить компьютер на следующий день после выключения.
                 */
                if (DateTime.Now.ToShortDateString() != rk.GetValue("lastDate").ToString())
                    IdCounter = 0;
            }
            catch
            {
                /*
                 * Если в реестре не было найдено значение lastDate, 
                 * то создается новое значение с текущей датой.
                 * Происходит во время самого первого запуска программы.
                 */
                rk.SetValue("lastDate", DateTime.Now.ToShortDateString());
            }

            /* 
             * Создаем диспетчер данного приложения, 
             * чтобы с ним можно было взаимодействовать из обработчика заданий печати.
             */
            dispatcher = Application.Current.Dispatcher;

            /*
             * Создаем обработчик смены дня.
             * Как только наступает новый день, 
             * обработчик сбрасывается и в реестре обновляется значение lastDate.
             */
            MidnightNotifier.DayChanged += (s, e) => 
            {
                IdCounter = 0;
                rk.SetValue("lastDate", DateTime.Now.ToShortDateString());
            };

            // Создаем список, в котором будем хранить приостановленные задания печати.
            pausedJobs = new SortedDictionary<int, JobPagesInfo>();

            // Создаем список, в котором будем хранить удаляемые задания печати.
            deletingJobs = new List<int>();

            // Получаем ссылку на очередь печати по умолчанию.
            pq = LocalPrintServer.GetDefaultPrintQueue();
            printerName = pq.Name;

            // Проверяем статус принтера при запуске программы.
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            var unit = from ManagementObject x in searcher.Get()
                       where x.Properties["Name"].Value.ToString() == printerName
                       select x;
            currentPrinterStatus = bool.Parse(unit.First()["WorkOffline"].ToString());
            CheckPrinterStatus(currentPrinterStatus);

            // Инициализируем монитор печати и обработчик заданий печати.
            pqm = new PrintQueueMonitor(printerName);
            pqm.OnJobStatusChange += new PrintJobStatusChanged(Pqm_OnJobStatusChange);

            // Создаем обработчик изменения статуса принтера.
            string wmiQuery = "Select * From __InstanceModificationEvent Within 1 " +
            "Where TargetInstance ISA 'Win32_Printer' AND TargetInstance.Name ='" + printerName + "'";
            ManagementEventWatcher watcher = new ManagementEventWatcher(new ManagementScope("\\root\\CIMV2"), new EventQuery(wmiQuery));
            watcher.EventArrived += new EventArrivedEventHandler(WmiEventHandler);
            watcher.Start();
        }


        /*
         * Обработчик заданий печати.
         * В переменной e хранится информация о задании 
         * (идентификатор, имя, кол-во страниц (учитывая копии), кол-во копий, статус, объект PrintSystemJobInfo).
         */
        void Pqm_OnJobStatusChange(object Sender, PrintJobChangeEventArgs e)
        {
            /* 
             * Если последнее задание - это финальный файл orderInfo, 
             * то сбрасываем значения и скрываем окно.
             */
            if (e.JobName == "orderInfo")
                return;

            // Если задание в процессе удаления, то пропускаем его.
            if(deletingJobs.Contains(e.JobID))
            {
                if ((e.JobStatus & JOBSTATUS.JOB_STATUS_DELETED) == JOBSTATUS.JOB_STATUS_DELETED)
                {
                    deletingJobs.Remove(e.JobID);
                }
                return;
            }

            /*
             * Если задание было добавлено в очередь, когда принтер был выключен,
             * а затем принтер включили, то статус задания станет равным 0,
             * пропускаем его.
             */
            if (e.JobStatus == 0 && pausedJobs.ContainsKey(e.JobID))
                return;

            // Обработка удаляемого задания.
            if (((e.JobStatus & JOBSTATUS.JOB_STATUS_DELETING) == JOBSTATUS.JOB_STATUS_DELETING) && orderIsCreated)
            {
                /*
                 * Отнимаем кол-во страниц задания от общего кол-ва и
                 * удаляем задание из списка приостановленных заданий.
                 * Срабатывает, если вручную удалять задание из общего списка заданий.
                 */
                if (pausedJobs.ContainsKey(e.JobID))
                {
                    totalPages -= pausedJobs[e.JobID].TotalPages;
                    pausedJobs.Remove(e.JobID);
                    dispatcher.Invoke(() => { UpdateFilesList(); });
                }

                /*
                 * В случае, если общий список или список приостановленных заданий пуст,
                 * сбрасываем значения и скрываем окно.
                 */
                if (!LocalPrintServer.GetDefaultPrintQueue().GetPrintJobInfoCollection().Any())
                    pausedJobs.Clear();
                if (!pausedJobs.Any())
                    dispatcher.Invoke(() => { CloseOrder(); });
                return;
            }

            // Если задания больше нет в очереди печати, то пропускаем его.
            if (e.JobInfo == null)
                return;

            /*
             * Когда задание подсчитает страницы (задание перестанет "спулить"),
             * оно приостановится, добавится в список приостановленных заданий,
             * и кол-во его страниц добавится к общему кол-ву. 
             * Если это первое добавленное задание, то появится новое окно заказа, 
             * иначе - список файлов в существующем окне обновится.
             */
            if ((e.JobStatus & JOBSTATUS.JOB_STATUS_SPOOLING) == JOBSTATUS.JOB_STATUS_SPOOLING)
            {
                dispatcher.Invoke(() => { BlockButtons(); });
            }
            else
            {
                if(!pausedJobs.ContainsKey(e.JobID))
                {
                    e.JobInfo.Pause();
                    JobPagesInfo jobPagesInfo = new JobPagesInfo(e.JobInfo.NumberOfPages, e.JobCopies);
                    pausedJobs.Add(e.JobID, jobPagesInfo);
                    totalPages += jobPagesInfo.TotalPages;
                    if (!orderIsCreated)
                    {
                        orderIsCreated = true;
                        dispatcher.Invoke(() => { CreateOrder(); });
                    }
                    else
                    {
                        dispatcher.Invoke(() => { UpdateFilesList(); });
                    }
                    dispatcher.Invoke(() => { UnblockButtons(); });
                }
            }
        }       

        /*
         * Создание нового заказа.
         * Номер заказа создается по принципу
         * <буква, полученная из IP-адреса компьютера><случайная буква>-<последний октет IP-адреса><счетчик из реестра>
         * После этого обновляется список файлов.
         */
        private void CreateOrder()
        {
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

        /*
         * Обновление списка файлов в окне.
         * После вывода всей информации,
         * показывается окно если оно было скрыто или свернуто.
         */
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
            {
                Show();
                Activate();
            }
            if(WindowState == WindowState.Minimized)
                WindowState = WindowState.Normal;
        }

        /* 
         * Закрытие заказа.
         * Общее количество страниц становится равным 0, 
         * окно скрывается, и 
         * флаги orderIsPrinting и orderIsCreated становятся равными false.
         */
        void CloseOrder()
        {
            totalPages = 0;
            Hide();
            orderIsCreated = false;
        }

        /* 
         * При нажатии на кнопку "Печатать еще файл", 
         * окно сворачивается.
         */
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        /* 
         * При нажатии на кнопку "Отменить все", 
         * все задания из списка приостановленных заданий отменяются.
         */
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                deletingJobs.Add(job.Key);
                pq.GetJob(job.Key).Cancel();
            }
            pausedJobs.Clear();
            CloseOrder();
        }

        /*
         * При нажатии на кнопку "Готово", 
         * все задания из списка приостановленных заданий возобнавляются,
         * в папку C:\Users\<Username>\Documents\PrintOrders и
         * в очередь добавляются соответственно файлы 
         * orders_info_<дата> и orderInfo с информацией о задании
         * (дата, время, номер заказа, кол-во страниц),
         * и окно скрывается.
         */
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            /* 
             * Если на момент нажатия кнопки принтер не подключен,
             * то запрещаем пользователю закончить заказ и 
             * показываем окно с ошибкой. 
             */ 
            if (currentPrinterStatus)
            {
                MessageBox.Show("Принтер (" + printerName + ") не подключен!", "Внимание!",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (KeyValuePair<int, JobPagesInfo> job in pausedJobs)
            {
                deletingJobs.Add(job.Key);
                pq.GetJob(job.Key).Resume();
            }

            PrintDocument printDocument = new PrintDocument
            {
                DocumentName = "orderInfo",
                PrinterSettings = new PrinterSettings { Copies = 1 }
            };
            printDocument.PrintPage += new PrintPageEventHandler(PrintPageHandler);
            printDocument.Print();
            
            void PrintPageHandler(object s, PrintPageEventArgs j)
            {
                string pages = totalPages.ToString();
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
            pausedJobs.Clear();
            CloseOrder();
        }

        /* 
         * Закрытие приложения при нажатии на кнопку "Выход" 
         * в контекстном меню у иконки в трее.
         */
        private void Exit(object sender, EventArgs e)
        {
            Close();       
        }

        /*
         * При закрытии окна, иконка скрывается и 
         * в реестр записывается текущая дата.
         */
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            notifyIcon.Visible = false;
            rk.SetValue("lastDate", DateTime.Now.ToShortDateString());
        }

        // Блокировка кнопок.
        private void BlockButtons()
        {            
            addButton.IsEnabled = false;
            cancelButton.IsEnabled = false;
            okButton.IsEnabled = false;
        }

        // Разблокировка кнопок.
        private void UnblockButtons()
        {            
            addButton.IsEnabled = true;
            cancelButton.IsEnabled = true;
            okButton.IsEnabled = true;
        }

        /* 
         * Обработчик изменения статуса принтера.
         * Если состояние изменилось, 
         * то перезапускаем монитор печати и обновляем иконку в трее.
         */
        private void WmiEventHandler(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject printer = (ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value;
            bool newPrinterStatus = bool.Parse(printer["WorkOffline"].ToString());
            if (currentPrinterStatus != newPrinterStatus)
            {
                currentPrinterStatus = newPrinterStatus;
                pqm?.Stop();
                pqm = null;
                pqm = new PrintQueueMonitor(printerName);
                pqm.OnJobStatusChange += new PrintJobStatusChanged(Pqm_OnJobStatusChange);
                CheckPrinterStatus(newPrinterStatus);
            }
        }

        // Обновление иконки в трее в зависимости от статуса принтера.
        private void CheckPrinterStatus(bool printerIsOffline)
        {
            if (!printerIsOffline)
            {
                notifyIcon.Icon = PrintOrders.Properties.Resources.printer;
                notifyIcon.Text = "PrintOrders\nРаботает (" + printerName + ")";
            }
            else
            {
                notifyIcon.Icon = PrintOrders.Properties.Resources.printer_warning;
                notifyIcon.Text = "PrintOrders\nПринтер не подключен (" + printerName + ")";
            }
        }

        // Обработка смены дня.
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
