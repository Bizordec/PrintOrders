using System;
using System.Collections.Generic;
using System.Linq;
using System.Printing;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using PrintSpool;


namespace Monitors
{

    public class PrintJobChangeEventArgs : EventArgs
    {
        public int JobID { get; } = 0;
        public string JobName { get; } = "";
        public JOBSTATUS JobStatus { get; } = new JOBSTATUS();
        public PrintSystemJobInfo JobInfo { get; } = null;
        public uint JobTotalPages { get; set; } = 0;
        public short JobCopies { get; set; } = 0;
        public PrintJobChangeEventArgs(int intJobID, string strJobName, JOBSTATUS jStatus, PrintSystemJobInfo objJobInfo, uint uintJobTotalPages, short shortJobCopies)
            : base()
        {
            JobID = intJobID;
            JobName = strJobName;
            JobStatus = jStatus;
            JobInfo = objJobInfo;
            JobTotalPages = uintJobTotalPages;
            JobCopies = shortJobCopies;
        }
    }

    public delegate void PrintJobStatusChanged(object Sender, PrintJobChangeEventArgs e);

    public class PrintQueueMonitor  
    {
        #region DLL Import Functions
        [DllImport("winspool.drv", EntryPoint = "OpenPrinterA",SetLastError = true, CharSet = CharSet.Ansi,ExactSpelling = true,CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter(String pPrinterName,
        out IntPtr phPrinter,
        Int32 pDefault);


        [DllImport("winspool.drv", EntryPoint = "ClosePrinter",
            SetLastError = true,
            ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter
        (Int32 hPrinter);

        [DllImport("winspool.drv",
        EntryPoint = "FindFirstPrinterChangeNotification",
        SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = true,
        CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr FindFirstPrinterChangeNotification
                            ([InAttribute()] IntPtr hPrinter,
                            [InAttribute()] Int32 fwFlags,
                            [InAttribute()] Int32 fwOptions,
                            [InAttribute(), MarshalAs(UnmanagedType.LPStruct)] PRINTER_NOTIFY_OPTIONS pPrinterNotifyOptions);

        [DllImport("winspool.drv", EntryPoint = "FindNextPrinterChangeNotification",
        SetLastError = true, CharSet = CharSet.Ansi,
        ExactSpelling = false,
        CallingConvention = CallingConvention.StdCall)]
        public static extern bool FindNextPrinterChangeNotification
                            ([InAttribute()] IntPtr hChangeObject,
                             [OutAttribute()] out Int32 pdwChange,
                             [InAttribute(), MarshalAs(UnmanagedType.LPStruct)] PRINTER_NOTIFY_OPTIONS pPrinterNotifyOptions,
                            [OutAttribute()] out IntPtr lppPrinterNotifyInfo
                                 );

        [DllImport(
        "winspool.drv",
        EntryPoint = "GetJob",
        SetLastError = true,
        ExactSpelling = false,
        CallingConvention = CallingConvention.StdCall)]
        private static extern bool GetJob
        ([InAttribute()] IntPtr hPrinter,
        [InAttribute()] UInt32 JobId,
        [InAttribute()] UInt32 Level,
        [OutAttribute()] IntPtr pJob,
        [InAttribute()] UInt32 cbBuf,
        [OutAttribute()] out UInt32 pcbNeeded);        
        #endregion

        #region Constants
        const int PRINTER_NOTIFY_OPTIONS_REFRESH = 1;
        #endregion
        #region Events
        public event PrintJobStatusChanged OnJobStatusChange;
        #endregion

        #region private variables
        private IntPtr _printerHandle = IntPtr.Zero;
        private string _spoolerName = "";
        private ManualResetEvent _mrEvent = new ManualResetEvent(false);
        private RegisteredWaitHandle _waitHandle = null;
        private IntPtr _changeHandle = IntPtr.Zero;
        private PRINTER_NOTIFY_OPTIONS _notifyOptions = new PRINTER_NOTIFY_OPTIONS();
        private Dictionary<int, string> objJobDict = new Dictionary<int, string>();
        private PrintQueue _spooler = null;
        #endregion

        #region constructor

        public PrintQueueMonitor(string strSpoolName) 
        {
            // Let us open the printer and get the printer handle.
            _spoolerName = strSpoolName;
            //Start Monitoring
            Start();

        }
        #endregion

        #region destructor
        ~PrintQueueMonitor() 
        {
            Stop();
        }
        #endregion

        #region StartMonitoring
        public void Start()
        {
            OpenPrinter(_spoolerName, out _printerHandle, 0);
            if (_printerHandle != IntPtr.Zero)
            {
                //We got a valid Printer handle.  Let us register for change notification....
                _changeHandle = FindFirstPrinterChangeNotification(_printerHandle, (int)PRINTER_CHANGES.PRINTER_CHANGE_JOB, 0, _notifyOptions);
                // We have successfully registered for change notification.  Let us capture the handle...
                _mrEvent.SafeWaitHandle = new Microsoft.Win32.SafeHandles.SafeWaitHandle(_changeHandle, true);
                //Now, let us wait for change notification from the printer queue....
                _waitHandle = ThreadPool.RegisterWaitForSingleObject(_mrEvent, new WaitOrTimerCallback(PrinterNotifyWaitCallback), _mrEvent, -1, true);
            }

            _spooler = new PrintQueue(new PrintServer(), _spoolerName);
            foreach (PrintSystemJobInfo psi in _spooler.GetPrintJobInfoCollection())
            {
                objJobDict[psi.JobIdentifier] = psi.Name;
            }
        }
        #endregion

        #region StopMonitoring
        public void Stop()
        {
            if (_printerHandle != IntPtr.Zero)
            {
                ClosePrinter((int)_printerHandle);
                _printerHandle = IntPtr.Zero;
            }
        }
        #endregion

        private const int ERROR_INSUFFICIENT_BUFFER = 122;
        #region Callback Function
        public void PrinterNotifyWaitCallback(Object state,bool timedOut)
        {
            if (_printerHandle == IntPtr.Zero) return;
            #region read notification details
            _notifyOptions.Count = 1;
            int pdwChange = 0;
            IntPtr pNotifyInfo = IntPtr.Zero;
            bool bResult = FindNextPrinterChangeNotification(_changeHandle, out pdwChange, _notifyOptions, out pNotifyInfo);
            //If the Printer Change Notification Call did not give data, exit code
            if ((bResult == false) || (((int)pNotifyInfo) == 0)) return;

            //If the Change Notification was not relgated to job, exit code
            bool bJobRelatedChange = ((pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_ADD_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_ADD_JOB) || 
                                     ((pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_SET_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_SET_JOB) ||
                                     ((pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_DELETE_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_DELETE_JOB) ||
                                     ((pdwChange & PRINTER_CHANGES.PRINTER_CHANGE_WRITE_JOB) == PRINTER_CHANGES.PRINTER_CHANGE_WRITE_JOB);
            if (!bJobRelatedChange) return;
            #endregion 

            #region populate Notification Information
            //Now, let us initialize and populate the Notify Info data            
            PRINTER_NOTIFY_INFO info = (PRINTER_NOTIFY_INFO)Marshal.PtrToStructure(pNotifyInfo, typeof(PRINTER_NOTIFY_INFO));
            long pData = (int)pNotifyInfo + (long)Marshal.OffsetOf(typeof(PRINTER_NOTIFY_INFO), "aData");
            PRINTER_NOTIFY_INFO_DATA[] data = new PRINTER_NOTIFY_INFO_DATA[info.Count];
            for (uint i = 0; i < info.Count; i++)
            {
                data[i] = (PRINTER_NOTIFY_INFO_DATA)Marshal.PtrToStructure((IntPtr)pData, typeof(PRINTER_NOTIFY_INFO_DATA));
                pData += Marshal.SizeOf(typeof(PRINTER_NOTIFY_INFO_DATA));
            }
            #endregion

            #region iterate through all elements in the data array
            for (int i = 0; i < data.Count(); i++)
            {

                if ((data[i].Field == (ushort)PRINTERJOBNOTIFICATIONTYPES.JOB_NOTIFY_FIELD_STATUS) && 
                    (data[i].Type == (ushort)PRINTERNOTIFICATIONTYPES.JOB_NOTIFY_TYPE))
                {                    
                    JOBSTATUS jStatus  = (JOBSTATUS)Enum.Parse(typeof(JOBSTATUS), data[i].NotifyData.Data.cbBuf.ToString());
                    int intJobID = (int)data[i].Id;
                    string strJobName = ""; 
                    PrintSystemJobInfo pji = null;
                    try
                    {
                        _spooler = new PrintQueue(new PrintServer(), _spoolerName);
                        pji = _spooler.GetJob(intJobID);                       
                        if (!objJobDict.ContainsKey(intJobID))
                            objJobDict[intJobID] = pji.Name;
                        strJobName = pji.Name;
                    }
                    catch
                    {
                        pji = null;
                        objJobDict.TryGetValue(intJobID, out strJobName);
                        if (strJobName == null) strJobName = "";
                    }
                    bool result;
                    short shortJobCopies = 0;
                    uint uintJobTotalPages = 0;
                    result = GetJob(_printerHandle, (uint)intJobID, 2, IntPtr.Zero, 0, out uint needed);
                    if (Marshal.GetLastWin32Error() == ERROR_INSUFFICIENT_BUFFER)
                    {
                        IntPtr buffer = Marshal.AllocHGlobal((int)needed);
                        result = GetJob(_printerHandle, (uint)intJobID, 2, buffer, needed, out needed);
                        JOB_INFO_2 jobInfo = (JOB_INFO_2)Marshal.PtrToStructure(buffer, typeof(JOB_INFO_2));
                        DEVMODE dMode = (DEVMODE)Marshal.PtrToStructure(jobInfo.pDevMode, typeof(DEVMODE));
                        shortJobCopies = dMode.dmCopies;
                        uintJobTotalPages = jobInfo.TotalPages;
                        Marshal.FreeHGlobal(buffer);
                    }
                    //Let us raise the event
                    OnJobStatusChange?.Invoke(this, new PrintJobChangeEventArgs(intJobID, strJobName, jStatus, pji, uintJobTotalPages, shortJobCopies));                    
                }
            }
            #endregion

            #region reset the Event and wait for the next event
            _mrEvent.Reset();
            _waitHandle = ThreadPool.RegisterWaitForSingleObject(_mrEvent, new WaitOrTimerCallback(PrinterNotifyWaitCallback), _mrEvent, -1, true);
            #endregion
        }
        #endregion
    }

}
