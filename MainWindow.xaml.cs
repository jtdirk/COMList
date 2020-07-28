using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Windows;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Media.Imaging;

namespace COMList
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IntPtr m_hNotifyDevNode;
        
        public MainWindow()
        {
            InitializeComponent();
            listCOMPorts();
        }

        private void listCOMPorts()
        {
            ListBoxCOMPorts.Items.Clear();
            List<string> tList = new List<string>();

            Thread thread = new Thread(() =>
                {
                    //using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
                    using (var searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\""))
                    {
                        string[] portnames = SerialPort.GetPortNames();

                        var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
                        //tList = (from n in portnames join p in ports on n equals p["DeviceID"].ToString() select n + " - " + p["Caption"]).ToList();
                        //tList = (from n in portnames join p in ports on n equals p["Caption"].ToString().Substring(p["Caption"].ToString().IndexOf("(COM") + 1, p["Caption"].ToString().LastIndexOf(")") - p["Caption"].ToString().IndexOf("(COM")) select n + " - " + p["Caption"]).ToList();
                        foreach(string port in portnames)
                        {
                            foreach(ManagementBaseObject p in ports)
                            {
                                string x = p["Caption"].ToString().Substring(p["Caption"].ToString().IndexOf("(COM") + 1, p["Caption"].ToString().LastIndexOf(")") - p["Caption"].ToString().IndexOf("(COM") - 1);
                                if (x.Equals(port))
                                {
                                    tList.Add(port + " - " + p["Caption"]);
                                }
                            }
                        }

                        tList.Sort(new Comparison<string>(delegate (string a, string b)
                        {
                            return int.Parse(a.Substring(3, a.IndexOf(' ') - 2)) -
                                   int.Parse(b.Substring(3, b.IndexOf(' ') - 2));
                        }));
                    }
                });

            thread.Start();
            thread.Join();

            if (tList.Count == 0)
            {
                tList.Add("Kein COM-Port vorhanden");
            }

            foreach (string s in tList)
            {
                if(ListBoxCOMPorts.Items.IndexOf(s) < 0)
                {
                    ListBoxCOMPorts.Items.Add(s);
                }
            }
        }

        private void RegisterNotification(Guid guid)
        {
            Dbt.DEV_BROADCAST_DEVICEINTERFACE devIF = new Dbt.DEV_BROADCAST_DEVICEINTERFACE();
            IntPtr devIFBuffer;

            // Set to HID GUID
            devIF.dbcc_size = Marshal.SizeOf(devIF);
            devIF.dbcc_devicetype = Dbt.DBT_DEVTYP_DEVICEINTERFACE;
            devIF.dbcc_reserved = 0;
            devIF.dbcc_classguid = guid;

            // Allocate a buffer for DLL call
            devIFBuffer = Marshal.AllocHGlobal(devIF.dbcc_size);

            // Copy devIF to buffer
            Marshal.StructureToPtr(devIF, devIFBuffer, true);

            // Register for HID device notifications
            m_hNotifyDevNode = Dbt.RegisterDeviceNotification((new WindowInteropHelper(this)).Handle, devIFBuffer, Dbt.DEVICE_NOTIFY_WINDOW_HANDLE);

            // Copy buffer to devIF
            Marshal.PtrToStructure(devIFBuffer, devIF);

            // Free buffer
            Marshal.FreeHGlobal(devIFBuffer);
        }

        // Unregister HID device notification
        private void UnregisterNotification()
        {
            uint ret = Dbt.UnregisterDeviceNotification(m_hNotifyDevNode);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(WndProc);

            Guid hidGuid = new Guid("4d1e55b2-f16f-11cf-88cb-001111000030");
            Guid usbXpressGuid = new Guid("3c5e1462-5695-4e18-876b-f3f3d08aaf18");
            Guid cp210xGuid = new Guid("993f7832-6e2d-4a0f-b272-e2c78e74f93e");
            Guid newCP210xGuid = new Guid("a2a39220-39f4-4b88-aecb-3d86a35dc748");
            Guid usbGuid1 = new Guid("A5DCBF10-6530-11D2-901F-00C04FB951ED");
            Guid usbGuid2 = new Guid("4d36e978-e325-11ce-bfc1-08002be10318");

            RegisterNotification(usbGuid1);
            RegisterNotification(usbGuid2);
        }

        protected IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // Intercept the WM_DEVICECHANGE message
            if (msg == Dbt.WM_DEVICECHANGE)
            {
                // Get the message event type
                int nEventType = wParam.ToInt32();

                // Check for devices being connected or disconnected
                if (nEventType == Dbt.DBT_DEVICEARRIVAL ||
                    nEventType == Dbt.DBT_DEVICEREMOVECOMPLETE)
                {
                    Dbt.DEV_BROADCAST_HDR hdr = new Dbt.DEV_BROADCAST_HDR();

                    // Convert lparam to DEV_BROADCAST_HDR structure
                    Marshal.PtrToStructure(lParam, hdr);

                    if (hdr.dbch_devicetype == Dbt.DBT_DEVTYP_DEVICEINTERFACE)
                    {
                        Dbt.DEV_BROADCAST_DEVICEINTERFACE_1 devIF = new Dbt.DEV_BROADCAST_DEVICEINTERFACE_1();

                        // Convert lparam to DEV_BROADCAST_DEVICEINTERFACE structure
                        Marshal.PtrToStructure(lParam, devIF);

                        // Get the device path from the broadcast message
                        string devicePath = new string(devIF.dbcc_name);

                        // Remove null-terminated data from the string
                        int pos = devicePath.IndexOf((char)0);
                        if (pos != -1)
                        {
                            devicePath = devicePath.Substring(0, pos);
                        }

                        // An HID device was connected or removed
                        if (nEventType == Dbt.DBT_DEVICEREMOVECOMPLETE || nEventType == Dbt.DBT_DEVICEARRIVAL)
                        {
                            //MessageBox.Show("Device \"" + devicePath + "\" was removed");
                            listCOMPorts();
                        }
                    }
                }
            }
            return IntPtr.Zero;
        }

        private void ButtonAlwaysOnTop_Checked(object sender, RoutedEventArgs e)
        {
            Window parent = Window.GetWindow(this);
            parent.Topmost = true;

            BitmapImage bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri("Resources/Pin.png", UriKind.Relative);
            bi.EndInit();
            ImagePin.Source = bi;
        }

        private void ButtonAlwaysOnTop_Unchecked(object sender, RoutedEventArgs e)
        {
            Window parent = Window.GetWindow(this);
            parent.Topmost = false;

            BitmapImage bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri("Resources/Unpin.png", UriKind.Relative);
            bi.EndInit();
            ImagePin.Source = bi;
        }
    }
}
