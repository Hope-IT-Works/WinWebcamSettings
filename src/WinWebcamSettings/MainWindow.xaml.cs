using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using DirectShowLib;

namespace WinWebcamSettings
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int OleCreatePropertyFrame(
            IntPtr hwndOwner,
            int x,
            int y,
            [MarshalAs(UnmanagedType.LPWStr)] string lpszCaption,
            int cObjects,
            [MarshalAs(UnmanagedType.Interface, ArraySubType=UnmanagedType.IUnknown)]
            ref object ppUnk,
            int cPages,
            IntPtr lpPageClsID,
            int lcid,
            int dwReserved,
            IntPtr lpvReserved
        );
        public MainWindow()
        {
            InitializeComponent();
            RefreshList();
            ButtonRefresh.Click += ButtonRefresh_Click;
            ButtonSettings.Click += ButtonSettings_Click;
        }

        private List<string> GetAllConnectedCameras()
        {
            var cameraNames = new List<string>();
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE (PNPClass = 'Image' OR PNPClass = 'Camera')"))
            {
                foreach (var device in searcher.Get())
                {
                    string caption = device["Caption"]?.ToString() ?? "";
                    if (caption.Length > 0)
                    {
                        cameraNames.Add(caption);
                    }
                }
            }
            return cameraNames;
        }

        private void RefreshList()
        {
            var connectedCameras = GetAllConnectedCameras();
            WebcamList.ItemsSource = connectedCameras;
            if (connectedCameras.Count > 0)
            {
                WebcamList.SelectedIndex = 0;
            }
        }

        private IBaseFilter? CreateFilter(Guid category, string? friendlyname)
        {
            object? source = null;
            Guid iid = typeof(IBaseFilter).GUID;
            foreach (DsDevice device in DsDevice.GetDevicesOfCat(category))
            {
                if (device.Name.CompareTo(friendlyname) == 0)
                {
                    device.Mon.BindToObject(null, null, ref iid, out source);
                    break;
                }
            }

            return source as IBaseFilter;
        }

        private void OpenCameraProperties()
        {
            IBaseFilter? dev = CreateFilter(FilterCategory.VideoInputDevice, WebcamList.SelectedItem.ToString());
            if (dev == null)
            {
                return;
            }

            // Get the ISpecifyPropertyPages for the filter
            ISpecifyPropertyPages? pProp = dev as ISpecifyPropertyPages;
            if (pProp == null)
            {
                // If the filter doesn't implement ISpecifyPropertyPages, try displaying IAMVfwCompressDialogs instead
                IAMVfwCompressDialogs? compressDialog = dev as IAMVfwCompressDialogs;
                if (compressDialog != null)
                {
                    int hr = compressDialog.ShowDialog(VfwCompressDialogs.Config, IntPtr.Zero);
                    DsError.ThrowExceptionForHR(hr);
                }
                return;
            }

            try
            {
                // Get the name of the filter from the FilterInfo struct
                FilterInfo filterInfo;
                int hr = dev.QueryFilterInfo(out filterInfo);
                DsError.ThrowExceptionForHR(hr);

                // Get the property pages from the property bag
                DsCAUUID caGUID;
                hr = pProp.GetPages(out caGUID);
                DsError.ThrowExceptionForHR(hr);

                try
                {
                    // Create and display the OlePropertyFrame
                    object oDevice = (object)dev;
                    hr = OleCreatePropertyFrame(new WindowInteropHelper(this).Handle, 0, 0, filterInfo.achName, 1, ref oDevice, caGUID.cElems, caGUID.pElems, 0, 0, IntPtr.Zero);
                    DsError.ThrowExceptionForHR(hr);
                }
                finally
                {
                    // Release COM objects
                    Marshal.FreeCoTaskMem(caGUID.pElems);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(pProp);
                Marshal.ReleaseComObject(dev);
            }
        }

        private void ButtonRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshList();
        }

        private void ButtonSettings_Click(object sender, RoutedEventArgs e)
        {
            OpenCameraProperties();
        }
    }
}
