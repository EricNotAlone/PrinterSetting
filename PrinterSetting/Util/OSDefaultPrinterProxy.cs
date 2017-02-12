using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Runtime.InteropServices;


namespace PrinterSetting.Util
{
    public class OSDefaultPrinterProxy
    {

        private const long HWND_BROADCAST = 0xffffL;
        private const long WM_WININICHANGE = 0x1a;

        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        private static extern int SendMessage(long hwnd, long wMsg, long wParam, string lParam);
        
        [DllImport("kernel32.dll")]
        static extern bool WriteProfileString(string lpAppName, string lpKeyName, string lpString);


        /// <summary>
        /// 設定作業系統預設印表機。True:代表更新成功，False:代表更新失敗。
        /// </summary>
        /// <param name="SelectedPrinterName">印表機名稱</param>
        public bool SetDefaultPrinter(string SelectedPrinterName)
        {
            try
            {
                // 使用 WriteProfileString 設定預設印表機
                WriteProfileString("windows", "Device", (SelectedPrinterName + ",,"));

                // 使用 SendMessage 傳送正確的通知給所有最上層的層級視窗。
                // WIN.INI 要在意的應用程式接聽此訊息，並且視需要重新讀取 WIN.ini
                SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows");

                return true;
            }
            catch
            {
                return false;       
            }
        }

        /// <summary>
        /// 取得作業系統目前的預設印表機名稱。True:代表有找到，False:代表沒有找到，回傳值沒有意義。
        /// </summary>
        /// <param name="OSDefaultPrinter">回傳的作業系統目前的預設印表機名稱</param>
        public bool GetOSDefaultPrinter(out string OSDefaultPrinter)
        {
            try
            {
                using (PrintDocument pd = new PrintDocument())
                {

                    OSDefaultPrinter = pd.PrinterSettings.PrinterName;

                    return true;
                }

            }
            catch
            {
                OSDefaultPrinter = string.Empty;
                return false;
            }
        }

        /// <summary>
        /// 取得作業系統所有已安裝的印表機名稱。True:代表有找到，False:代表沒有找到，回傳值沒有意義。
        /// </summary>
        /// <param name="OSPrinterList">回傳的印表機清單</param>
        public bool GetOSPrinterList(out List<string> OSPrinterList)
        {
            try
            {
                using (PrintDocument printDoc = new PrintDocument())
                {
                    string OSdefaultPrinter;


                    OSPrinterList = new List<string>();
                    OSdefaultPrinter = printDoc.PrinterSettings.PrinterName;


                    foreach (string printerName in PrinterSettings.InstalledPrinters)
                    {
                        OSPrinterList.Add(printerName);
                    }

                    return true;
                }

            }
            catch(Exception EX)
            {
                OSPrinterList = null;
                return false;
            }
        }


    }
}
