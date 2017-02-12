using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Drawing.Printing;
using PrinterSetting.Util;


namespace testOSDefaultPrinterProxy
{
    [TestClass]
    public class testOSDefaultPrinterProxy
    {
        private OSDefaultPrinterProxy osPrintProxy = new OSDefaultPrinterProxy();


        [TestMethod]
        public void TestGetOSDefaultPrinter()
        {
            string printerName;
            string expect = new PrintDocument().PrinterSettings.PrinterName;


            if (!osPrintProxy.GetOSDefaultPrinter(out printerName))
            {
                throw new NullReferenceException("取得系統預設印表機失敗");
            }


            if (!expect.Equals(printerName))
            {
                throw new ArgumentException("Method failed");
            }

        }

        [TestMethod]
        public void TestGetOSPrinterList()
        {
            List<string> printerList;
            string errorPrinterName = "errorPrinter";
            

            if (!osPrintProxy.GetOSPrinterList(out printerList))
            {
                throw new NullReferenceException("查無印表機清單");
            }


            foreach (string printerName in PrinterSettings.InstalledPrinters)
            {

                if (!printerList.Contains(printerName))
                {
                    throw new ArgumentException("Method failed");
                }
            }


            if (printerList.Contains(errorPrinterName))
            {
                throw new ArgumentException("Method failed");
            }
             
        }

        [TestMethod]
        public void TestSetDefaultPrinter()
        {
            string expectPrinterName = "Fax";
            string testPrinterName;


            if (!osPrintProxy.SetDefaultPrinter(expectPrinterName))
            {
                throw new NullReferenceException("Method failed");
            }


            testPrinterName = new PrintDocument().PrinterSettings.PrinterName;


            if (!expectPrinterName.Equals(testPrinterName))
            {
                throw new ArgumentException("系統預設印表機設定值不符合測試期望值");
            }



            if (!osPrintProxy.SetDefaultPrinter(string.Empty))
            {
                throw new NullReferenceException("Method failed");
            }


            testPrinterName = new PrintDocument().PrinterSettings.PrinterName;


            foreach (string printerName in PrinterSettings.InstalledPrinters)
            {

                if (testPrinterName.Contains(printerName))
                {
                    throw new ArgumentException("沒有取消系統預設印表機");
                }
            }


        }


    }
}
