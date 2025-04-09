using System;
using System.IO;
using System.Runtime.InteropServices;

namespace BulkZPLPrinter
{
    /// <summary>
    /// Raw printer helper class for sending raw data directly to printers
    /// This is an alternative implementation that uses Windows API for better ZPL printing
    /// </summary>
    public class RawPrinterHelper
    {
        // Structure and API declarations
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        /// <summary>
        /// Send raw data to a printer
        /// </summary>
        /// <param name="szPrinterName">Printer name</param>
        /// <param name="szString">Raw data to print</param>
        /// <returns>True on success, false on failure</returns>
        public static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr hPrinter = IntPtr.Zero;
            DOCINFOA di = new DOCINFOA();
            di.pDocName = "ZPL Document";
            di.pDataType = "RAW";

            // Open the printer
            if (!OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                return false;
            }

            // Start a document
            if (!StartDocPrinter(hPrinter, 1, di))
            {
                ClosePrinter(hPrinter);
                return false;
            }

            // Start a page
            if (!StartPagePrinter(hPrinter))
            {
                EndDocPrinter(hPrinter);
                ClosePrinter(hPrinter);
                return false;
            }

            // Send the data to the printer
            bool success = false;
            IntPtr pBytes;
            Int32 dwCount = szString.Length;

            // Allocate unmanaged memory and copy the string to it
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            Int32 dwWritten = 0;
            success = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
            Marshal.FreeCoTaskMem(pBytes);

            // End the page
            EndPagePrinter(hPrinter);

            // End the document
            EndDocPrinter(hPrinter);

            // Close the printer
            ClosePrinter(hPrinter);

            // Return success/failure
            return success;
        }

        /// <summary>
        /// Send raw bytes to a printer
        /// </summary>
        /// <param name="szPrinterName">Printer name</param>
        /// <param name="bytes">Raw data to print</param>
        /// <returns>True on success, false on failure</returns>
        public static bool SendBytesToPrinter(string szPrinterName, byte[] bytes)
        {
            IntPtr hPrinter = IntPtr.Zero;
            DOCINFOA di = new DOCINFOA();
            di.pDocName = "ZPL Document";
            di.pDataType = "RAW";

            // Open the printer
            if (!OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                return false;
            }

            // Start a document
            if (!StartDocPrinter(hPrinter, 1, di))
            {
                ClosePrinter(hPrinter);
                return false;
            }

            // Start a page
            if (!StartPagePrinter(hPrinter))
            {
                EndDocPrinter(hPrinter);
                ClosePrinter(hPrinter);
                return false;
            }

            // Send the data to the printer
            bool success = false;
            IntPtr pBytes;
            Int32 dwCount = bytes.Length;

            // Allocate unmanaged memory and copy the bytes to it
            pBytes = Marshal.AllocCoTaskMem(dwCount);
            Marshal.Copy(bytes, 0, pBytes, dwCount);
            Int32 dwWritten = 0;
            success = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
            Marshal.FreeCoTaskMem(pBytes);

            // End the page
            EndPagePrinter(hPrinter);

            // End the document
            EndDocPrinter(hPrinter);

            // Close the printer
            ClosePrinter(hPrinter);

            // Return success/failure
            return success;
        }

        /// <summary>
        /// Send a file to a printer
        /// </summary>
        /// <param name="szPrinterName">Printer name</param>
        /// <param name="szFileName">File to print</param>
        /// <returns>True on success, false on failure</returns>
        public static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            try
            {
                // Read file content
                FileStream fs = new FileStream(szFileName, FileMode.Open);
                BinaryReader br = new BinaryReader(fs);
                byte[] bytes = new byte[fs.Length];
                br.Read(bytes, 0, (int)fs.Length);
                br.Close();
                fs.Close();

                // Send the file content to printer
                return SendBytesToPrinter(szPrinterName, bytes);
            }
            catch (Exception ex)
            {
                throw new Exception("Error sending file to printer", ex);
            }
        }
    }
}