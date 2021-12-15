using System;
using System.Threading;
using OposScanner_CCO;

namespace OPOSScannerSample
{
    class Program
    {

        private const string USBCOM_SCANNER_ANY = "RS232Imager";
        private const string USBOEM_SCANNER_HANDHELD = "USBHHScanner";
        private const string USBOEM_SCANNER_FIXED_RETAIL = "USBScanner";
        private const string RS232_SCANNER_ANY = "RS232Scanner";
        private const string RS232SC_SCANNER_FIXED_RETAIL = "SCRS232Scanner";

        private static OPOSScannerClass scanner;

        // The names of the commonly used scanner profiles seen under
        //    HKLM\Software\Wow6432Node\OLEforRetail\ServiceOPOS\SCANNER
        private static String[] names =
        {
            USBCOM_SCANNER_ANY,
            USBOEM_SCANNER_FIXED_RETAIL,
            USBOEM_SCANNER_HANDHELD,
            RS232_SCANNER_ANY,
            RS232SC_SCANNER_FIXED_RETAIL
        };

        static void Main(string[] args)
        {
            // Setup the console program to exit gracefully.
            var exitEvent = new ManualResetEvent(false);
            Console.CancelKeyPress += (sender, eventArgs) =>
            {
                eventArgs.Cancel = true;
                exitEvent.Set();
            };

            // Create a scanner object.
            scanner = new OPOSScannerClass();

            // Attempt to open and claim a scanner (first come, first serve).
            int index, count = names.Length;
            for(index=0; index < count; ++index)
            {
                scanner.Open(names[index]);
                scanner.ClaimDevice(1000);
                if (scanner.Claimed)
                    break;

                scanner.Close();
            }

            if (scanner.Claimed)
            {
                // The scanner has been opened and claimed.
                string profileName = names[index];
                Console.WriteLine("Connected to: " + profileName);

                // Subscribe to the delegate.
                scanner.DataEvent += DataEvent;

                // Enable the scanner and decoding events.
                scanner.DeviceEnabled = true;
                scanner.DataEventEnabled = true;
                scanner.DecodeData = true;

                GoodBeep(ref scanner, profileName);
                Console.WriteLine("Press \'Ctrl + C\' to quit.");

                // Wait for exit event.
                exitEvent.WaitOne();

                // Unsubscribe from the delegate.
                scanner.DataEvent -= DataEvent;

                // Disable, release and close the scanner.
                scanner.DataEventEnabled = false;
                scanner.ReleaseDevice();
                scanner.Close();
            }
            else
            {
                Console.WriteLine("Failed to connect to any scanner.");
            }
        }

        // GoodBeep() provides a simple example of using DirectIO().
        // DirectIO commands are interface-specific.
        // In a commercial application, it would be better to subclass OPOSScannerClass
        // and add interface-specific methods to the subclass.
        static void GoodBeep(ref OPOSScannerClass scanner, string profileName)
        {
            string command = string.Empty;

            if ((profileName == USBOEM_SCANNER_FIXED_RETAIL) || (profileName == USBOEM_SCANNER_HANDHELD))
                command = "30 00 04";
            else if ((profileName == USBCOM_SCANNER_ANY) || (profileName == RS232_SCANNER_ANY))
                command = "42";
            else if (profileName == RS232SC_SCANNER_FIXED_RETAIL)
                command = "33 33 34";  // scanner/scale configuration
                // command = "33 34";  // scanner-only configuration

            if (command != string.Empty)
            {
                Console.WriteLine("GoodBeep()");
                int fodder = 0;
                scanner.DirectIO(-1, ref fodder, ref command);
            }
        }

        static private void DataEvent(int value)
        {
            Console.WriteLine("Data: " + scanner.ScanDataLabel);
            scanner.DataEventEnabled = true;
        }
    }
}