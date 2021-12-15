using System;
using System.Threading;
using OposScale_CCO;

// For more convenient shorthand.
using OPOSConstants = OPOSCONSTANTSLib.OPOS_Constants;
using OPOSScaleConstants = OPOSCONSTANTSLib.OPOSScaleConstants;

namespace OPOSScaleSample
{
    class Program
    {
        private static OPOSScaleClass scale;

        // The names of the commonly used scanner profiles seen under
        //    HKLM\Software\Wow6432Node\OLEforRetail\ServiceOPOS\SCALE
        private static String[] names =
        {
            "USBScale",
            "RS232Scale",
            "SCRS232Scale"
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

            // Create a scale object.
            scale = new OPOSScaleClass();

            // Attempt to open and claim a scale (first come, first serve).
            int index, count = names.Length;
            for (index = 0; index < count; ++index)
            {
                scale.Open(names[index]);
                scale.ClaimDevice(1000);
                if (scale.Claimed)
                    break;

                scale.Close();
            }

            if (scale.Claimed)
            {
                // The scale has been opened and claimed.
                string profileName = names[index];
                Console.WriteLine("Connected to: " + profileName);

                if (scale.CapStatusUpdate)
                {
                    // Tell the scale we intend to perform "live" weighing.
                    scale.StatusNotify = (int)OPOSScaleConstants.SCAL_SN_ENABLED;
                    if (scale.ResultCode == (int)OPOSConstants.OPOS_SUCCESS)
                    {
                        // Subscribe to the delegate.
                        scale.StatusUpdateEvent += StatusUpdateEvent;

                        // Enable scale events.
                        scale.DeviceEnabled = true;
                        if (scale.DeviceEnabled)
                        {
                            scale.DataEventEnabled = true;

                            Console.WriteLine("Live weighing enabled.");

                            // Wait for exit event.
                            Console.WriteLine("Press \'Ctrl + C\' to quit.");
                            exitEvent.WaitOne();

                            // Disable, release and close the scale.
                            scale.DataEventEnabled = false;
                        }
                        else
                        {
                            Console.WriteLine("Failed to enable the scale. Error code: " + scale.ResultCode);
                        }

                        // Unsubscribe from the delegate.
                        scale.StatusUpdateEvent -= StatusUpdateEvent;

                        // Release and close the scale.
                        scale.ReleaseDevice();
                        scale.Close();
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to any scale.");
            }
        }

        static private void StatusUpdateEvent(int value)
        {
            int status = (int) scale.ResultCode;

            if (value == (int)OPOSScaleConstants.SCAL_SUE_STABLE_WEIGHT)
            {
                Console.WriteLine(WeightFormat(scale.ScaleLiveWeight));
            }
            else if (value == (int)OPOSScaleConstants.SCAL_SUE_WEIGHT_UNSTABLE)
            {
                Console.WriteLine("Scale weight unstable");
            }
            else if (value == (int)OPOSScaleConstants.SCAL_SUE_WEIGHT_ZERO)
            {
                Console.WriteLine(WeightFormat(scale.ScaleLiveWeight));
            }
            else if (value == (int)OPOSScaleConstants.SCAL_SUE_WEIGHT_OVERWEIGHT)
            {
                Console.WriteLine("Weight limit exceeded.");
            }
            else if (value == (int)OPOSScaleConstants.SCAL_SUE_NOT_READY)
            {
                Console.WriteLine("Scale not ready.");
            }
            else if (value == (int)OPOSScaleConstants.SCAL_SUE_WEIGHT_UNDER_ZERO)
            {
                Console.WriteLine("Scale under zero weight.");
            }
            else
            {
                Console.WriteLine("Unknown status [{0}]", value);
            }
        }

        static private string WeightFormat(int weight)
        {
            string weightStr = string.Empty;

            string units = UnitAbbreviation(scale.WeightUnits);
            if (units == string.Empty)
            {
                weightStr = string.Format("Unknown weight unit");
            }
            else
            {
                double dWeight = 0.001 * (double)weight;
                weightStr = string.Format("{0:0.000} {1}", dWeight, units);
            }

            return weightStr;
        }

        static private string UnitAbbreviation(int units)
        {
            string unitStr = string.Empty;

            switch ((OPOSScaleConstants)units)
            {
                case OPOSScaleConstants.SCAL_WU_GRAM: unitStr = "gr."; break;
                case OPOSScaleConstants.SCAL_WU_KILOGRAM: unitStr = "Kg."; break;
                case OPOSScaleConstants.SCAL_WU_OUNCE: unitStr = "oz."; break;
                case OPOSScaleConstants.SCAL_WU_POUND: unitStr = "lb."; break;
            }

            return unitStr;
        }
    }
}
