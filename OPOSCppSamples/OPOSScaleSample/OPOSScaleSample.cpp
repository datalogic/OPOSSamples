
#include "ScaleSink.h"

static DWORD threadID;
static BOOL handler(DWORD event);

int main()
{
    // Setup the console program to exit gracefully.
    threadID = GetCurrentThreadId();
    SetConsoleCtrlHandler((PHANDLER_ROUTINE)(handler), TRUE);

    // The names of the commonly used scale profiles seen under
    //   HKLM\Software\Wow6432Node\OLEforRetail\ServiceOPOS\SCALE
    std::vector<std::string> keys =
    {
        "USBScale",
        "RS232Scale",
        "SCRS232Scale"
    };

    // Initialize COM.
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
    {
        std::cout << "Failed to initialize COM library. Error code = 0x"
            << std::hex << hr << std::endl;
        return hr;
    }

    // Create a COM object and point to it.
    OposScale_CCO::IOPOSScalePtr scale;
    scale.CreateInstance("OPOS.Scale");

    // Attempt to open and claim a scale (first come, first serve).
    size_t index, count = keys.size();
    for (index = 0; index < count; ++index)
    {
        scale->Open(keys[index].c_str());
        scale->ClaimDevice(1000);
        if (scale->Claimed)
            break;

        scale->Close();
    }

    if (scale->Claimed)
    {
        // The scale has been opened and claimed.
        std::string profileName(keys[index]);
        std::cout << "Connected to: " << profileName << std::endl;

        // Determine whether scale is connectable
        IConnectionPointContainer* cpc;
        bool isConnectable = (scale->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc) == S_OK);

        if (isConnectable)
        {
            // Determine whether _IOPOSScaleEvents connection point is supported.
            IConnectionPoint* cp;
            bool haveConnectionPoint = (cpc->FindConnectionPoint(__uuidof(OposScale_CCO::_IOPOSScaleEvents), &cp) == S_OK);
            cpc->Release();

            if (haveConnectionPoint)
            {
                ScaleSink* sink = new ScaleSink(*scale);

                // Connect cp with sink (subscribe to the sink).
                // cookie is a token representing the connection,
                // used later when deleting the connection.
                DWORD cookie;
                cp->Advise(sink, &cookie);

                if (scale->CapStatusUpdate)
                {
                    // Tell the scale we intend to perform "live" weighing.
                    scale->StatusNotify = SCAL_SN_ENABLED;
                    if (scale->ResultCode == OPOS_SUCCESS)
                    {
                        // Enable scale events.
                        scale->DeviceEnabled = true;
                        if (scale->DeviceEnabled)
                        {
                            scale->DataEventEnabled = true;

                            std::cout << "Live weighing enabled." << std::endl;
                            std::cout << "Press \'Ctrl + C\' to quit." << std::endl;

                            // The scale message loop. Events will be handled by the methods of the sink.
                            static MSG msg = { 0 };
                            while (GetMessage(&msg, 0, 0, 0))
                            {
                                TranslateMessage(&msg);
                                DispatchMessage(&msg);
                            }

                            scale->DataEventEnabled = false;
                        }
                    }
                }

                // Delete the connection (unsubscribe from the sink).
                cp->Unadvise(cookie);
                cp->Release();
            }
        }

        // Disable, release and close the scale.
        scale->DeviceEnabled = false;
        scale->ReleaseDevice();
        scale->Close();
    }
    else
    {
        std::cout << "Failed to connect to any scale." << std::endl;
    }

    // Release the COM object
    scale.Release();

    // Unload libraries on this thread.
    CoUninitialize();
    return 0;
}

BOOL handler(DWORD event)
{
    PostThreadMessage(threadID, WM_QUIT, 0, 0);
    return TRUE;
}
