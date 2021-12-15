
#include "ScannerSink.h"

static DWORD threadID;
static BOOL handler(DWORD event);

int main()
{
    // Setup the console program to exit gracefully.
    threadID = GetCurrentThreadId();
    SetConsoleCtrlHandler((PHANDLER_ROUTINE)(handler), TRUE);

    // The names of the commonly used scanner profiles seen under
    //    HKLM\Software\Wow6432Node\OLEforRetail\ServiceOPOS\SCANNER
    // For the sake of this example project, the profile names used
    // here are #define'd in essential.h because they are also referenced
    // in ScannerSink.cpp
    std::vector<std::string> keys =
    {
        USBCOM_SCANNER_ANY,
        USBOEM_SCANNER_HANDHELD,
        USBOEM_SCANNER_FIXED_RETAIL,
        RS232_SCANNER_ANY,
        RS232SC_SCANNER_FIXED_RETAIL
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
    OposScanner_CCO::IOPOSScannerPtr scanner;
    scanner.CreateInstance("OPOS.Scanner");

    // Attempt to open and claim a scanner (first come, first serve).
    size_t index, count = keys.size();
    for (index = 0; index < count; ++index)
    {
        scanner->Open( keys[index].c_str() );
        scanner->ClaimDevice(1000);
        if (scanner->Claimed)
            break;

        scanner->Close();
    }

    if (scanner->Claimed)
    {
        // The scanner has been opened and claimed.
        std::string profileName(keys[index]);
        std::cout << "Connected to: " << profileName << std::endl;

        // Enable the scanner and decoding events.
        scanner->DeviceEnabled = true;
        scanner->DataEventEnabled = true;
        scanner->DecodeData = true;

        // Determine whether scanner is connectable
        IConnectionPointContainer* cpc;
        bool isConnectable = (scanner->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc) == S_OK);

        if (isConnectable)
        {
            // Determine whether _IOPOSScannerEvents connection point is supported.
            IConnectionPoint* cp;
            bool haveConnectionPoint = (cpc->FindConnectionPoint(__uuidof(OposScanner_CCO::_IOPOSScannerEvents), &cp) == S_OK);
            cpc->Release();

            if (haveConnectionPoint)
            {
                ScannerSink* sink = new ScannerSink(*scanner, profileName);

                // Connect cp with sink (subscribe to the sink).
                // cookie is a token representing the connection,
                // used later when deleting the connection.
                DWORD cookie;
                cp->Advise(sink, &cookie);

                sink->GoodBeep();
                std::cout << "Press \'Ctrl + C\' to quit." << std::endl;

                // The scanner message loop. Events will be handled by the methods of the sink.
                static MSG msg = { 0 };
                while (GetMessage(&msg, 0, 0, 0))
                {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }

                // Delete the connection (unsubscribe from the sink).
                cp->Unadvise(cookie);
                cp->Release();
            }
        }

        // Disable, release and close the scanner.
        scanner->DeviceEnabled = false;
        scanner->ReleaseDevice();
        scanner->Close();
    }
    else
    {
        std::cout << "Failed to connect to any scanner." << std::endl;
    }

    // Release the COM object
    scanner.Release();

    // Unload libraries on this thread.
    CoUninitialize();
    return 0;
}

BOOL handler(DWORD event)
{
    PostThreadMessage(threadID, WM_QUIT, 0, 0);
    return TRUE;
}
