#pragma once

#include "essential.h"

class ScannerSink : public OposScanner_CCO::_IOPOSScannerEvents
{
public:

    ScannerSink(OposScanner_CCO::IOPOSScanner& scannerObject, const std::string& profileName);

    // IUnknown methods 
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IDispatch methods
    IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo);
    IFACEMETHODIMP GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** iti);
    IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* names,
        UINT size, LCID lcid, DISPID* rgDispId);
    IFACEMETHODIMP Invoke(DISPID dispid, REFIID riid, LCID lcid,
        WORD flags, DISPPARAMS* dispparams, VARIANT* result,
        EXCEPINFO* exceptioninfo, UINT* argerr);

    // _IOPOSScannerEvents methods (from OPOSScanner.tlh)
    HRESULT DataEvent(long Status);
    HRESULT DirectIOEvent(long EventNumber, long* Data, BSTR* String);
    HRESULT ErrorEvent(long ResultCode, long ResultCodeExtended, long ErrorLocus, long* ErrorResponse);
    HRESULT StatusUpdateEvent(long Data);

    void GoodBeep();

private:

    // Event "dispatch ids".
    enum ScannerEvent : DISPID
    {
        Unused   = 0,
        Data     = 1,
        DirectIO = 2,
        Error    = 3,
        Reserved = 4,
        StatusUpdate = 5,
        Count    = 6
    };

    enum ScannerInterfaceType
    {
        Undefined   = 0,
        UsbOem      = 1,
        Serial      = 2,
        SingleCable = 3
    };

private:

    LONG ref;
    OposScanner_CCO::IOPOSScanner& scanner;
    ScannerInterfaceType ifType;
};
