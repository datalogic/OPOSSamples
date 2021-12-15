#pragma once

#include "essential.h"

class ScaleSink : public OposScale_CCO::_IOPOSScaleEvents
{
public:

    ScaleSink(OposScale_CCO::IOPOSScale& scaleObject)
        : scale(scaleObject)
        , ref(0)
    {};

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

    // _IOPOSScaleEvents methods (from OPOSScale.tlh)
    HRESULT DataEvent(long Status);
    HRESULT DirectIOEvent(long EventNumber, long* Data, BSTR* String);
    HRESULT ErrorEvent(long ResultCode, long ResultCodeExtended, long ErrorLocus, long* ErrorResponse);
    HRESULT StatusUpdateEvent(long Data);

private:

    // Event "dispatch ids".
    enum ScaleEvent : DISPID
    {
        Unused   = 0,
        Data     = 1,
        DirectIO = 2,
        Error    = 3,
        Reserved = 4,
        StatusUpdate = 5,
        Count    = 6
    };

private:

    std::string WeightFormat(int weight);
    std::string UnitAbbreviation(int units);

private:

    LONG ref;
    OposScale_CCO::IOPOSScale& scale;
};
