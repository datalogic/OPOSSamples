
#include "ScannerSink.h"


ScannerSink::ScannerSink(OposScanner_CCO::IOPOSScanner& scannerObject, const std::string& profileName)
	: scanner(scannerObject)
	, ifType(Undefined)
	, ref(0)
{
	if ((profileName == USBOEM_SCANNER_HANDHELD) || (profileName == USBOEM_SCANNER_FIXED_RETAIL))
		ifType = UsbOem;
	else if ((profileName == USBCOM_SCANNER_ANY) || (profileName == RS232_SCANNER_ANY))
		ifType = Serial;
	else if (profileName == RS232SC_SCANNER_FIXED_RETAIL)
		ifType = SingleCable;
}

// IUnknown methods
// Returns a pointer to a scanner sink interface for this object, allowing
// the COM object to call this.Invoke()
IFACEMETHODIMP ScannerSink::QueryInterface(REFIID riid, void** ppv)
{
	*ppv = nullptr;
	IID id = __uuidof(OposScanner_CCO::_IOPOSScannerEvents);
	HRESULT hr = E_NOINTERFACE;
	if (riid == IID_IUnknown || riid == IID_IDispatch || riid == id)
	{
		*ppv = static_cast<OposScanner_CCO::_IOPOSScannerEvents*>(this);
		AddRef();
		hr = S_OK;
	}
	return hr;
}

// For reference counting. When IConnectionPoint.Advise() is called on
// this, AddRef() is called.
IFACEMETHODIMP_(ULONG) ScannerSink::AddRef()
{
	return ++ref;
}

// For reference counting. When IConnectionPoint.UnAdvise() is called on
// this, Release() is called and this is destroyed.
IFACEMETHODIMP_(ULONG) ScannerSink::Release()
{
	if (--ref == 0)
		delete this;
	return ref;
}

// IDispatch methods
IFACEMETHODIMP ScannerSink::GetTypeInfoCount(UINT* pctinfo)
{
	*pctinfo = 0;
	return E_NOTIMPL;
}

IFACEMETHODIMP ScannerSink::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** iti)
{
	*iti = nullptr;
	return E_NOTIMPL;
}

// Called by the COM object to retrieve information about the specific
// properties or methods it implements. Sets dispids with the "dispatch id"
// corresponding to the name passed in through names. The dispatch id is
// used later during calls to Invoke().
IFACEMETHODIMP ScannerSink::GetIDsOfNames(REFIID riid, LPOLESTR* names, UINT size, LCID lcid, DISPID* dispids)
{
	if (wcscmp(names[0], L"StatusUpdateEvent") == 0)
		dispids[0] = ScannerEvent::StatusUpdate;
	else if (wcscmp(names[0], L"DirectIOEvent") == 0)
		dispids[0] = ScannerEvent::DirectIO;
	else if (wcscmp(names[0], L"ErrorEvent") == 0)
		dispids[0] = ScannerEvent::Error;
	else if (wcscmp(names[0], L"DataEvent") == 0)
		dispids[0] = ScannerEvent::Data;
	else
		dispids[0] = -1;

	return ((dispids[0] == -1) ? E_NOTIMPL : S_OK);
}

// Called by the COM object when a scanner event occurs. The scanner event
// corresponds to a dispatch id determined by GetIDsOfNames().
IFACEMETHODIMP ScannerSink::Invoke(DISPID dispid, REFIID riid, LCID lcid,
	WORD flags, DISPPARAMS* dispparams, VARIANT* result,
	EXCEPINFO* exceptioninfo, UINT* argerr)
{
	// From Microsoft documentation:
	//    dispParams: MUST point to a DISPPARAMS structure that defines the arguments
	//    passed to the method. Arguments MUST be stored in dispParams.rgvarg in reverse
	//    order, so that the first argument is the one with the highest index in the array.

	switch (dispid)
	{
	case ScannerEvent::Data:
		return DataEvent(dispparams->rgvarg[0].lVal);

	case ScannerEvent::DirectIO:
		return DirectIOEvent(dispparams->rgvarg[2].lVal, dispparams->rgvarg[1].plVal, dispparams->rgvarg[0].pbstrVal);

	case ScannerEvent::Error:
		return ErrorEvent(dispparams->rgvarg[3].lVal, dispparams->rgvarg[2].lVal, dispparams->rgvarg[1].lVal, dispparams->rgvarg[0].plVal);

	case ScannerEvent::StatusUpdate:
		return StatusUpdateEvent(dispparams->rgvarg[0].lVal);

	default:
		return S_OK;
	}
}

// _IOPOSScannerEvents methods
HRESULT ScannerSink::DataEvent(long Status)
{
	std::cout << "Data: " << scanner.ScanDataLabel << std::endl;
	// scanner.DataEventEnabled is set to false when a DataEvent is invoked
	// and so the we must reset it to true to continue recieving DataEvents.
	scanner.DataEventEnabled = true;
	return S_OK;
}

HRESULT ScannerSink::DirectIOEvent(long EventNumber, long* pData, BSTR* pString)
{
	std::cout << "DirectIOEvent( " << EventNumber << ", " << pData << ", " << pString << " )" << std::endl;
	scanner.DirectIO(EventNumber, pData, pString);
	return S_OK;
}

HRESULT ScannerSink::ErrorEvent(long ResultCode, long ResultCodeExtended, long ErrorLocus, long* pErrorResponse)
{
	std::cout << "ErrorEvent( " << ResultCode << ", " << ResultCodeExtended << ", " << ErrorLocus << ", " << pErrorResponse << " )" << std::endl;
	return S_OK;
}

HRESULT ScannerSink::StatusUpdateEvent(long Data)
{
	std::cout << "StatusUpdateEvent( " << Data << " )" << std::endl;
	return S_OK;
}

// GoodBeep() provides a simple example of using DirectIOEvent().
// DirectIO commands are interface-specific.
void ScannerSink::GoodBeep()
{
	std::string command;

	if (ifType == UsbOem)
		command = "30 00 04";
	else if (ifType == Serial)
		command = "42";
	else if (ifType == SingleCable)
		command = "33 33 34";  // scanner/scale configuration
		// command = "33 34";  // scanner-only configuration

	if (!command.empty())
	{
		std::cout << "GoodBeep()" << std::endl;
		long fodder = 0;
		BSTR ioBytes = _com_util::ConvertStringToBSTR( command.c_str() );
		DirectIOEvent(-1, &fodder, &ioBytes);
	}
}
