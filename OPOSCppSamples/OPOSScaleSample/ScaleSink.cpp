
#include "ScaleSink.h"

// IUnknown methods
// Returns a pointer to a scale sink interface for this object, allowing
// the COM object to call this.Invoke()
IFACEMETHODIMP ScaleSink::QueryInterface(REFIID riid, void** ppv)
{
	*ppv = nullptr;
	IID id = __uuidof(OposScale_CCO::_IOPOSScaleEvents);
	HRESULT hr = E_NOINTERFACE;
	if (riid == IID_IUnknown || riid == IID_IDispatch || riid == id)
	{
		*ppv = static_cast<OposScale_CCO::_IOPOSScaleEvents*>(this);
		AddRef();
		hr = S_OK;
	}
	return hr;
}

// For reference counting. When IConnectionPoint.Advise() is called on
// this, AddRef() is called.
IFACEMETHODIMP_(ULONG) ScaleSink::AddRef()
{
	return ++ref;
}

// For reference counting. When IConnectionPoint.UnAdvise() is called on
// this, Release() is called and this is destroyed.
IFACEMETHODIMP_(ULONG) ScaleSink::Release()
{
	if (--ref == 0)
		delete this;
	return ref;
}

// IDispatch methods
IFACEMETHODIMP ScaleSink::GetTypeInfoCount(UINT* pctinfo)
{
	*pctinfo = 0;
	return E_NOTIMPL;
}

IFACEMETHODIMP ScaleSink::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** iti)
{
	*iti = nullptr;
	return E_NOTIMPL;
}

// Called by the COM object to retrieve information about the specific
// properties or methods it implements. Sets dispids with the "dispatch id"
// corresponding to the name passed in through names. The dispatch id is
// used later during calls to Invoke().
IFACEMETHODIMP ScaleSink::GetIDsOfNames(REFIID riid, LPOLESTR* names, UINT size, LCID lcid, DISPID* dispids)
{
	if (wcscmp(names[0], L"StatusUpdateEvent") == 0)
		dispids[0] = ScaleEvent::StatusUpdate;
	else if (wcscmp(names[0], L"DirectIOEvent") == 0)
		dispids[0] = ScaleEvent::DirectIO;
	else if (wcscmp(names[0], L"ErrorEvent") == 0)
		dispids[0] = ScaleEvent::Error;
	else if (wcscmp(names[0], L"DataEvent") == 0)
		dispids[0] = ScaleEvent::Data;
	else
		dispids[0] = -1;

	return ((dispids[0] == -1) ? E_NOTIMPL : S_OK);
}

// Called by the COM object when a scanner event occurs. The scanner event
// corresponds to a dispatch id determined by GetIDsOfNames().
IFACEMETHODIMP ScaleSink::Invoke(DISPID dispid, REFIID riid, LCID lcid,
	WORD flags, DISPPARAMS* dispparams, VARIANT* result,
	EXCEPINFO* exceptioninfo, UINT* argerr)
{
	// From Microsoft documentation:
	//    dispParams: MUST point to a DISPPARAMS structure that defines the arguments
	//    passed to the method. Arguments MUST be stored in dispParams.rgvarg in reverse
	//    order, so that the first argument is the one with the highest index in the array.

	switch (dispid)
	{
	case ScaleEvent::Data:
		return DataEvent(dispparams->rgvarg[0].lVal);
	case ScaleEvent::DirectIO:
		return DirectIOEvent(dispparams->rgvarg[2].lVal, dispparams->rgvarg[1].plVal, dispparams->rgvarg[0].pbstrVal);
	case ScaleEvent::Error:
		return ErrorEvent(dispparams->rgvarg[3].lVal, dispparams->rgvarg[2].lVal, dispparams->rgvarg[1].lVal, dispparams->rgvarg[0].plVal);
	case ScaleEvent::StatusUpdate:
		return StatusUpdateEvent(dispparams->rgvarg[0].lVal);
	default:
		return S_OK;
	}
}

// _IOPOSScaleEvents methods
HRESULT ScaleSink::DataEvent(long Status)
{
	std::cout << "DataEvent( " << Status << " )" << std::endl;
	return S_OK;
}

HRESULT ScaleSink::DirectIOEvent(long EventNumber, long* pData, BSTR* pString)
{
	std::cout << "DirectIOEvent( " << EventNumber << ", " << pData << ", " << pString << " )" << std::endl;
	return S_OK;
}

HRESULT ScaleSink::ErrorEvent(long ResultCode, long ResultCodeExtended, long ErrorLocus, long* pErrorResponse)
{
	std::cout << "ErrorEvent( " << ResultCode << ", " << ResultCodeExtended << ", " << ErrorLocus << ", " << pErrorResponse << " )" << std::endl;
	return S_OK;
}

HRESULT ScaleSink::StatusUpdateEvent(long Data)
{
	if (Data == SCAL_SUE_STABLE_WEIGHT)
		std::cout << WeightFormat(scale.ScaleLiveWeight) << std::endl;
	else if (Data == SCAL_SUE_WEIGHT_UNSTABLE)
		std::cout << "Scale weight unstable" << std::endl;
	else if (Data == SCAL_SUE_WEIGHT_ZERO)
		std::cout << WeightFormat(scale.ScaleLiveWeight) << std::endl;
	else if (Data == SCAL_SUE_WEIGHT_OVERWEIGHT)
		std::cout << "Weight limit exceeded." << std::endl;
	else if (Data == SCAL_SUE_NOT_READY)
		std::cout << "Scale not ready." << std::endl;
	else if (Data == SCAL_SUE_WEIGHT_UNDER_ZERO)
		std::cout << "Scale under zero weight." << std::endl;
	else
		std::cout << "Unknown status [" << Data << "]" << std::endl;

	return S_OK;
}

std::string ScaleSink::WeightFormat(int weight)
{
	std::string weightStr;

	std::string units = UnitAbbreviation(scale.WeightUnits);
	if (units.empty())
	{
		weightStr = "Unknown weight unit";
	}
	else
	{
		std::stringstream ss;
		ss << std::fixed << std::setw(6) << std::setprecision(3) << (0.001 * (double)weight) << " " << units;
		weightStr = ss.str();
	}

	return weightStr;
}

std::string ScaleSink::UnitAbbreviation(int units)
{
	std::string unitStr;

	switch (units)
	{
	case SCAL_WU_GRAM:     unitStr = "gr."; break;
	case SCAL_WU_KILOGRAM: unitStr = "kg."; break;
	case SCAL_WU_OUNCE:    unitStr = "oz."; break;
	case SCAL_WU_POUND:    unitStr = "lb."; break;
	}

	return unitStr;
}
