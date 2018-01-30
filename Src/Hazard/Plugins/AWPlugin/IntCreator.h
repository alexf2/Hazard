// IntCreator.h : Declaration of the CIntCreator

#ifndef __INTCREATOR_H_
#define __INTCREATOR_H_

#include "resource.h"       // main symbols

//#include "FactorTable.h"
//#include "GostTable.h"



/////////////////////////////////////////////////////////////////////////////
// CIntCreator
class ATL_NO_VTABLE CIntCreator : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CIntCreator, &CLSID_IntCreator>,
	public ISupportErrorInfo,
	public IDispatchImpl<IIntCreator, &IID_IIntCreator, &LIBID_AWPLUGINLib>
	
{
public:
	CIntCreator()  throw()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_INTCREATOR)
DECLARE_NOT_AGGREGATABLE(CIntCreator)
DECLARE_CLASSFACTORY_SINGLETON(CIntCreator)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CIntCreator)
	COM_INTERFACE_ENTRY(IIntCreator)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)  throw();

// IIntCreator
public:
	STDMETHOD(NewICollGosts)(/*[in]*/BSTR Name, /*[in]*/IStorage* Stg, /*[out, retval]*/ICollGosts** pCollGosts);
	STDMETHOD(NewIGostTable)(/*[in]*/BSTR Name,/*[out, retval]*/ IGostTable** GostTable)  throw();
	STDMETHOD(NewIFactorTable)(/*[in]*/BSTR Name,/*[out, retval]*/ IFactorTable** FactorTable)  throw();

	STDMETHOD(FetchGostTable)( /*[in]*/IStorage* StmRoot, /*[in]*/SAFEARRAY** Path, /*[in]*/long OFThrough, /*[in,optional]*/VARIANT* OFEnd, /*[in,optional]*/VARIANT* CreCached, /*[out,retval]*/ IGostTable** pGT ) throw();
	STDMETHOD(FetchFactorTable)( /*[in]*/IStorage* StmRoot, /*[in]*/SAFEARRAY** Path, /*[in]*/long OFThrough, /*[in,optional]*/VARIANT* OFEnd, /*[in,optional]*/VARIANT* CreCached, /*[out,retval]*/ IFactorTable** pFT ) throw();

protected:
   HRESULT __fastcall Fetch( IStorage* StmRoot, DWORD OpenFlagsT, DWORD OpenFlagsE, SAFEARRAY** Path, bool bLastIsStream, CComPtr<IUnknown>* pRes ) throw();
};


#endif //__INTCREATOR_H_
