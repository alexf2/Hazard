// CatRegistrar.h : Declaration of the CCatRegistrar

#ifndef __CATREGISTRAR_H_
#define __CATREGISTRAR_H_

#include "resource.h"       // main symbols

extern const CATID CATID_FacValMonitors;
// {2495CD20-6AE3-11d4-8FBA-00504E02C39D};


/////////////////////////////////////////////////////////////////////////////
// CCatRegistrar
class ATL_NO_VTABLE CCatRegistrar : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCatRegistrar, &CLSID_CatRegistrar>,
	public ISupportErrorInfo,
	public IDispatchImpl<ICatRegistrar, &IID_ICatRegistrar, &LIBID_GERTNETLib>
{
public:
	CCatRegistrar() throw()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CATREGISTRAR)
DECLARE_NOT_AGGREGATABLE(CCatRegistrar)
DECLARE_CLASSFACTORY_SINGLETON(CCatRegistrar)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCatRegistrar)
	COM_INTERFACE_ENTRY(ICatRegistrar)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid) throw();

// ICatRegistrar
public:
	STDMETHOD(GetProgIDs)(/*[out, retval]*/ SAFEARRAY** arrProgIDs) throw();
	STDMETHOD(Unregister)() throw();
	STDMETHOD(Register)() throw();

protected:
  HRESULT CreArrrBSTR( list<BSTR>& rL, SAFEARRAY** ppArr ) throw();
};

#endif //__CATREGISTRAR_H_
