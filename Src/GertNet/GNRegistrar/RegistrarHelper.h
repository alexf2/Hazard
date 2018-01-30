// RegistrarHelper.h : Declaration of the CRegistrarHelper

#ifndef __REGISTRARHELPER_H_
#define __REGISTRARHELPER_H_

#include "resource.h"       // main symbols
#include <statreg.h>

namespace ATL {

/////////////////////////////////////////////////////////////////////////////
// CRegistrarHelper
class ATL_NO_VTABLE CRegistrarHelper : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CRegistrarHelper, &CLSID_RegistrarHelper>,
	public ISupportErrorInfo,
	public IDispatchImpl<IRegistrarHelper, &IID_IRegistrarHelper, &LIBID_GNREGISTRARLib>
{
public:
	CRegistrarHelper()
	{
	  m_hr = S_OK;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_REGISTRARHELPER)
DECLARE_NOT_AGGREGATABLE(CRegistrarHelper)

DECLARE_PROTECT_FINAL_CONSTRUCT()

HRESULT FinalConstruct();
void FinalRelease()
 {
   //if( m_pReg ) delete m_pReg, m_pReg = NULL;
 }
 
BEGIN_COM_MAP(CRegistrarHelper)
	COM_INTERFACE_ENTRY(IRegistrarHelper)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IRegistrarHelper
public:
	STDMETHOD(get_LastError)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(get_Hresult)(/*[out, retval]*/ long *pVal);
	STDMETHOD(Register)(/*[in]*/ BSTR ObjName, /*[in]*/ BSTR Path, /*[in]*/ short RegFlag, short TypeLibFlag, BSTR PathHelp );

private:
   //CComObject<CDLLRegObject>* m_pReg;
   ATL::CRegObject mReg;
   HRESULT m_hr;
   CComBSTR m_bsErr;

   void ClearError()
	{
	  m_hr = S_OK;
	  m_bsErr.Empty();
	}
   void SetErr( HRESULT hr, LPCWSTR lpw )
	{
	  m_hr = hr;	  
	  m_bsErr = lpw;
	}
};

 } //nm ATL
#endif //__REGISTRARHELPER_H_
