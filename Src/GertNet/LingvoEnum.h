// LingvoEnum.h : Declaration of the CLingvoEnum

#ifndef __LINGVOENUM_H_
#define __LINGVOENUM_H_

#include "resource.h"       // main symbols

#include "CollFunctors.h"

#define NUMBER_LD 11
struct LingvoDescr
 {   
   LPCOLESTR   m_bsMark;
   double      m_dQuality;
 };

struct LingvoDescr2
 {   
   CComBSTR   m_cbsMark;
   double     m_dQuality;
 };

class CFChange;
class CFactor;

/////////////////////////////////////////////////////////////////////////////
// CLingvoEnum
class ATL_NO_VTABLE CLingvoEnum : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CLingvoEnum, &CLSID_LingvoEnum>,
	public ISupportErrorInfo,
	public IDispatchImpl<ILingvoEnum, &IID_ILingvoEnum, &LIBID_GERTNETLib>,
    public IPersistStreamInit,
	public ILongKey,
	public IClonable
	
{

friend class CMGertNet;
friend class CFactor;
friend struct CApplyDirty<CLingvoEnum>;
friend static void ApplyFCToFactor( CComObject<CFChange>* pFC, CComObject<CFactor>* pFactor, bool& bFlOverWrap );

public:
	CLingvoEnum()
	{	  
	  m_bRequiresSave = 0;
	}

	BOOL m_bRequiresSave;

	void SetDirty( BOOL bDirty )
	 {
	   m_bRequiresSave = bDirty;
	 }
	
	BOOL GetDirty()
	 {
	   return m_bRequiresSave;
	 }
	

DECLARE_REGISTRY_RESOURCEID(IDR_LINGVOENUM)
DECLARE_NOT_AGGREGATABLE(CLingvoEnum)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CLingvoEnum)
	COM_INTERFACE_ENTRY(ILingvoEnum)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(ILongKey)
	COM_INTERFACE_ENTRY(IClonable)
END_COM_MAP()


// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


public:
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk);

// ILingvoEnum
	STDMETHOD(get_Mark)(/*[in]*/ long lOrder, /*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Mark)(/*[in]*/ long lOrder, /*[in]*/ BSTR newVal);
	STDMETHOD(get_Quality)(/*[in]*/ long lOrder, /*[out, retval]*/ double *pVal);
	STDMETHOD(put_Quality)(/*[in]*/ long lOrder, /*[in]*/ double newVal);
	STDMETHOD(get_Count)(/*[out, retval]*/ long *pVal);
	STDMETHOD(MarkForValue)(/*[in]*/ double dVal, /*[out, retval]*/ BSTR* pbsMark);

//ILongKey
	STDMETHOD(Set)( long lKey )
	 {
	   m_lID = lKey;
	   m_bRequiresSave = true;
	   return S_OK;
	 }
	STDMETHOD(Get)( long* plKey )
	 {
	   if( plKey == NULL ) return E_POINTER;
	   *plKey = m_lID;
       return S_OK; 
	 }

// IPersistStreamInit
public:
	STDMETHOD(ValueIdx)(/*[in]*/double Val, /*[out,retval]*/short* shIdx);
	STDMETHOD(RoundS)(/*[in]*/double Val, /*[out,retval]*/double* dV);
	STDMETHOD(UpdateFrom)(/*[in]*/ILingvoEnum* pLEnum);
	STDMETHOD(get_IsDirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_ID)(/*[out, retval]*/ VARIANT *pVal);
				
    STDMETHOD(GetClassID)(CLSID *pClassID)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::GetClassID\n"));
		*pClassID = GetObjectCLSID();
		return S_OK;
	 }	
	STDMETHOD(IsDirty)()
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::IsDirty\n"));		
		return m_bRequiresSave ? S_OK : S_FALSE;
	 }

	STDMETHOD(Load)(LPSTREAM pStm);	
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty);
	
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER FAR* /* pcbSize */)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		ATLTRACENOTIMPL(_T("IPersistStreamInit::GetSizeMax"));
		return E_NOTIMPL;
	 }
	
	STDMETHOD(InitNew)();
	 

protected:
   LingvoDescr2 arrLD[ NUMBER_LD ];
   long m_lID;

   static LingvoDescr arrLD_Default[ NUMBER_LD ];


   void Set_OrderOutRange();   
   static long GetCount()
	{
	  return NUMBER_LD;
	}
};

#endif //__LINGVOENUM_H_
