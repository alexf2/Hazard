// FChange.h : Declaration of the CFChange

#ifndef __FCHANGE_H_
#define __FCHANGE_H_

#include "resource.h"       // main symbols

class CFactor;

/////////////////////////////////////////////////////////////////////////////
// CFChange
class ATL_NO_VTABLE CFChange : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFChange, &CLSID_FChange>,
	public ISupportErrorInfo,
	public IDispatchImpl<IFChange, &IID_IFChange, &LIBID_GERTNETLib>,
	public IPersistStreamInit,
	public ILongKey,
	public IClonable
{

friend class CMGertNet;
friend static void ApplyFCToFactor( CComObject<CFChange>* pFC, CComObject<CFactor>* pFactor, bool& bFlOverWrap );

public:
	CFChange()
	{
	  m_bRequiresSave = false;
	}

	bool m_bRequiresSave;

DECLARE_REGISTRY_RESOURCEID(IDR_FCHANGE)
DECLARE_NOT_AGGREGATABLE(CFChange)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CFChange)
	COM_INTERFACE_ENTRY(IFChange)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(ILongKey)
	COM_INTERFACE_ENTRY(IClonable)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IFChange
public:
	STDMETHOD(get_Delta)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_Delta)(/*[in]*/ short newVal);
	STDMETHOD(get_NameFactor)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_NameFactor)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_TCh)(/*[out, retval]*/ TrustChange *pVal);
	STDMETHOD(put_TCh)(/*[in]*/ TrustChange newVal);
	STDMETHOD(get_ID)(/*[out, retval]*/ VARIANT *pVal);

public:
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk);

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
  long m_lID;

  short m_shDelta;
  TrustChange m_tcTrustChange;
  CComBSTR m_bsName;

#pragma pack(push, 1)
  typedef struct
   {
	 long m_lID;

	 short m_shDelta;
	 TrustChange m_tcTrustChange;

	 void LoadToObj( CFChange& rObj );	  
	 void LoadFromObj( const CFChange& rObj );
	  
   } TInternalData;
#pragma pack(pop)
  friend struct TInternalData;
};

inline void CFChange::TInternalData::LoadToObj( CFChange& rObj )
 {
   rObj.m_lID = m_lID;
   rObj.m_shDelta = m_shDelta;
   rObj.m_tcTrustChange = m_tcTrustChange;
 }
inline void CFChange::TInternalData::LoadFromObj( const CFChange& rObj )
 {
   m_lID = rObj.m_lID;
   m_shDelta = rObj.m_shDelta;
   m_tcTrustChange = rObj.m_tcTrustChange;
 }


#endif //__FCHANGE_H_
