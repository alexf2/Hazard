// SafetyPrecaution.h : Declaration of the CSafetyPrecaution

#ifndef __SAFETYPRECAUTION_H_
#define __SAFETYPRECAUTION_H_

#include "resource.h"       // main symbols
#include "PassErr.h"

#include <vector>
using namespace std;

/////////////////////////////////////////////////////////////////////////////
// CSafetyPrecaution
class ATL_NO_VTABLE CSafetyPrecaution : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSafetyPrecaution, &CLSID_SafetyPrecaution>,
	public ISupportErrorInfo,
	public IDispatchImpl<ISafetyPrecaution, &IID_ISafetyPrecaution, &LIBID_GERTNETLib>,
	public IPersistStreamInit,
	public ILongKey,
	public IClonable
{

friend class CMGertNet;
friend class CCollSF;
friend struct TSampleData;
friend class TOptStep;
friend class TOptInteger;
friend struct TIntTableSlot;

public:
	CSafetyPrecaution()
	{
	  m_bRequiresSave = false;
	}

	bool m_bRequiresSave;

DECLARE_REGISTRY_RESOURCEID(IDR_SAFETYPRECAUTION)
DECLARE_NOT_AGGREGATABLE(CSafetyPrecaution)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CSafetyPrecaution)
	COM_INTERFACE_ENTRY(ISafetyPrecaution)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(ILongKey)
	COM_INTERFACE_ENTRY(IClonable)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ISafetyPrecaution
public:
	STDMETHOD(get_DeltaQ)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_DeltaQ)(/*[in]*/ double newVal);
	STDMETHOD(get_Cost)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_Cost)(/*[in]*/ CURRENCY newVal);
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Name)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_FChanges)(/*[out, retval]*/ IICollFChange* *pVal);
	STDMETHOD(get_ID)(/*[out, retval]*/ VARIANT *pVal);
	STDMETHOD(get_NonCompatibles)(SAFEARRAY** pparrVal);
	STDMETHOD(put_NonCompatibles)(SAFEARRAY** parrVal);

	STDMETHOD(get_Ke)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_Ke)(/*[in]*/ double newVal);
	STDMETHOD(get_Profit)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_Profit)(/*[in]*/ CURRENCY newVal);
	STDMETHOD(get_IsDirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	
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

    STDMETHOD(Load)(LPSTREAM pStm);
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty);
//IPersistStreamInit
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
	   
	   if( m_bRequiresSave ) return S_OK;
	   if( m_spCollFChange.p )
		{
		  CComQIPtr<IPersistStorage> spStor( m_spCollFChange.p );
		  if( spStor.p && spStor->IsDirty() == S_OK ) return S_OK;
		}	   
	   

	   return S_FALSE;
	 }

	
	
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER FAR* /* pcbSize */)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		ATLTRACENOTIMPL(_T("IPersistStreamInit::GetSizeMax"));
		return E_NOTIMPL;
	 }
	
	STDMETHOD(InitNew)()
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::InitNew\n"));	   

	   return InternalIN();
	 }

   double& DeltaRef()
	{
	  return m_dDeltaQ;
	}

protected:
  long m_lID;

  double m_dDeltaQ;
  COleCurrency m_ocCost;
  CComBSTR m_bsName;

  CComPtr<IICollFChange> m_spCollFChange;
  vector<long> m_vecCollNC;

  COleCurrency m_ocProfit;
  double m_dKe;

  HRESULT InternalIN();
  HRESULT InitCollections();

#pragma pack(push, 1)
  typedef struct
   {
     long m_lID;
     double m_dDeltaQ;
	 CURRENCY m_ocCost;
	 	 
	 void LoadToObj( CSafetyPrecaution& rObj );	  
	 void LoadFromObj( const CSafetyPrecaution& rObj );	 
	  
   } TInternalData;
#pragma pack(pop)
  friend struct TInternalData;
};

inline void CSafetyPrecaution::TInternalData::LoadToObj( CSafetyPrecaution& rObj )
 {
   rObj.m_lID = m_lID;
   rObj.m_dDeltaQ = m_dDeltaQ;
   rObj.m_ocCost = m_ocCost;
 }

inline void CSafetyPrecaution::TInternalData::LoadFromObj( const CSafetyPrecaution& rObj )
 {
   m_lID = rObj.m_lID;
   m_dDeltaQ = rObj.m_dDeltaQ;
   m_ocCost = rObj.m_ocCost.m_cur;
 }

#endif //__SAFETYPRECAUTION_H_
