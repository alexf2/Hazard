// Factor.h : Declaration of the CFactor

#ifndef __FACTOR_H_
#define __FACTOR_H_

#include "resource.h"       // main symbols

#include "CollFunctors.h"
#include "FactorCP.h"
#include "idx.h"

#include "vector"
using namespace std;

class CMGertNet;
class CFChange;
struct TFacSnapshot;
		

/////////////////////////////////////////////////////////////////////////////
// CFactor
class ATL_NO_VTABLE CFactor : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFactor, &CLSID_Factor>,
	public ISupportErrorInfo,
	public IDispatchImpl<IFactor, &IID_IFactor, &LIBID_GERTNETLib>,	
	public IPersistStreamInit,
	public IBSTRKey,
	public IClonable,
	public CProxy_IFacEvents< CFactor >,
	public IConnectionPointContainerImpl<CFactor>
{

friend class CMGertNet;
friend class CRndFunction;
friend struct CApplyDirty<CFactor>;
friend struct TFacSnapshot;
friend static void ApplyFCToFactor( CComObject<CFChange>* pFC, CComObject<CFactor>* pFactor, bool& bFlOverWrap );


public:
	CFactor()
	{
	  m_bRequiresSave = 0;
	  m_sOverwrap = 0;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_FACTOR)
DECLARE_NOT_AGGREGATABLE(CFactor)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CFactor)
	COM_INTERFACE_ENTRY(IFactor)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)	
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(IBSTRKey)
	COM_INTERFACE_ENTRY(IClonable)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


public:
	STDMETHOD(get_IsDirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_ID)(/*[out, retval]*/ VARIANT *pVal);
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk);

// IFactor
	STDMETHOD(get_TrustLvl)(/*[out, retval]*/ TrustLevel *pVal);
	STDMETHOD(put_TrustLvl)(/*[in]*/ TrustLevel newVal);
	STDMETHOD(get_Value)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_Value)(/*[in]*/ short newVal);
	STDMETHOD(get_IDEnum)(/*[out, retval]*/ long *pVal);	
	STDMETHOD(put_IDEnum)(/*[in]*/ long lVal);	
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Name)(/*[in]*/ BSTR newVal);
              
	STDMETHOD(AddOwerWrap)();
	STDMETHOD(ResetOverWrap)();
	STDMETHOD(get_Overwrap)(/*[out, retval]*/ short *pVal);


	//ILongKey
	STDMETHOD(Set)( BSTR bsKey )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   m_bsShortName = bsKey;
	   m_bRequiresSave = true;
	   Fire_OnDirtyChanged( VARIANT_TRUE );
	   return S_OK;
	 }
	STDMETHOD(Get)( BSTR* pbsKey )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   if( pbsKey == NULL ) return E_POINTER;
	   *pbsKey = m_bsShortName.Copy();
       return S_OK; 
	 }

// IPersistStreamInit
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

	short GetNIdx();

protected:
    CComBSTR m_bsName, m_bsShortName;
	long m_lIDEnum;
	short m_sValue;
	TrustLevel m_tr;

	float arrshInd[ NUMBER_IDX ];

	PrpbDistrTyp m_pdt;
	float m_fPlacement, m_fScale;

	short m_sOverwrap;

	BOOL m_bRequiresSave;

#pragma pack(push,1)
	typedef struct
	 {
	   long m_lIDEnum;
	   short m_sValue;
	   TrustLevel m_tr;

	   PrpbDistrTyp m_pdt;
	   float m_fPlacement, m_fScale;

	   float arrshInd[ NUMBER_IDX ];

	   void LoadToObj( CFactor& rObj );	  
	   void LoadFromObj( const CFactor& rObj );
		
	 } TInternalData;
#pragma pack(pop)
    friend struct TInternalData;

public :
	STDMETHOD(get_CochiScale)(/*[out, retval]*/ float *pVal);
	STDMETHOD(put_CochiScale)(/*[in]*/ float newVal);
	STDMETHOD(get_CochiPlacement)(/*[out, retval]*/ float *pVal);
	STDMETHOD(put_CochiPlacement)(/*[in]*/ float newVal);
	STDMETHOD(get_DistrType)(/*[out, retval]*/ PrpbDistrTyp *pVal);
	STDMETHOD(put_DistrType)(/*[in]*/ PrpbDistrTyp newVal);
	STDMETHOD(get_NIdx)(/*[out, retval]*/ short *pVal);
	STDMETHOD(get_Idx)(/*[in]*/short sIdx, /*[out, retval]*/ float *pVal);
	STDMETHOD(put_Idx)(/*[in]*/short sIdx, /*[in]*/ float newVal);
	
	

BEGIN_CONNECTION_POINT_MAP(CFactor)
	CONNECTION_POINT_ENTRY(DIID__IFacEvents)
END_CONNECTION_POINT_MAP()

};

inline void CFactor::TInternalData::LoadToObj( CFactor& rObj )
 {
   rObj.m_lIDEnum = m_lIDEnum;
   rObj.m_sValue = m_sValue;
   rObj.m_tr = m_tr;
   
   rObj.m_pdt = m_pdt;
   rObj.m_fPlacement = m_fPlacement;
   rObj.m_fScale = m_fScale;

   memcpy( rObj.arrshInd, arrshInd, sizeof arrshInd );
 }
inline void CFactor::TInternalData::LoadFromObj( const CFactor& rObj )
 {
   m_lIDEnum = rObj.m_lIDEnum;
   m_sValue = rObj.m_sValue;
   m_tr = rObj.m_tr;
   
   m_pdt = rObj.m_pdt;
   m_fPlacement = rObj.m_fPlacement;
   m_fScale = rObj.m_fScale;

   memcpy( arrshInd, rObj.arrshInd, sizeof arrshInd );
 }



struct TFacSnapshot
 {
    TFacSnapshot()
	 {
	 }
    TFacSnapshot( const CComObject<CFactor>& rFac )
	 {
	   operator=( rFac );
	 }
	TFacSnapshot( const TFacSnapshot& rSn )
	 {
	   operator=( rSn );
	 }

	TFacSnapshot& operator=( const TFacSnapshot& rSn )
	 {
	   m_sValue = rSn.m_sValue;
	   m_tr = rSn.m_tr;
	   memcpy( arrshInd, rSn.arrshInd, sizeof(arrshInd) );
	   m_pdt = rSn.m_pdt;
	   m_fPlacement = rSn.m_fPlacement;
	   m_fScale = rSn.m_fScale;
	   //m_bRequiresSave = rSn.m_bRequiresSave;
	   return *this;
	 }

	TFacSnapshot& operator=( const CComObject<CFactor>& rFac )
	 {
	   m_sValue = rFac.m_sValue;
	   m_tr = rFac.m_tr;
	   memcpy( arrshInd, rFac.arrshInd, sizeof(arrshInd) );
	   m_pdt = rFac.m_pdt;
	   m_fPlacement = rFac.m_fPlacement;
	   m_fScale = rFac.m_fScale;
	   //m_bRequiresSave = rFac.m_bRequiresSave;
	   return *this;
	 }

	void LoadTo( CComObject<CFactor>& rFac )
	 {
	   rFac.m_sValue = m_sValue;
	   rFac.m_tr = m_tr;
	   memcpy( rFac.arrshInd, arrshInd, sizeof(arrshInd) );
	   m_pdt = rFac.m_pdt;
	   m_fPlacement = rFac.m_fPlacement;
	   m_fScale = rFac.m_fScale;
	 }

    short m_sValue;
	TrustLevel m_tr;
	float arrshInd[ NUMBER_IDX ];

	PrpbDistrTyp m_pdt;
	float m_fPlacement, m_fScale;
	//BOOL m_bRequiresSave;
 };


typedef vector<TFacSnapshot> VEC_TFacSnapshot;
typedef VEC_TFacSnapshot::iterator IT_VEC_TFacSnapshot;

#endif //__FACTOR_H_
