// FactorTable.h : Declaration of the CFactorTable

#ifndef __FACTORTABLE_H_
#define __FACTORTABLE_H_

#include "resource.h"       // main symbols

struct CFactorSlot
 {
   CFactorSlot() throw()
	{
	  m_fValuation = m_fAveWeighted = m_fWeight = 0;
	  m_bCalculated = false; m_bLocked = false;
	  m_ftTyp = FT_None;
	  memset( &m_Lnk, 0, sizeof m_Lnk );
	}

   CFactorSlot( int ) throw():
	  m_bsComponentName( L"<Нет>" ),
	  m_bsLingvoVal( L"<Нет>" )
	{      
	  m_fWeight = m_fValuation = m_fAveWeighted = 0;
	  m_bCalculated = false; m_bLocked = false;
	  m_ftTyp = FT_None;
	  memset( &m_Lnk, 0, sizeof m_Lnk );
	}

   CFactorSlot( const CFactorSlot& rObj ) throw()
	{
	  m_bsComponentName.m_str = 0;	  
	  m_bsLingvoVal.m_str = 0;
	  m_bsComment.m_str = 0;
	  m_ftTyp = FT_None;
	  memset( &m_Lnk, 0, sizeof m_Lnk );
	  
	  this->operator=( rObj );
	}

   ~CFactorSlot() throw()
	{ 	
	  ReleasePtr();
	}

   void ReleasePtr() throw()
	{
	  if( m_ftTyp == FT_Self )
	   {
	     if( m_Lnk.Self.pFT ) m_Lnk.Self.pFT->Release();
		 m_Lnk.Self.pFT = NULL;
	   }
	  else if( m_ftTyp == FT_Gost )
	   {
	     if( m_Lnk.Gost.pGT ) m_Lnk.Gost.pGT->Release();
		 m_Lnk.Gost.pGT = NULL;
	   }	  
	}
   
   CFactorSlot& operator=( const CFactorSlot& rObj ) throw()
	{
	  
	  m_bsComponentName = rObj.m_bsComponentName;
      m_fWeight = rObj.m_fWeight;
      m_bsLingvoVal = rObj.m_bsLingvoVal;
	  m_bsComment = rObj.m_bsComment;
      m_fValuation = rObj.m_fValuation;
	  m_fAveWeighted = rObj.m_fAveWeighted;
	  m_bCalculated = rObj.m_bCalculated;
	  m_bLocked = rObj.m_bLocked;

	  ReleasePtr();

	  m_ftTyp = rObj.m_ftTyp;
	  memcpy( &m_Lnk, &rObj.m_Lnk, sizeof m_Lnk );
	  if( m_ftTyp == FT_Self )
	   {
	     if( m_Lnk.Self.pFT ) m_Lnk.Self.pFT->AddRef();
	   }
	  else if( m_ftTyp == FT_Gost )
	   {
	     if( m_Lnk.Gost.pGT ) m_Lnk.Gost.pGT->AddRef();
	   }
	  
	  return *this;
	}

   bool operator==( const CFactorSlot& rObj ) throw()
	{
	  return CmpBStrNoCase( m_bsComponentName, rObj.m_bsComponentName );	         
	}

   HRESULT WriteToStream( LPSTREAM ) throw();
   HRESULT ReadFromStream( LPSTREAM ) throw();

   void MkNone() throw()
	{
	  ReleasePtr();
	  m_ftTyp = FT_None;
	  //m_bCalculated = false;
	  Uncalc();
	}

   void MkSelf( long lKey, IFactorTable* pFT ) throw()
	{
	  ReleasePtr();
	  m_ftTyp = FT_Self;
	  m_Lnk.Self.lKey = lKey;
	  m_Lnk.Self.pFT = pFT; if( pFT ) pFT->AddRef();	  
	  //m_bCalculated = false;
	  Uncalc();
	}

   void MkGost( long lKeyTopic, long lKeyGost, IGostTable* pGT ) throw()
	{
	  ReleasePtr();
	  m_ftTyp = FT_Gost;
	  m_Lnk.Gost.lKeyTopic = lKeyTopic;
	  m_Lnk.Gost.lKeyGost = lKeyGost;
	  m_Lnk.Gost.pGT = pGT; if( pGT ) pGT->AddRef();
	  m_Lnk.Gost.iSlot = -1;
	  //m_bCalculated = false;
	  Uncalc();
	}

   bool IsAllAssigned() const  throw()
	{
	  bool bFl;
	  if( m_ftTyp == FT_None ) 
	    bFl = false;
	  else if( m_ftTyp == FT_Gost )
	    bFl = (m_Lnk.Gost.iSlot == -1 || m_Lnk.Gost.pGT == NULL) ? false:true;
	  else if( m_Lnk.Self.pFT == NULL ) 
		bFl = false;
	  else
	   {
		 VARIANT_BOOL vbFl;
		 HRESULT hr = m_Lnk.Self.pFT->get_AllComponentsIsAssigned( &vbFl );
		 bFl = (SUCCEEDED(hr) && vbFl == VARIANT_TRUE);
	   }
	  return bFl;
	}

   
   enum  EN_AssErrors
	{
	  AE_OK,
	  AE_SlotUnbounded,
	  AE_Failed,
	  AE_SlotIdxOutOfRange,
	  AE_CanntGetValue
	};

   EN_AssErrors AssignValue( int iSlotNumber ) throw()
	{
	  if( m_ftTyp != FT_Gost || m_Lnk.Gost.pGT == NULL ) 
	     return AE_SlotUnbounded;

      if( iSlotNumber == -1 )
	   {
	     m_fAveWeighted = m_fValuation = 0;
	     m_Lnk.Gost.iSlot = -1; m_bsLingvoVal = L"<Нет>"; 
		 //m_bCalculated = false;
		 Uncalc();
		 return AE_OK;
	   }

	  long lNSlots;
	  HRESULT hr = m_Lnk.Gost.pGT->get_NumberSlots( &lNSlots );
	  if( FAILED(hr) ) return AE_Failed;
	  if( iSlotNumber < 0 || iSlotNumber >= lNSlots ) return AE_SlotIdxOutOfRange;
	  
	  //m_bCalculated = false;
	  Uncalc();
      m_Lnk.Gost.iSlot = iSlotNumber;

	  return Calculate();	  
	}

   EN_AssErrors Calculate() throw()
	{	  
	  EN_AssErrors enAE = AE_OK;

	  if( m_bLocked )
	    m_fAveWeighted = m_fWeight * m_fValuation,
		m_bCalculated = true;
	  else
	   {
		 HRESULT hr = m_Lnk.Gost.pGT->get_NValue( m_Lnk.Gost.iSlot, &m_fValuation );
		 if( FAILED(hr) ) enAE = AE_CanntGetValue;
		 else
		  {
			m_bsLingvoVal.Empty();	  
			hr = m_Lnk.Gost.pGT->get_NLingvoVal( m_Lnk.Gost.iSlot, &m_bsLingvoVal );
			if( FAILED(hr) ) enAE = AE_CanntGetValue;
			else
			 {	  	  		 			
			   m_fAveWeighted = m_fWeight * m_fValuation;
			   m_bCalculated = true;
			 }
		  }
	   }

	  return enAE;
	}

   void Uncalc()
	{
	  m_bCalculated = false;
	  if( !m_bLocked ) m_bsLingvoVal.Empty();
	}

   CComBSTR m_bsComponentName;
   float m_fWeight;
   CComBSTR m_bsLingvoVal, m_bsComment;
   float m_fValuation, m_fAveWeighted;
   bool m_bCalculated, m_bLocked;

   FTSlotTyp m_ftTyp;
   union {

	struct {
	   long lKey;
       IFactorTable *pFT;
	 } Self;

	struct {
	   long lKeyTopic, lKeyGost;
       IGostTable *pGT;
	   short iSlot;
	 } Gost;

	} m_Lnk;

#pragma pack(push, 1)  
  typedef struct
   {
     float m_fWeight, m_fValuation;
	 FTSlotTyp m_ftTyp;
	 bool m_bLocked;

	 long lKey_Self;

	 long lKeyTopic_Gost, lKeyGost_Gost;
	 short iSlot_Gost;

	 void LoadToObj( CFactorSlot& rObj ) throw();	  
	 void LoadFromObj( const CFactorSlot& rObj ) throw();
	  
   } TInternalData;
#pragma pack(pop)
  friend struct TInternalData;
 };

inline bool operator<( const CFactorSlot& rObj1, const CFactorSlot& rObj2 ) throw()
 {
   return rObj1.m_bsComponentName.operator<( (BSTR)rObj2.m_bsComponentName );
 }

inline void CFactorSlot::TInternalData::LoadToObj( CFactorSlot& rObj ) throw()
 {
   rObj.m_fWeight = m_fWeight;   
   rObj.m_fValuation = m_fValuation;
   rObj.m_ftTyp = m_ftTyp;
   rObj.m_bLocked = m_bLocked;
   if( m_ftTyp == FT_Self )
	 rObj.m_Lnk.Self.lKey = lKey_Self;
   else if( m_ftTyp == FT_Gost )
	 rObj.m_Lnk.Gost.lKeyTopic  = lKeyTopic_Gost,
	 rObj.m_Lnk.Gost.lKeyGost   = lKeyGost_Gost,
	 rObj.m_Lnk.Gost.iSlot  = iSlot_Gost;
 }
inline void CFactorSlot::TInternalData::LoadFromObj( const CFactorSlot& rObj ) throw()
 {
   m_fWeight = rObj.m_fWeight;   
   m_fValuation = rObj.m_fValuation;
   m_ftTyp = rObj.m_ftTyp;
   m_bLocked = rObj.m_bLocked;
   if( rObj.m_ftTyp == FT_Self )
	 lKey_Self = rObj.m_Lnk.Self.lKey;
   else if( rObj.m_ftTyp == FT_Gost )
	 lKeyTopic_Gost = rObj.m_Lnk.Gost.lKeyTopic,
	 lKeyGost_Gost = rObj.m_Lnk.Gost.lKeyGost,
	 iSlot_Gost = rObj.m_Lnk.Gost.iSlot;
 }


typedef vector<CFactorSlot> VEC_CFactorSlot;
typedef VEC_CFactorSlot::iterator IT_VEC_CFactorSlot;
typedef VEC_CFactorSlot::const_iterator CIT_VEC_CFactorSlot;


/////////////////////////////////////////////////////////////////////////////
// CFactorTable
class ATL_NO_VTABLE CFactorTable : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CFactorTable, &CLSID_FactorTable>,
	public ISupportErrorInfo,
	public IDispatchImpl<IFactorTable, &IID_IFactorTable, &LIBID_AWPLUGINLib>,
	public IPersistStreamInit,
	public IStCollItem
{

friend class CIntCreator;

public:
	CFactorTable() throw()
	 {
	 }

	HRESULT FinalConstruct() throw()
	 {
	   m_lKey = -1;	   
       m_os = OS_None;       
	   m_lSign = 0;
	   m_bDefaultCU = true;
       m_fWeightSumm = m_fValuationSumm = m_fAveWeightedSumm = 0;
	   m_bCalculated = false;
	   return S_OK;
	 }
	void Modify() throw()
	 {
	   if( m_os == OS_None ) m_os = OS_Updated;
	 }


//DECLARE_REGISTRY_RESOURCEID(IDR_FACTORTABLE)
DECLARE_NOT_AGGREGATABLE(CFactorTable)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CFactorTable)
	COM_INTERFACE_ENTRY(IFactorTable)
	COM_INTERFACE_ENTRY(IStCollItem)	
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)

	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid) throw();

// IPersistStreamInit
	STDMETHOD(GetClassID)(CLSID *pClassID) throw()
	 {
	   /*ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::GetClassID\n"));
	   *pClassID = GetObjectCLSID();
	   return S_OK;*/
	   return E_NOTIMPL;
	 }	
	STDMETHOD(IsDirty)() throw()
	 {
	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::IsDirty\n"));		
	   return m_os != OS_None ? S_OK : S_FALSE;
	 }

	STDMETHOD(Load)(LPSTREAM pStm) throw();	
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty) throw();
	STDMETHOD(InitNew)() throw();
	
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER FAR* /* pcbSize */) throw()
	 {
	   ATLTRACENOTIMPL(_T("IPersistStreamInit::GetSizeMax"));
	   return E_NOTIMPL;
	 }


	
public:			
		
//IStCollItem 
	STDMETHOD(get_Key)(/*[out, retval]*/ long *pVal) throw();	
	STDMETHOD(put_Key)(/*[out, retval]*/ long NewVal) throw();	
	STDMETHOD(get_Status)(/*[out, retval]*/ ObjStatus *pVal) throw();
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_Name)(/*[out, retval]*/ BSTR NewVal) throw();
	STDMETHOD(get_Sign)(/*[out, retval]*/ long *pVal) throw();	
	STDMETHOD(put_Sign)(/*[out, retval]*/ long NewVal) throw();	
	STDMETHOD(get_Storage)( /*[out, retval]*/ IStorage** Stg ) throw();
	STDMETHOD(put_DefaultCU)( /*[in]*/VARIANT_BOOL bMode ) throw();
    STDMETHOD(get_DefaultCU)( /*[out,retval]*/VARIANT_BOOL* pbMode ) throw();
	/*[local]*/BSTR STDMETHODCALLTYPE NameByRef( void );


// IFactorTable
	STDMETHOD(NormalizeWeights)() throw();
	STDMETHOD(GetRefToGost)(/*[in]*/long KeyTopic, /*[in]*/long KeyGost, /*[out, retval]*/VARIANT* RefSlot) throw();
	STDMETHOD(GetRefToFTable)(/*[in]*/long Key, /*[out, retval]*/VARIANT* RefSlot) throw();
	STDMETHOD(HandsOffObjects)() throw();
	STDMETHOD(NSlotGost)(/*[in]*/long ZeroIndex, /*[out, optional]*/long* KeyTopic, /*[out, optional]*/long* KeyGost, /*[out, optional]*/IGostTable** IGT) throw();
	STDMETHOD(NSlotSelf)(/*[in]*/long ZeroIndex, /*[out, optional]*/long* Key, /*[out, optional]*/IFactorTable** IFT) throw();
	STDMETHOD(NBindSlot)(/*[in]*/long ZeroIndex, /*[in]*/IDispatch* Object) throw();
	STDMETHOD(NAssSlotType)(/*[in]*/long ZeroIndex, /*[in]*/FTSlotTyp NewTyp, /*[in, optional]*/ long K1, /*[in, optional]*/ long K2) throw();
		
	STDMETHOD(CalcValues)(/*[out, optional]*/VARIANT*  fValuationSumm, /*[out, optional]*/VARIANT*  fAveWeightedSumm) throw();	
	STDMETHOD(get_IsWeightsCorrect)(/*in,out,optional*/VARIANT* Summ,/*[out, retval]*/ VARIANT_BOOL *pVal) throw();
	STDMETHOD(SetWeights)(/*[in, optional]*/VARIANT StartIdx) throw();
    STDMETHOD(get_AllComponentsIsAssigned)( /*[out, retval]*/ VARIANT_BOOL* bRes ) throw();	
		
	STDMETHOD(get_NWeight)(/*[in]*/long ZeroIndex,  /*[out, retval]*/ float *pVal) throw();	
	STDMETHOD(put_NWeight)(/*[in]*/long ZeroIndex,  /*[out, retval]*/ float newVal) throw();	
	STDMETHOD(get_NComment)(/*[in]*/long ZeroIndex,  /*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_NComment)(/*[in]*/long ZeroIndex,  /*[in]*/ BSTR newVal) throw();
	STDMETHOD(get_NLingvoVal)(/*[in]*/long ZeroIndex, /*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_NLingvoVal)(/*[in]*/long ZeroIndex,  /*[in]*/ BSTR Val) throw();
	STDMETHOD(get_NComponentName)(/*[in]*/long ZeroIndex,  /*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_NComponentName)(/*[in]*/long ZeroIndex,  /*[in]*/ BSTR newVal) throw();
	STDMETHOD(get_NumberSlots)(/*[out, retval]*/ long *pVal) throw();	
	STDMETHOD(RemoveSlots)(/*[in, optional, defaultvalue(-1)]*/long Number, /*[in, optional, defaultvalue(-1)]*/long IdxStart) throw();
	STDMETHOD(InsertSlots)(/*[in]*/long Number, /*[in, optional, defaultvalue(-1)]*/long IdxStart) throw();
	
	

	STDMETHOD(get_WeightSumm)(/*[out, retval]*/ float *pVal) throw();
	//STDMETHOD(put_WeightSumm)(/*[in]*/ float newVal) throw();
    STDMETHOD(get_ValuationSumm)(/*[out, retval]*/ VARIANT *pfVal) throw();
	//STDMETHOD(put_ValuationSumm)(/*[in]*/ float newVal) throw();
	STDMETHOD(get_AveWeightedSumm)(/*[out, retval]*/  VARIANT *pfVal) throw();
	//STDMETHOD(put_AveWeightedSumm)(/*[in]*/ float newVal) throw();

	STDMETHOD(get_NAveWeighted)(/*[in]*/long ZeroIndex,  /*[out, retval]*/ VARIANT *pfVal) throw();
	//STDMETHOD(put_NAveWeighted)(/*[in]*/long ZeroIndex,  /*[in]*/ float newVal) throw();
	STDMETHOD(get_NValuation)(/*[in]*/long ZeroIndex,  /*[out, retval]*/ VARIANT *pfVal) throw();	
	STDMETHOD(put_NValuation)(/*[in]*/long ZeroIndex,  /*[in]*/ VARIANT *pfVal) throw();	
	STDMETHOD(get_NSlotType)(long ZeroIndex,/*[out, retval]*/ FTSlotTyp *pVal) throw();

	STDMETHOD(get_NSlotValue)(/*[in]*/long ZeroIndex, /*[out, retval]*/ long *pVal) throw();
	STDMETHOD(put_NSlotValue)(/*[in]*/long ZeroIndex, /*[in]*/ long newVal) throw();

	STDMETHOD(get_NLocked)(/*[in]*/long ZeroIndex, /*[out, retval]*/ VARIANT_BOOL *pVal) throw();
	STDMETHOD(put_NLocked)(/*[in]*/long ZeroIndex, /*[in]*/ VARIANT_BOOL newVal) throw();


	bool operator==( const CFactorTable& rObj )  throw()
	 {
	   return CmpBStrNoCase( m_bsName, rObj.m_bsName );
	 }

	

protected:
   long m_lKey;
   ObjStatus m_os;
   long m_lSign;
   bool m_bDefaultCU;   

   CComBSTR m_bsName;
   VEC_CFactorSlot m_vec;

   float m_fWeightSumm;   
   float m_fValuationSumm, m_fAveWeightedSumm;

   bool m_bCalculated;

   HRESULT CheckArgs( long Number, long IdxStart ) throw();
   HRESULT CheckIdx( long lIdx ) throw();   
   void UncalcN( long ZeroIndex )
	{ 
	  if( ZeroIndex != -1 ) m_vec[ ZeroIndex ].Uncalc();
	  m_bCalculated = false;
	}
};

#endif //__FACTORTABLE_H_
