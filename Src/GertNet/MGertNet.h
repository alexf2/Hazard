// MGertNet.h : Declaration of the CMGertNet

#ifndef __MGERTNET_H_
#define __MGERTNET_H_

#include "resource.h"       // main symbols
#include "GertNetCP.h"
#include "PassErr.h"

#include "CollFac.h"
#include "CollLingvo.h"
#include "CRndFunction.h"
#include "CSimplMatrix.h"

#include "CollFunctors.h"
#include "Fac64NM.h"
#include "TNodeSfTree.h"

#include "map"
#include "vector"
#include "set"
using namespace std;



typedef vector< CComPtr<ISafetyPrecaution> > TV_SF;
typedef TV_SF::iterator IT_TV_SF;
struct TSampleData
 {

   typedef enum {
      TC_Money,
	  TC_DQ
	} TCmpMode;

   TSampleData(): m_ocSumm( 0, 0 ), m_dSummDQ( 0 )
	{
	  m_mode = TC_Money;
	  m_shOverWrap = 0;
	}
   ~TSampleData()
	{
	}

   TSampleData( const TSampleData& rD )
	{
	  this->operator=( rD );
	}
   TSampleData& operator=( const TSampleData& rD )
	{ 
	  m_vec = rD.m_vec;
	  m_ocSumm = rD.m_ocSumm;
	  m_dSummDQ = rD.m_dSummDQ;
	  m_mode = rD.m_mode;
	  m_shOverWrap = rD.m_shOverWrap;
      return *this;
	}


   bool operator==( const TSampleData& rD )
	{
      if( m_mode == TC_Money )
	    return  m_ocSumm == rD.m_ocSumm;
	  else
	    return m_dSummDQ == rD.m_dSummDQ;

	  //return m_dSummDQ == rD.m_dSummDQ;
	}

   bool operator<( const TSampleData& rD )
	{
      if( m_mode == TC_Money )
	    return  m_ocSumm > rD.m_ocSumm;
	  else
	    return m_dSummDQ < rD.m_dSummDQ;

	  //return m_dSummDQ < rD.m_dSummDQ;
	}

   void SummMoney();
   void SummDQ();
   void Summ();
	
   HRESULT GetCollection( ICollSF** ppColl );

   TV_SF m_vec;
   COleCurrency m_ocSumm;
   double m_dSummDQ;
   TSampleData::TCmpMode m_mode;
   short m_shOverWrap;
 };



typedef auto_ptr<TSampleData> AP_TSampleData;

struct TSD_less : binary_function<AP_TSampleData, AP_TSampleData, bool> {
	bool operator()(const AP_TSampleData& _X, const AP_TSampleData& _Y) const
	 {
	   return _X->operator<( *_Y.get() );
	 }
 };

typedef multiset< AP_TSampleData, TSD_less > TS_SetRes;
typedef TS_SetRes::iterator IT_TS_SetRes;

void DumpSetRes( TS_SetRes& rS ); 

#define TNCNT __int64

#define NUMBER_VENTILS 4

#define NUMBER_STATES  8

#define C_B 0
#define C_C 1
#define C_D 2
#define C_E 3
#define C_F 4
#define C_G 5
#define C_H 6
#define C_I 7



 //InSummI, InSummII,  InSumIII,  InSummIV;


struct LongMin { long operator()() { return LONG_MIN; } };
struct LongMax { long operator()() { return LONG_MAX; } };
struct DblMin { long operator()() { return DBL_MIN; } };
struct DblMax { long operator()() { return 1.7976931348623158e+308; } };



/////////////////////////////////////////////////////////////////////////////
// CMGertNet
class ATL_NO_VTABLE CMGertNet : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMGertNet, &CLSID_MGertNet>,
	public ISupportErrorInfo,
	public IDispatchImpl<IMGertNet, &IID_IMGertNet, &LIBID_GERTNETLib>,
	public IPersistStorage,
	public IProvideClassInfo2Impl<&CLSID_MGertNet, &DIID__IGertNetEvents, &LIBID_GERTNETLib>,	
	public IConnectionPointContainerImpl<CMGertNet>,
	public CProxy_IGertNetEvents< CMGertNet >,
	public IClonable,
	public IMGertNet_Debug
{

friend class TOptStep;
friend class TOptInteger;

public:
	CMGertNet():
	   m_evCancel( FALSE, TRUE ), 
	   m_evInternalEndCalc( FALSE, TRUE ), 
	   m_tfFunc( otlMakeFunctor(*this, &CMGertNet::MakeWork), COINIT_MULTITHREADED ),
	   m_tfFuncCDQ( otlMakeFunctor(*this, &CMGertNet::MakeWork2), COINIT_MULTITHREADED ),
	   m_tfFuncOpt( otlMakeFunctor(*this, &CMGertNet::MakeWorkOpt), COINIT_MULTITHREADED )
	 {	  	   	   
	 }
	~CMGertNet()
	 {	   
	 }
	
	HRESULT FinalConstruct();

	void FinalRelease()
	 {	   
	   m_bFinalize = true;

	   ResumeThread( m_tfFunc.m_hthread );
	   ResumeThread( m_tfFuncCDQ.m_hthread );
	   ResumeThread( m_tfFuncOpt.m_hthread );

	   HANDLE hW[ 3 ]  = { m_tfFunc.m_hthread, m_tfFuncCDQ.m_hthread, m_tfFuncOpt.m_hthread };
	   WaitForMultipleObjects( sizeof(hW)/sizeof(hW[0]), hW, TRUE, INFINITE );

	   m_spCollFac.Release();
       m_spCollEnum.Release();
	 }	

	/*STDMETHOD(FindConnectionPoint)(
  REFIID riid,  //Requested connection point's interface identifier
  IConnectionPoint **ppCP
                //Address of output variable that receives the 
                // IConnectionPoint interface pointer
)
	 {
	 MessageBox( NULL, "uu", "kk", MB_OK );
	 //return IConnectionPointContainerImpl<CMGertNet>::FindConnectionPoint(riid, ppCP);
	 return S_OK;
	 }
	STDMETHOD(EnumConnectionPoints)(
  IEnumConnectionPoints **ppEnum  //Address of output variable that 
                                  // receives the IEnumConnectionPoints 
                                  // interface pointer
)
	 {
	   MessageBox( NULL, "uu2", "kk", MB_OK );
	  //return IConnectionPointContainerImpl<CMGertNet>::EnumConnectionPoints(ppEnum);
	   return S_OK;
	 }*/
 
	 

DECLARE_REGISTRY_RESOURCEID(IDR_MGERTNET)
DECLARE_NOT_AGGREGATABLE(CMGertNet)
DECLARE_GET_CONTROLLING_UNKNOWN()

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_CONNECTION_POINT_MAP(CMGertNet)	
	CONNECTION_POINT_ENTRY(DIID__IGertNetEvents)
END_CONNECTION_POINT_MAP()


BEGIN_COM_MAP(CMGertNet)
	COM_INTERFACE_ENTRY(IMGertNet)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)	
	COM_INTERFACE_ENTRY(IClonable)
	COM_INTERFACE_ENTRY(IMGertNet_Debug)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


public:																								
	STDMETHOD(get_ModuleConfig)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_ModuleConfig)(/*[in]*/ BSTR newVal);
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppMGN);	

//IMGertNet_Debug
	STDMETHOD(GetPtsForFactor)( /*[in]*/BSTR bsFShortName, /*[out, retval]*/SAFEARRAY** pparrPoints );
	STDMETHOD(GetDensityForFactor)( /*[in]*/BSTR bsFShortName, /*[out, retval]*/SAFEARRAY** pparrPoints );
	STDMETHOD(TestFunc)(/*[in]*/ BSTR bsFac, /*[in]*/ double dX, /*[out, retval]*/ double* dY);

// IMGertNet
	STDMETHOD(get_VProbability)(/*[in]*/ short Ventil, /*[out, retval]*/ double *pVal);
	STDMETHOD(put_VProbability)(/*[in]*/ short Ventil, /*[in]*/ double newVal);
	STDMETHOD(get_VParamIndistinct)(/*[in]*/ short Ventil, /*[out, retval]*/ double *pVal);
	STDMETHOD(put_VParamIndistinct)(/*[in]*/ short Ventil, /*[in]*/ double newVal);
	STDMETHOD(get_VParam)(/*[in]*/ short Ventil, /*[out, retval]*/ long *pVal);
	STDMETHOD(put_VParam)(/*[in]*/ short Ventil, /*[in]*/ long newVal);
	STDMETHOD(get_NFormat)(/*[in]*/NFormatTyp nf, /*[out, retval]*/ BSTR *pVal);

	STDMETHOD(get_NDigitsOther)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_NDigitsOther)(/*[in]*/ short newVal);
	STDMETHOD(get_NDigitsPr)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_NDigitsPr)(/*[in]*/ short newVal);
	STDMETHOD(get_NFormatOther)(/*[out, retval]*/ NumberFormat *pVal);
	STDMETHOD(put_NFormatOther)(/*[in]*/ NumberFormat newVal);

	STDMETHOD(get_NFormatPr)(/*[out, retval]*/ NumberFormat *pVal);
	STDMETHOD(put_NFormatPr)(/*[in]*/ NumberFormat newVal);
	STDMETHOD(get_IntegralAccuracy)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_IntegralAccuracy)(/*[in]*/ double newVal);
	STDMETHOD(get_UnionThreshold)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_UnionThreshold)(/*[in]*/ double newVal);
	STDMETHOD(get_NDiv)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_NDiv)(/*[in]*/ short newVal);

	STDMETHOD(get_ImitateCount)(/*[out, retval]*/ DATE* pVal);
	STDMETHOD(get_OptCount)(/*[out, retval]*/ DATE *pVal);
	STDMETHOD(get_RngCount)(/*[out, retval]*/ DATE *pVal);

	STDMETHOD(get_ModuleProgID)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_ModuleProgID)(/*[in]*/ BSTR newVal);

	STDMETHOD(TstApplyTCH)(/*[in]*/IFactor* pFac, /*[in]*/IFChange* pFC, /*[out,retval]*/TrustLevel* pLvl);
	STDMETHOD(get_StatIn)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_StatIn)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_StatOn)(/*[out, retval]*/ TStatFlag *pVal);
	STDMETHOD(put_StatOn)(/*[in]*/ TStatFlag newVal);

	STDMETHOD(get_TotalIterStr)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(get_CurrIterStr)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(get_RndBase)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_RndBase)(/*[in]*/ long newVal);

	STDMETHOD(get_UserProp)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_UserProp)(/*[in]*/ long newVal);

	STDMETHOD(Revert)();
	STDMETHOD(Snapshot)(/*[in]*/VARIANT_BOOL bMake);

	STDMETHOD(get_MoneyForSF)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_MoneyForSF)(/*[in]*/ CURRENCY newVal);
	STDMETHOD(get_AverageDamage)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_AverageDamage)(/*[in]*/ CURRENCY newVal);

    STDMETHOD(get_RemovLingvoOverwrap)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_RemovLingvoOverwrap)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_OptimizationType)(/*[out, retval]*/ OptType *pVal);
	STDMETHOD(put_OptimizationType)(/*[in]*/ OptType newVal);
	

	STDMETHOD(get_CalcSpeed)(/*[out, retval]*/ ThrdPriority *pVal);
	STDMETHOD(put_CalcSpeed)(/*[in]*/ ThrdPriority newVal);
	STDMETHOD(get_IsDirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);

	STDMETHOD(GetModelStat)(short shIdx, TypANodesStat* pNS, SAFEARRAY** ppVal, VARIANT_BOOL bFullGet);
	STDMETHOD(get_LastCalcError)(/*[out, retval]*/ BSTR *pVal);

	STDMETHOD(CalibrateModel)(/*[in]*/double dP1, /*[in]*/double dP2, /*[in]*/double dP3, /*[in]*/double dP4);
	STDMETHOD(get_TimeWork3)(/*[out, retval]*/ DATE *pVal);
	STDMETHOD(get_OptimizResultsGetAndClear)(/*[out, retval]*/ SAFEARRAY** ppVal);
	STDMETHOD(get_IsRunningOpt)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_IsRunning2)(/*[out, retval]*/ VARIANT_BOOL *pVal);

	STDMETHOD(OptimizeSelectSF)(/*[in]*/OptTask otTask, /*[in]*/ICollSF *pCollSF, CURRENCY cyMax, double dDeltaQ, double dP0, short shNRetAlternate );
	STDMETHOD(get_TimeWork2)(/*[out, retval]*/ DATE *pVal);

	STDMETHOD(CalcDeltaQ)(ICollSF* pCollSF, long lK, long lN, short shMedVal);
	STDMETHOD(ChangeDirty)(/*[in]*/ VARIANT_BOOL bFlDirty);

	STDMETHOD(GetInfSampleKIdx)(/*[in]*/WCHAR cState, /*[out]*/double* pMin, /*[out]*/double* pMax, /*[out]*/double* pMx, /*[out]*/double* pDx);	
	STDMETHOD(GetInfSampleK)(WCHAR cState,/*[out]*/long* pMin, /*[out]*/long* pMax, /*[out]*/double* pMx, /*[out]*/double* pDx);

	STDMETHOD(get_IsRunning)(/*[out, retval]*/ VARIANT_BOOL *pVal);	

	STDMETHOD(GetEnumForFactor)(/*[int]*/ IFactor* pF, /*[out, retval]*/ ILingvoEnum** ppLE);
	STDMETHOD(GetFactorPresent)(/*[in]*/ IFactor* pF, /*[out,optional]*/ VARIANT* strDescription, /*[out,optional]*/ VARIANT* strTrustLvl, /*[out,optional]*/ VARIANT* dVal);
	STDMETHOD(GetFactorPresentForName)(/*[in]*/ BSTR bsShortName, /*[out]*/ BSTR* pbsDescription, /*[out]*/ BSTR* pbsTrustLvl, /*[out]*/ double* pdVal);
	STDMETHOD(get_NotifyStep)(/*[out, retval]*/ short *pVal);
	STDMETHOD(put_NotifyStep)(/*[in]*/ short newVal);

	STDMETHOD(get_SampleN)(/*[in]*/ long lIndex, /*[out, retval]*/ SAFEARRAY** pparrSmpl );
	STDMETHOD(get_N)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_N)(long pVal);
	STDMETHOD(get_K)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_K)(long pVal);
	STDMETHOD(get_NumberStates)(/*[out, retval]*/ short *pVal)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   if( !pVal ) return E_POINTER;
	   *pVal = NUMBER_STATES; //B, C, D, E, F, G, H, I
	   return S_OK;
	 }
	STDMETHOD(get_TimeWork)(/*[out, retval]*/ DATE *pVal);
	STDMETHOD(get_Dx)(/*[in]*/ WCHAR cState, /*[out, retval]*/ double *pVal);
	STDMETHOD(get_Mx)(/*[in]*/ WCHAR cState, /*[out, retval]*/ double *pVal);
	  //cState is: B, C, D, E, F, G, H, I
	STDMETHOD(get_Counter)(/*[in]*/ WCHAR cState, /*[out, retval]*/ long *pVal);	
		

	STDMETHOD(Cancel)();
	STDMETHOD(Run)(/*[in]*/long lNExperience, /*[in]*/long lNRunInOne, /*[in]*/VARIANT_BOOL bResetModel, /*[in]*/long lRndBase, VARIANT_BOOL bPrepareOnly);

	STDMETHOD(get_RunMode)(/*[out, retval]*/ ModelType *pVal);
	STDMETHOD(put_RunMode)(/*[in]*/ ModelType newVal);
	STDMETHOD(get_NotifyWnd)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_NotifyWnd)(/*[in]*/ long newVal);

	STDMETHOD(get_Factors)(/*[out, retval]*/ ICollFac* *pVal);
	STDMETHOD(get_Enumerators)(/*[out, retval]*/ ICollLingvo* *pVal);

	STDMETHOD(ApplySF)(/*[in]*/ ICollSF* pSF, VARIANT_BOOL* pbOverWrap);

	

//IPersistStorage
	STDMETHOD(GetClassID)( CLSID *pClassID )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::GetClassID\n"));
	   *pClassID = GetObjectCLSID();
	   return S_OK;
	 }
	STDMETHOD(IsDirty)(void);
	STDMETHOD(InitNew)(IStorage* pStor)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::InitNew\n"));

	   return InternalIN( pStor );
	 }
	STDMETHOD(Load)(IStorage* pStorage);	
    STDMETHOD(Save)(IStorage* pStorage, BOOL fSameAsLoad);
	
	STDMETHOD(SaveCompleted)(IStorage* /* pStorage */)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::SaveCompleted\n"));
		return S_OK;
	}
	STDMETHOD(HandsOffStorage)(void)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::HandsOffStorage\n"));
		return S_OK;
	}

protected:
  COleDateTime m_dtSta, m_dtEnd, m_dtSta2, m_dtEnd2, m_dtSta3, m_dtEnd3;
  long m_K, m_N, m_lRndBase;
  //long** m_ppSampleN; //size is NUMBER_STATES*m_K
  //long m_larrCounters[ NUMBER_STATES ];
  CSimpleMatrix<long, __int64, LongMin, LongMax> m_smSampleK;
  CSMOutParams<double> m_arropSampleK[ NUMBER_STATES ];

  //CSimpleMatrix<double, double, DblMin, DblMax> m_smSampleIdx;
  //CSMOutParams<double> m_arropSampleIdx[ NUMBER_STATES ];
  TSampleStat m_ssSampleK[ NUMBER_STATES ];
    
  

  double m_darrMx[ NUMBER_STATES ], m_darrDx[ NUMBER_STATES ];

  long m_lInTest1, m_lInTest2, m_lInTest3, m_lInTest4;
  double m_dInTest1D, m_dInTest2D, m_dInTest3D, m_dInTest4D;
  double m_dArrPVent[ 4 ];

  CComBSTR m_bsModuleProgID, m_bsModuleCnf;

  bool m_bRemovLingvoOverwrap;
  OptType m_otOptimizationType;
  

  bool m_bRequiresSave;

  CComPtr<ICollFac> m_spCollFac;
  CComPtr<ICollLingvo> m_spCollEnum;

  ModelType m_mtModelType;
  HWND m_hwNotify;
  short m_shNotifyStep;

  HRESULT __fastcall InternalIN( IStorage* );
  HRESULT __fastcall InitCollections();
  HRESULT __fastcall GetFactorInfo( IFactor* pF, BSTR *pbsDescription, BSTR *pbsTrustLvl, double *pdVal);
  HRESULT __fastcall Internal_GetEnumForFactor( IFactor *pF, ILingvoEnum **ppLE );

  
  void __fastcall Reset_arrRF();
  HRESULT Update_arrF();
  HRESULT Bind_arrRF();
  CRndFunction m_arrRF[ NUMBER_SINKS ];

  CEvent m_evCancel;
  CEvent m_evInternalEndCalc; //using by multi calc
  bool __fastcall CheckCancel()
   {
     return  ::WaitForSingleObject(m_evCancel, 0) != WAIT_TIMEOUT;	 
   }
  bool __fastcall CheckRunning()
   {
     return CheckRunningThread( m_tfFunc.m_hthread );
   }
  bool __fastcall CheckRunning2()
   {
     return CheckRunningThread( m_tfFuncCDQ.m_hthread );
   }
  bool __fastcall CheckRunningOpt()
   {
     return CheckRunningThread( m_tfFuncOpt.m_hthread );
   }
  bool __fastcall CheckRunningThread( HANDLE h )
   {
     if( h == NULL ) 
	   return false;

	 if( ::WaitForSingleObject(h, 0) != WAIT_TIMEOUT )
	   return false;

	 //CONTEXT ctx; ctx.ContextFlags = CONTEXT_FULL;
	 //return GetThreadContext( m_tfFunc.m_hthread, &ctx ) == FALSE;
	 DWORD dwCnt = SuspendThread( h );
	 ResumeThread( h );

     return dwCnt == 0;
   }
  StingrayOTL::COtlThreadFunction m_tfFunc;
  StingrayOTL::COtlThreadFunction m_tfFuncCDQ;
  StingrayOTL::COtlThreadFunction m_tfFuncOpt;
  CComPtr<ICollSF> m_spCollSFKeep;

  void MakeWork();
  bool __fastcall Work_CancelHelper();
  bool __fastcall Work_CancelHelper3();
  bool m_bC3Notified, m_bCNotified;

  void  MakeWork2();
  void  MakeWorkOpt();
  short m_shMedVal;
  HRESULT __fastcall InternalRun( long lNExperience, long lNRunInOne, VARIANT_BOOL bResetModel, long lRndBase, VARIANT_BOOL bPrepareOnly );

  void (__fastcall CMGertNet::*m_pfnNotify)(NotifyMsg, void*);
  void __fastcall Notify_PostMessage(NotifyMsg, void*);
  void __fastcall Notify_Callback(NotifyMsg, void*);          
  TNCNT m_i64NotifyStepAbs;

//using by OptimizeSelectSF
  int m_iNSelectMax;
  __int64 m_i64TotalOptCycles;
  OptTask m_otTask;
  COleCurrency m_cyMax;
  double m_dDeltaQ;
  short m_shNRetAlternate;
  __int64 m_i64NotifyStepAbsOpt;
  CComPtr<ICollSF> m_spCollSF_Opt;
  TS_SetRes m_sOptResults;

  void __fastcall ModelImitate();
  void __fastcall ModelAnalytic();  
  void __fastcall ModelAnalytic2();
  void __fastcall SetCalibrateError( const double dbl );  
  void __fastcall SetGenericError( const HRESULT hr );
  basic_string<wchar_t> m_bsError;

  double m_dP1, m_dP2, m_dP3, m_dP4; //ventil's Ps (used by calibrate)
  bool m_bCalibrateAnalytic;  

  TANodesStat m_ansStat[ NUMBER_STATES ];
  VEC_TOutNumber2 m_vton[ NUMBER_VENTILS ];
  SET_TOutNumber2 m_ston[ NUMBER_STATES ];

  void __fastcall ResetOverwraps();

  COleCurrency m_cyMoneyForSF;
  COleCurrency m_cyAverageDamage;

  VEC_TFacSnapshot m_vtfs;

  long m_lUser;

  HRESULT __fastcall ApplyVectorSF( TV_SF&, VARIANT_BOOL* pbOverWrap );
  bool __fastcall MakeOnePass( TSampleData& rDta, const double dP0, VARIANT_BOOL* pbOver );
  void __fastcall IncreaseAccuracyOfResult();

  double m_d_P0;
  __int64 m_i64CurrIterOpt;
  TStatFlag m_sfStatOn;

  void __fastcall CalcSampleK();
  bool m_bStatIn;

  void __fastcall InitStateArr();
  void __fastcall AddInSumm( long* pB, long* pC, long* pD, long* pE, long* pF, long* pG, long* pH, long* pI );
  void __fastcall AddPVal( SET_TOutNumber2& rS, const double dPVal );

  __int64 m_i64Count, m_i64Total;
  __int64 m_i64CountRng, m_i64TotalRng;

  double m_dIntegralAccuracy;
  double m_dUnionThreshold;
  short m_sNDiv;
  NumberFormat m_nfFormatPr, m_nfFormatOther;
  short m_sNDigitsPr, m_sNDigitsOther;

  bool m_bFinalize;

  CComBSTR m_bsFmtPr, m_bsFmtOther;
  void __fastcall InitFmts();

#pragma pack(push, 1)  
  typedef struct
   {
     double m_dIntegralAccuracy;
	 double m_dUnionThreshold;
	 short m_sNDiv;

	 NumberFormat m_nfFormatPr, m_nfFormatOther;
	 short m_sNDigitsPr, m_sNDigitsOther;

	 long m_K, m_N, m_lRndBase;

	 long m_lInTest[ 4 ];
     double m_dInTestD[ 4 ];
	 double m_dArrPVent[ 4 ];

	 void LoadToObj( CMGertNet& rObj );	  
	 void LoadFromObj( const CMGertNet& rObj );
	  
   } TInternalData;
#pragma pack(pop)
  friend struct TInternalData;
  
};

inline void CMGertNet::TInternalData::LoadToObj( CMGertNet& rObj )
 {
   rObj.m_dIntegralAccuracy = m_dIntegralAccuracy;
   rObj.m_dUnionThreshold = m_dUnionThreshold;
   rObj.m_sNDiv = m_sNDiv;
   rObj.m_nfFormatPr = m_nfFormatPr;
   rObj.m_nfFormatOther = m_nfFormatOther;
   rObj.m_sNDigitsPr = m_sNDigitsPr;
   rObj.m_sNDigitsOther = m_sNDigitsOther;

   rObj.m_K = m_K;
   rObj.m_N = m_N;
   rObj.m_lRndBase = m_lRndBase;

   rObj.m_lInTest1 = m_lInTest[ 0 ];
   rObj.m_lInTest2 = m_lInTest[ 1 ];
   rObj.m_lInTest3 = m_lInTest[ 2 ];
   rObj.m_lInTest4 = m_lInTest[ 3 ];   

   rObj.m_dInTest1D = m_dInTestD[ 0 ];
   rObj.m_dInTest2D = m_dInTestD[ 1 ];
   rObj.m_dInTest3D = m_dInTestD[ 2 ];
   rObj.m_dInTest4D = m_dInTestD[ 3 ];   

   memcpy( rObj.m_dArrPVent, m_dArrPVent, sizeof m_dArrPVent );
 }
inline void CMGertNet::TInternalData::LoadFromObj( const CMGertNet& rObj )
 {
   m_dIntegralAccuracy = rObj.m_dIntegralAccuracy;
   m_dUnionThreshold = rObj.m_dUnionThreshold;
   m_sNDiv = rObj.m_sNDiv;
   m_nfFormatPr = rObj.m_nfFormatPr;
   m_nfFormatOther = rObj.m_nfFormatOther;
   m_sNDigitsPr = rObj.m_sNDigitsPr;
   m_sNDigitsOther = rObj.m_sNDigitsOther;

   m_K = rObj.m_K;
   m_N = rObj.m_N;
   m_lRndBase = rObj.m_lRndBase;

   m_lInTest[ 0 ] = rObj.m_lInTest1;
   m_lInTest[ 1 ] = rObj.m_lInTest2;
   m_lInTest[ 2 ] = rObj.m_lInTest3;
   m_lInTest[ 3 ] = rObj.m_lInTest4;

   m_dInTestD[ 0 ] = rObj.m_dInTest1D;
   m_dInTestD[ 1 ] = rObj.m_dInTest2D;
   m_dInTestD[ 2 ] = rObj.m_dInTest3D;
   m_dInTestD[ 3 ] = rObj.m_dInTest4D;

   memcpy( m_dArrPVent, rObj.m_dArrPVent, sizeof m_dArrPVent );
 }


void __fastcall LRCut( SET_TOutNumber&, long, IT_SET_TOutNumber, IT_SET_TOutNumber );

#endif //__MGERTNET_H_
