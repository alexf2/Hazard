// MGertNet.cpp : Implementation of CMGertNet
#include "stdafx.h"

//#include <OtlThreads.h>
//using namespace StingrayOTL;

#include "GertNet.h"
#include "MGertNet.h"
#include "OleEnumIterator.h"
#include "CollSF.h"
#include "FChange.h"
#include "ICollFChange.h"
#include "LingvoEnum.h"
#include "HyperIterator.h"
#include "TNodeSFTree.h"
#include "CompaundFiles.h"


#include <tchar.h>


#include <sstream>
#include <string>
#include <iomanip>
#include <map>
using namespace std;


double dUnionThreshold;
short sNDiv;

//состояния
//#define STAT_IDX
//краны
//#define STAT_IDX0

//#define _SIZE_STATISTIC_

static HRESULT CreFormat( const NumberFormat nf, const short sNDigits, CComBSTR& bsResult )
 {
   static const wchar_t const* wcPrefNormal = L"#0.0";   
   static const wchar_t const* wcPrefSc = L"0.0";
   static const wchar_t const* wcSuffNormal = L";;0";
   static const wchar_t const* wcSuffSc = L"e-###;;0";
   static const int iSzPrefNorm = 4;
   static const int iSzPrefSc = 3;
   static const int iSzSuffNorm = 3;
   static const int iSzSuffSc = 8;

   bsResult.Empty(); 
   const int iSzExtra = (nf == NF_Normal ? iSzPrefNorm + iSzSuffNorm:iSzPrefSc+iSzSuffSc);   
   int iSzBody = sNDigits - 1;

   BSTR bsTmp_ = SysAllocStringByteLen(NULL, 2*(iSzExtra + iSzBody + 1));
   if( bsTmp_ == NULL ) return E_OUTOFMEMORY;
   bsResult.Attach( bsTmp_ );
   //memset( bsResult.m_str, 1, 2 * (iSzExtra + iSzBody) );

   //int i1 = *(bsResult.m_str - 1);
   //int i2 = *(bsResult.m_str - 2);
   //int i3 = SysStringLen( bsResult.m_str );

   if( nf == NF_Normal )
	{
	  memcpy( bsResult.m_str, wcPrefNormal, sizeof(wchar_t) * iSzPrefNorm );
	  BSTR bs = bsResult.m_str + iSzPrefNorm;
	  for( int i = iSzBody; i > 0; --i, *bs++ = L'#' );
	  memcpy( bs, wcSuffNormal, sizeof(wchar_t) * iSzSuffNorm );
	}
   else
	{
	  wcsncpy( bsResult.m_str, wcPrefSc, iSzPrefSc );
	  BSTR bs = bsResult.m_str + iSzPrefSc;
	  for( int i = iSzBody; i > 0; --i, *bs++ = L'#' );	  	  
	  wcsncpy( bs, wcSuffSc, iSzSuffSc );
	}   

   *(bsResult.m_str + iSzExtra + iSzBody) = 0;

   //wchar_t ww[ 1024 ];
   //swprintf( ww, L"%s", bsResult.m_str );

   return S_OK;
 }

 
void __fastcall WaitThrdSleeped( HANDLE h )
 {
   DWORD dwCnt;
   do {
	 if( ::WaitForSingleObject(h, 0) != WAIT_TIMEOUT ) break;
	 dwCnt = SuspendThread( h );
	 ResumeThread( h );
	 if( dwCnt == 0 )
	  {
	    Sleep( 20 );
		continue;
	  }
	 else break;
	} while( 1 );
 }

struct TCmpSB_BS
 {
   TCmpSB_BS( BSTR p ):m_p( p )
	{
	}
   int __fastcall operator()( TSincBnd& rBnd )
	{
	  return wcscmp( rBnd.m_bs, m_p ) == 0;
	}

   BSTR m_p;
 };

/////////////////////////////////////////////////////////////////////////////
// CMGertNet

HRESULT CMGertNet::FinalConstruct()
 {
   //m_ppSampleN = NULL;

   m_bFinalize = false;

   m_dIntegralAccuracy = 0.1;
   m_dUnionThreshold = 0.05;
   m_sNDiv = 5;

   m_nfFormatOther = m_nfFormatPr = NF_Normal;
   m_sNDigitsPr = 15; m_sNDigitsOther = 4;
   InitFmts();

   m_i64Total = m_i64Count = 0;
   m_i64CountRng = m_i64TotalRng = 0;

   m_bCalibrateAnalytic = false;
   m_sfStatOn = TSF_No;

   m_bStatIn = false;

   memset( m_ansStat, 0, sizeof(m_ansStat) );

   m_bRequiresSave = 0;
   m_lRndBase = -1;
   m_N = m_K = 0;
   m_lUser = 0;
   m_bC3Notified = false;
   m_bCNotified = false;

   m_dtSta = m_dtEnd = COleDateTime::GetCurrentTime();
   m_dtEnd2 = m_dtSta2 = COleDateTime::GetCurrentTime();
   m_dtEnd3 = m_dtSta3 = COleDateTime::GetCurrentTime();
   
   
   m_hwNotify = NULL;


   for( int i = 0; i < NUMBER_SINKS; ++i )
	 m_arrRF[ i ].Reset(),
	 m_arrRF[ i ].m_bsShortName_Factor = arrbsFacBnd[ i ].m_bs;

   

   return S_OK;
 }

typedef map<WCHAR, short> TM_WI;
static TM_WI::value_type vtarr[ NUMBER_STATES ] = 
	{ 
	  TM_WI::value_type( L'B', 0 ),
	  TM_WI::value_type( L'C', 1 ),
	  TM_WI::value_type( L'D', 2 ),
	  TM_WI::value_type( L'E', 3 ),
	  TM_WI::value_type( L'F', 4 ),
	  TM_WI::value_type( L'G', 5 ),
	  TM_WI::value_type( L'H', 6 ),
	  TM_WI::value_type( L'I', 7 )
	};
static TM_WI mWI( &vtarr[0], &vtarr[NUMBER_STATES] );

static short __fastcall StateToIndex( WCHAR wc )
 {   
   TM_WI::iterator it = mWI.find( wc );
   return it == mWI.end() ? -1:it->second;
 };

static WCHAR __fastcall IndexToState( short sIdx )
 {   
   TM_WI::iterator it = mWI.begin();
   for( ; it != mWI.end(); ++it )
	  if( it->second == sIdx ) return it->first;

   return L'B' - 1;
 };

STDMETHODIMP CMGertNet::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IMGertNet,
		&IID_IPersistStorage
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if( InlineIsEqualGUID(*arr[i],riid) )
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CMGertNet::get_Enumerators(ICollLingvo **pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = NULL;
	return m_spCollEnum.CopyTo( pVal );	
}

STDMETHODIMP CMGertNet::get_Factors(ICollFac **pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	
	 *pVal = NULL;
	return m_spCollFac.CopyTo( pVal );	
}

STDMETHODIMP CMGertNet::get_NotifyWnd(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = (long)m_hwNotify;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NotifyWnd(long newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( CheckRunning() || CheckRunning2() || CheckRunningOpt() )
	{
      Error( L"CMGertNet::Run: is already running", IID_IMGertNet, E_FAIL );   
      return E_FAIL;
	}

	m_hwNotify = (HWND)newVal;	

	return S_OK;
}

STDMETHODIMP CMGertNet::get_RunMode(ModelType *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_mtModelType;	

	return S_OK;
}

STDMETHODIMP CMGertNet::put_RunMode(ModelType newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( CheckRunning() || CheckRunning2() )
	 {
       Error( L"CMGertNet::put_RunMode: is already running", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

    if( newVal != MT_Imitate && 
	    newVal != MT_ImitateIndistinct &&
		newVal != MT_Analytical &&
		newVal != MT_AnalyticalIndistinct &&
		newVal != MT_Analytical2
	  )
	 {
       Error( L"CMGertNet::put_RunMode: invalid arg", IID_IMGertNet, E_INVALIDARG );   
       return E_INVALIDARG;
	 }

	m_mtModelType = newVal;
	m_bRequiresSave = true;
	Reset_arrRF();

	return S_OK;
}



HRESULT __fastcall CMGertNet::InternalRun( long lNExperience, long lNRunInOne, VARIANT_BOOL bResetModel, long lRndBase, VARIANT_BOOL bPrepareOnly )
 {   
    
	m_evCancel.ResetEvent();

	
	if( lNExperience < 1 || lNRunInOne < 1 ) return E_INVALIDARG;
	if( !m_spCollFac || !m_spCollEnum )
	 {
	   PassError( L"CMGertNet::Run: Object MGertNet is not initialized", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	   return E_FAIL;
	 }

	long lCnt1 = 0, lCnt2 = 0;
	m_spCollFac->get_Count( &lCnt1 ), m_spCollEnum->get_Count( &lCnt2 );
	if( lCnt1 < 1 || lCnt2 < 1 )
	 {
	   PassError( L"CMGertNet::Run: Object MGertNet is not initialized", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	   return E_FAIL;
	 }
    

	if( bResetModel == VARIANT_TRUE ) Reset_arrRF();
	//srand( lRndBase == -1 ? time(NULL):lRndBase ); - need in MakeWork
	m_lRndBase = lRndBase;
	m_pfnNotify = ::IsWindow( m_hwNotify ) == TRUE ? &CMGertNet::Notify_PostMessage:&CMGertNet::Notify_Callback;
	 

	if( m_K != lNExperience || !m_smSampleK ) 
	 {
	   m_K = lNExperience, m_N = lNRunInOne;
	   if( !m_smSampleK.Init(NUMBER_STATES, 1) 
		   //!m_smSampleK.Init(NUMBER_STATES, m_K) 
		 )
		{
          Error( L"No memory for instance data", IID_IMGertNet, E_FAIL );
	      return E_FAIL;
		}
	 }	
	else m_N = lNRunInOne;
	
	
	HRESULT hr = Bind_arrRF();
	if( SUCCEEDED(hr) )
	 {
       hr = Update_arrF();
	   if( SUCCEEDED(hr) )
		{

          /*double dVV = m_arrRF[ ST2 ].operator()( 0.8 );
		  basic_string<WCHAR> ss;
		  m_arrRF[ ST2 ].Dump( ss );
		  OutputDebugStringW( ss.c_str() );*/

          m_i64NotifyStepAbs = double(TNCNT(m_K) * TNCNT(m_N)) / 
		    double(100.0) * double(TNCNT(m_shNotifyStep));
		  if( m_i64NotifyStepAbs < 1 ) m_i64NotifyStepAbs = 1;
	      
          if( bPrepareOnly == VARIANT_TRUE ) return S_OK;

		  if( m_tfFunc.Start() == FALSE )
		   {
		     Error( L"CMGertNet::Run: cann't start work", IID_IMGertNet, E_FAIL );   
             hr = E_FAIL;
		   }
		}
	 }
	

	return hr;
 }

STDMETHODIMP CMGertNet::Run(long lNExperience, long lNRunInOne, VARIANT_BOOL bResetModel, long lRndBase, VARIANT_BOOL bPrepareOnly)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	if( CheckRunning() || CheckRunning2() )
	 {
       Error( L"CMGertNet::Run: is already running", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	return InternalRun( lNExperience, lNRunInOne, bResetModel, lRndBase, bPrepareOnly );	
}

STDMETHODIMP CMGertNet::Cancel()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( CheckCancel() )
	 {
	   Error( L"CMGertNet::Cancel: object is already in cancel mode", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }
	m_evCancel.SetEvent();

	return S_OK;
}



STDMETHODIMP CMGertNet::Clone(IUnknown **ppMGN)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( !ppMGN ) return E_POINTER;

	CComObject<CMGertNet>* pObj;
	HRESULT hr = CComObject<CMGertNet>::CreateInstance( &pObj );
	if( FAILED(hr) )
	 {
       PassError( L"CMGertNet::Copy: CComObject<CMGertNet>::CreateInstance", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	CComPtr<IUnknown> spTmp;
	hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
	if( FAILED(hr) )
	 {
	   delete pObj;
       PassError( L"CMGertNet::Copy: CComObject<CMGertNet>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	pObj->m_dtSta = m_dtSta;
	pObj->m_dtEnd = m_dtEnd;
	pObj->m_dtSta2 = m_dtSta2;
	pObj->m_dtEnd2 = m_dtEnd2;
	pObj->m_dtSta3 = m_dtSta3;
	pObj->m_dtEnd3 = m_dtEnd3;

	pObj->m_K = m_K;
	pObj->m_N = m_N;
	pObj->m_lRndBase = m_lRndBase;

    pObj->m_smSampleK = m_smSampleK; 
	//pObj->m_smSampleIdx = m_smSampleIdx;		
	memcpy( pObj->m_ssSampleK, m_ssSampleK, sizeof(m_ssSampleK) );

	pObj->m_cyMoneyForSF = m_cyMoneyForSF;
    pObj->m_cyAverageDamage = m_cyAverageDamage;
	

	//memcpy( pObj->m_larrCounters, m_larrCounters, sizeof(m_larrCounters) );
	memcpy( pObj->m_darrMx, m_darrMx, sizeof(m_darrMx) );
	memcpy( pObj->m_darrDx, m_darrDx, sizeof(m_darrDx) );

	pObj->m_lInTest1 = m_lInTest1;
	pObj->m_lInTest2 = m_lInTest2;
	pObj->m_lInTest3 = m_lInTest3;
	pObj->m_lInTest4 = m_lInTest4;

	pObj->m_dInTest1D = m_dInTest1D;
	pObj->m_dInTest2D = m_dInTest2D;
	pObj->m_dInTest3D = m_dInTest3D;
	pObj->m_dInTest4D = m_dInTest4D;

	memcpy( pObj->m_dArrPVent, m_dArrPVent, sizeof m_dArrPVent );

	pObj->m_bsModuleProgID = m_bsModuleProgID;
	pObj->m_bsModuleCnf = m_bsModuleCnf;

	pObj->m_dIntegralAccuracy = m_dIntegralAccuracy;
	pObj->m_dUnionThreshold = m_dUnionThreshold;
	pObj->m_sNDiv = m_sNDiv;
	pObj->m_nfFormatPr = m_nfFormatPr;

	pObj->m_nfFormatOther = m_nfFormatOther;
	pObj->m_sNDigitsPr = m_sNDigitsPr;
	pObj->m_sNDigitsOther = m_sNDigitsOther;
	pObj->InitFmts();
	

	pObj->m_bRemovLingvoOverwrap = m_bRemovLingvoOverwrap;
	pObj->m_otOptimizationType = m_otOptimizationType;
	

	pObj->m_shNotifyStep = m_shNotifyStep;
	pObj->m_lUser = m_lUser;

	pObj->m_bRequiresSave = FALSE;

	
	CComQIPtr<IClonable> spClon( m_spCollFac.p );
	if( !spClon.p )
	 {	   
	   Error( L"CMGertNet::Copy: CComQIPtr<IClonable>", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }
	CComPtr<IUnknown> spUnkCopy;
	hr = spClon->Clone( &spUnkCopy );
	if( FAILED(hr) )
	 {
       PassError( L"CMGertNet::Copy: cann't clone", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }
	pObj->m_spCollFac.Release();
	hr = spUnkCopy.QueryInterface( &pObj->m_spCollFac.p );
	if( FAILED(hr) )
	 {
       PassError( L"CMGertNet::Copy: QueryInterface", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }


	spClon.Release();
	m_spCollEnum.QueryInterface( &spClon );
	if( !spClon.p )
	 {
	   PassError( L"CMGertNet::Copy: CComQIPtr<IClonable>", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
       return E_FAIL;
	 }
	spUnkCopy = NULL;
	hr = spClon->Clone( &spUnkCopy );
	if( FAILED(hr) )
	 {
       PassError( L"CMGertNet::Copy: cann't clone", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }
	pObj->m_spCollEnum.Release();
	hr = spUnkCopy.QueryInterface( &pObj->m_spCollEnum.p );
	if( FAILED(hr) )
	 {
       PassError( L"CMGertNet::Copy: QueryInterface", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }
	    

	pObj->m_mtModelType = m_mtModelType; 
    pObj->m_hwNotify = m_hwNotify;
	
	*ppMGN = NULL;
	spTmp.CopyTo( ppMGN );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_Counter(WCHAR cState, long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	short shInd = StateToIndex( cState );
	if( shInd == -1 )
	 {
	   PassError( L"IMGertNet.get_Counter: bad index of state", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet  );
       return E_INVALIDARG;
	 }	

	//*pVal = m_larrCounters[ shInd ];

	return S_OK;
}


STDMETHODIMP CMGertNet::get_Mx(WCHAR cState, double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	short shInd = StateToIndex( cState );
	if( shInd == -1 )
	 {
	   PassError( L"IMGertNet.get_Mx: bad index of state", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet  );
       return E_INVALIDARG;
	 }

	*pVal = m_darrMx[ shInd ];

	return S_OK;
}

STDMETHODIMP CMGertNet::get_Dx(WCHAR cState, double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	short shInd = StateToIndex( cState );
	if( shInd == -1 )
	 {
	   PassError( L"IMGertNet.get_Dx: bad index of state", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet  );
       return E_INVALIDARG;
	 }

	*pVal = m_darrDx[ shInd ];

	return S_OK;
}

STDMETHODIMP CMGertNet::get_TimeWork(DATE *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	*pVal = m_dtEnd - m_dtSta;

	return S_OK;
}


STDMETHODIMP CMGertNet::get_K(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_K;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_K( long newVal )
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	
	MPUT_PROPERTY( m_K, newVal );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_N(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_N;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_N( long newVal )
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	MPUT_PROPERTY( m_N, newVal );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_SampleN(long lIndex, SAFEARRAY** pparrSmpl )
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pparrSmpl ) return E_POINTER;
    if( m_K == 0 )
	 {
	   basic_stringstream<WCHAR> strm;
	   if( m_K == 0 ) strm << L"IMGertNet.get_SampleN: no runs";
	   else
	     strm << L"IMGertNet.get_SampleN: index" << lIndex << L"out of range - need (0.." << m_K << L")";
       PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet  );   
	   return E_FAIL;
	 }
	if( lIndex < 0 || lIndex >= NUMBER_STATES )
	 {
	   Error( L"CMGertNet::get_SampleN: index out of bounds", IID_IMGertNet, E_INVALIDARG );
	   return E_INVALIDARG;
	 }

	//SAFEARRAYBOUND bnd = { m_K, 0 };
	SAFEARRAYBOUND bnd = { 1, 0 };
	*pparrSmpl = SafeArrayCreate( VT_I4, 1, &bnd );

	if( !*pparrSmpl )
	 {
	   Error( L"Cann't array allocate", IID_IMGertNet, E_FAIL );
	   return E_FAIL;
	 }

	long* pDta;
	HRESULT hr = SafeArrayAccessData( *pparrSmpl, (void**)&pDta );
	if( FAILED(hr) )
	 {
	   SafeArrayDestroy( *pparrSmpl );
	   *pparrSmpl = NULL;	   
	   Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
	   return E_FAIL;
	 }


	//memcpy( pDta, m_smSampleK[ lIndex ], m_smSampleK.SizeElem() * m_K );
	memcpy( pDta, m_smSampleK[ lIndex ], m_smSampleK.SizeElem() * 1 );

	SafeArrayUnaccessData( *pparrSmpl );

	return S_OK;
}

HRESULT __fastcall CMGertNet::InitCollections()
 {
   m_spCollFac.Release();
   m_spCollEnum.Release();

   CComObject<CCollFac>* pCF;
   HRESULT hr = CComObject<CCollFac>::CreateInstance( &pCF );
   if( FAILED(hr) )
	{		  
	  PassError( L"CComObject<CCollFac>::CreateInstance", hr, GetObjectCLSID(), IID_IMGertNet  );
	  return hr;
	}
   hr = pCF->_InternalQueryInterface( IID_ICollFac, (void**)&m_spCollFac );
   if( FAILED(hr) )
	{
	  delete pCF;
	  PassError( L"_InternalQueryInterface of IID_ICollFac failed", hr, GetObjectCLSID(), IID_IMGertNet  );
	  return hr;
	}

   CComObject<CCollLingvo>* pCL;
   hr = CComObject<CCollLingvo>::CreateInstance( &pCL );
   if( FAILED(hr) )
	{		  
	  PassError( L"CComObject<CCollLingvo>::CreateInstance", hr, GetObjectCLSID(), IID_IMGertNet  );
	  return hr;
	}

   hr = pCL->_InternalQueryInterface( IID_ICollLingvo, (void**)&m_spCollEnum );
   if( FAILED(hr) )
	{
	  delete pCL;
	  PassError( L"_InternalQueryInterface of IID_ICollLingvo failed", hr, GetObjectCLSID(), IID_IMGertNet  );
	  return hr;
	}

   return S_OK;
 }

HRESULT __fastcall CMGertNet::InternalIN( IStorage* )
 {
   m_dIntegralAccuracy = 0.1;
   m_dUnionThreshold = 0.05;
   m_sNDiv = 5;

   m_nfFormatOther = m_nfFormatPr = NF_Normal;
   m_sNDigitsPr = 15; m_sNDigitsOther = 4;
   InitFmts ();
   
   m_lInTest1 = m_lInTest2 = m_lInTest3 = m_lInTest4 = 0;
   m_shNotifyStep = 10;

   m_bsModuleProgID = L"";
   m_bsModuleCnf = L"";

   m_dInTest1D = m_dInTest2D = m_dInTest3D = m_dInTest4D = 0;

   double dPrTmp[] = { 0.5, 0.05, 0.005, 0.0005 };
   memcpy( m_dArrPVent, dPrTmp, sizeof m_dArrPVent );

   m_bRemovLingvoOverwrap = true;
   m_otOptimizationType = OT_Quick;

   m_mtModelType = MT_Imitate;

   m_cyMoneyForSF = COleCurrency( 1500, 0 );
   m_cyAverageDamage = COleCurrency( 10000000, 0 );
   

   HRESULT hr = InitCollections();     
   if( SUCCEEDED(hr) )
	{
	  CComQIPtr<IPersistStorage> sp( m_spCollFac.p );
	  hr = sp->InitNew( NULL );
	  if( SUCCEEDED(hr) )
	   {
	     sp.Release();
		 m_spCollEnum.QueryInterface( &sp );
	     hr = sp->InitNew( NULL );
	   }
	}    

   if( SUCCEEDED(hr) ) 
     m_bRequiresSave = true;

   return hr;
 }

STDMETHODIMP CMGertNet::get_NotifyStep(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( !pVal ) return E_POINTER;
	*pVal = m_shNotifyStep;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NotifyStep(short newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( newVal < 1 || newVal > 99 )
	 {
	   PassError( L"put_NotifyStep: need argument 1..99", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet  );
	   return E_INVALIDARG;
	 }

	m_shNotifyStep = newVal;

	return S_OK;
}


//IPersistStorage
STDMETHODIMP CMGertNet::IsDirty(void)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::IsDirty\n"));		
   

   if( m_bRequiresSave ) return S_OK;
   if( m_spCollFac.p )
	{
	  CComQIPtr<IPersistStorage> spStor( m_spCollFac.p );
	  if( spStor.p && spStor->IsDirty() == S_OK ) return S_OK;
	}

   if( m_spCollEnum.p )
	{
	  CComQIPtr<IPersistStorage> spStor( m_spCollEnum.p );
	  if( spStor.p && spStor->IsDirty() == S_OK ) return S_OK;
	}

   return S_FALSE;
 }

STDMETHODIMP CMGertNet::Load( IStorage* pStorage )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Load\n"));

   if( !pStorage ) return E_POINTER;

   m_bRequiresSave = false;

   CComPtr<IStream> spInfoStream;
   HRESULT hr = pStorage->OpenStream( L"Content_MGertNet", 0,
	 STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spInfoStream );
   if( FAILED(hr) ) 
	  return ReportStgError( hr, L"IStorage.OpenStream", GetObjectCLSID(), IID_IMGertNet );
    /*{         
	  PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_IMGertNet  );   
	  return hr;
	}*/

   
   hr = spInfoStream->Read( &m_mtModelType, sizeof(m_mtModelType), NULL );
   if( SUCCEEDED(hr) )
	{
      hr = spInfoStream->Read( &m_shNotifyStep, sizeof(m_shNotifyStep), NULL );
	  if( SUCCEEDED(hr) )
	   {
		 hr = spInfoStream->Read( &m_bRemovLingvoOverwrap, sizeof(m_bRemovLingvoOverwrap), NULL );
		 if( SUCCEEDED(hr) )
		  {
			hr = spInfoStream->Read( &m_otOptimizationType, sizeof(m_otOptimizationType), NULL );
			if( SUCCEEDED(hr) )
			 {				  
				hr = spInfoStream->Read( &m_cyAverageDamage.m_cur, sizeof(m_cyAverageDamage.m_cur), NULL );
                if( SUCCEEDED(hr) )					  
				  hr = spInfoStream->Read( &m_cyMoneyForSF.m_cur, sizeof(m_cyMoneyForSF.m_cur), NULL );
			 }
		  }
	   }
	}
	
      
   if( FAILED(hr) )
	 return ReportStgError( hr, L"Load", GetObjectCLSID(), IID_IMGertNet );
	

   spInfoStream.Release();

   hr = InitCollections();
   if( FAILED(hr) ) return hr;

   CComPtr<IStorage> spStorage;
   hr = pStorage->OpenStorage( L"Factors collection", NULL, 
	 STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStorage );   

   if( FAILED(hr) )
	 return ReportStgError( hr, L"OpenStorage \"Factors collection\"", GetObjectCLSID(), IID_IMGertNet );	

   CComQIPtr<IPersistStorage> spPS( m_spCollFac.p );

   hr = spPS->Load( spStorage );
   if( SUCCEEDED(hr) )
	{
	  spStorage.Release();
      hr = pStorage->OpenStorage( L"LingvoEnums collection", NULL, 
	    STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStorage );   
	  if( FAILED(hr) )
	    return ReportStgError( hr, L"OpenStorage \"LingvoEnums collection\"", GetObjectCLSID(), IID_IMGertNet );	
	   
      
	  spPS.Release();
	  m_spCollEnum.QueryInterface( &spPS );
      hr = spPS->Load( spStorage ); 
	}

   if( FAILED(hr) ) 
	 return ReportStgError( hr, L"Load", GetObjectCLSID(), IID_IMGertNet );	

   
   hr = pStorage->OpenStream( L"ModuleProgID", 0,
	 STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spInfoStream );
   if( FAILED(hr) ) 
    {         
	  if( HRESULT_FACILITY(hr) == FACILITY_STORAGE && hr == STG_E_FILENOTFOUND )
	    m_bsModuleProgID = L"", m_bsModuleCnf = L"";
	  else
	    return ReportStgError( hr, L"IStorage.OpenStream", GetObjectCLSID(), IID_IMGertNet );		   
	}
   else
	{
	  hr = m_bsModuleProgID.ReadFromStream( spInfoStream );	  
	  if( SUCCEEDED(hr) )
	    hr = m_bsModuleCnf.ReadFromStream( spInfoStream );
	  
	  spInfoStream.Release();
	  if( FAILED(hr) )
	    return ReportStgError( hr, L"IStorage.OpenStream", GetObjectCLSID(), IID_IMGertNet );		   
	}   

   hr = pStorage->OpenStream( L"AlgThreshold", 0,
	 STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spInfoStream );
   if( FAILED(hr) ) 
    {         
	  if( HRESULT_FACILITY(hr) == FACILITY_STORAGE && hr == STG_E_FILENOTFOUND )
	   {
	     m_dIntegralAccuracy = 0.1;
		 m_dUnionThreshold = 0.05;
		 m_sNDiv = 5;

		 m_nfFormatOther = m_nfFormatPr = NF_Normal;
         m_sNDigitsPr = 15; m_sNDigitsOther = 4;
	   }
	  else
	    return ReportStgError( hr, L"IStorage.OpenStream", GetObjectCLSID(), IID_IMGertNet );		   	   
	}
   else
	{
	  CMGertNet::TInternalData th;
	  hr = spInfoStream->Read( &th, sizeof(th), NULL );
	  spInfoStream.Release();
	  if( FAILED(hr) ) 
	    return ReportStgError( hr, L"Load", GetObjectCLSID(), IID_IMGertNet );
	  th.LoadToObj( *this );
	}   
	
   return S_OK;
 }
STDMETHODIMP CMGertNet::Save( IStorage* pStorage, BOOL fSameAsLoad )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Save\n"));
   

   if( !pStorage ) return E_POINTER;

   CComPtr<IStream> spInfoStream;
   HRESULT hr = pStorage->CreateStream( L"Content_MGertNet",
	 STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spInfoStream );
   if( FAILED(hr) ) 
	 return ReportStgError( hr, L"IStorage.CreateStream", GetObjectCLSID(), IID_IMGertNet );    

   
   hr = spInfoStream->Write( &m_mtModelType, sizeof(m_mtModelType), NULL );
   if( SUCCEEDED(hr) )	   
	{
      hr = spInfoStream->Write( &m_shNotifyStep, sizeof(m_shNotifyStep), NULL );         	   
      if( SUCCEEDED(hr) )	   
	   {
		 hr = spInfoStream->Write( &m_bRemovLingvoOverwrap, sizeof(m_bRemovLingvoOverwrap), NULL );
		 if( SUCCEEDED(hr) )
		  {
			hr = spInfoStream->Write( &m_otOptimizationType, sizeof(m_otOptimizationType), NULL );
			if( SUCCEEDED(hr) )
			 {				  
			   hr = spInfoStream->Write( &m_cyAverageDamage.m_cur, sizeof(m_cyAverageDamage.m_cur), NULL );
			   if( SUCCEEDED(hr) )					  
				 hr = spInfoStream->Write( &m_cyMoneyForSF.m_cur, sizeof(m_cyMoneyForSF.m_cur), NULL );
			 }
		  }
	   }
	}

   
   if( FAILED(hr) )
	 return ReportStgError( hr, L"WriteToStream", GetObjectCLSID(), IID_IMGertNet );    	

   spInfoStream.Release();

   CComPtr<IStorage> spStorage;
   hr = pStorage->OpenStorage( L"Factors collection", NULL, 
	 STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStorage );
   

   CComQIPtr<IPersistStorage> spPS( m_spCollFac.p );
   if( FAILED(hr) || spPS->IsDirty() == S_OK )
	{      
	  if( !spStorage )
	    hr = pStorage->CreateStorage( L"Factors collection", 
	      STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spStorage );	   

	  if( FAILED(hr) )
	    return ReportStgError( hr, L"OpenStorage (Factors collection)", GetObjectCLSID(), IID_IMGertNet );	   

	  hr = spPS->Save( spStorage, fSameAsLoad );
	  if( FAILED(hr) ) 
	    return ReportStgError( hr, L"Save", GetObjectCLSID(), IID_IMGertNet );
	}


   spStorage.Release();

   hr = pStorage->OpenStorage( L"LingvoEnums collection", NULL, 
	 STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStorage );
   

   spPS.Release();
   m_spCollEnum.QueryInterface( &spPS );   
   if( FAILED(hr) || spPS->IsDirty() == S_OK )
	{      
	  if( !spStorage )
	    hr = pStorage->CreateStorage( L"LingvoEnums collection", 
	      STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spStorage );
	   //pStorage->CreateStorage( L"LingvoEnums collection", 
	   //STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStorage );

	  if( FAILED(hr) )
	    return ReportStgError( hr, L"CreateStorage", GetObjectCLSID(), IID_IMGertNet );	   

	  hr = spPS->Save( spStorage, fSameAsLoad );	  
	  if( FAILED(hr) ) 
	    return ReportStgError( hr, L"Save", GetObjectCLSID(), IID_IMGertNet );	   
	}


   hr = pStorage->CreateStream( L"ModuleProgID",
	 STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spInfoStream );
   if( FAILED(hr) ) 
	 return ReportStgError( hr, L"IStorage.CreateStream", GetObjectCLSID(), IID_IMGertNet );	       

   hr = m_bsModuleProgID.WriteToStream( spInfoStream );
   if( SUCCEEDED(hr) )
     hr = m_bsModuleCnf.WriteToStream( spInfoStream );   

   spInfoStream.Release();
   if( FAILED(hr) ) 
	 return ReportStgError( hr, L"Save", GetObjectCLSID(), IID_IMGertNet );	       
   
   

   hr = pStorage->CreateStream( L"AlgThreshold",
	 STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spInfoStream );
   if( FAILED(hr) ) 
	 return ReportStgError( hr, L"IStorage.CreateStream", GetObjectCLSID(), IID_IMGertNet );	       
    

   CMGertNet::TInternalData th; th.LoadFromObj( *this );
   hr = spInfoStream->Write( &th, sizeof(th), NULL );
   if( FAILED(hr) ) 
	 return ReportStgError( hr, L"Save", GetObjectCLSID(), IID_IMGertNet );

   spInfoStream.Release();


   m_bRequiresSave = false;

   return S_OK;
 }



STDMETHODIMP CMGertNet::GetFactorPresent(/*[in]*/ IFactor* pF, /*[out,optional]*/ VARIANT* strDescription, /*[out,optional]*/ VARIANT* strTrustLvl, /*[out,optional]*/ VARIANT* dVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pF ) return E_POINTER;

	BSTR *pbsDescription, *pbsTrustLvl;
	double *pdVal;
    HRESULT hr = S_OK;

	pbsDescription = GetBSTRDst( strDescription, hr );
	if( FAILED(hr) )
	  Error( L"Bad parametr 'strDescription' type. String is expected", IID_IMGertNet, E_INVALIDARG );
	else
	 {
	   pbsTrustLvl = GetBSTRDst( strTrustLvl, hr );	   
	   if( FAILED(hr) )
		 Error( L"Bad parametr 'strTrustLvl' type. String is expected", IID_IMGertNet, E_INVALIDARG );
	   else
		{
		  pdVal = GetDoubleDst( dVal, hr );	   
		  if( FAILED(hr) )
		    Error( L"Bad parametr 'dVal' type. Double is expected", IID_IMGertNet, E_INVALIDARG );
		  else
		   hr = GetFactorInfo( pF, 
	              pbsDescription, 
	              pbsTrustLvl, 
	              pdVal );
		}
	 }

	return hr == S_FALSE ? S_OK:hr;     			
}

HRESULT __fastcall CMGertNet::GetFactorInfo( IFactor* pF, BSTR *pbsDescription, BSTR *pbsTrustLvl, double *pdVal)
 {
    ILingvoEnum* pILE;
    HRESULT hr = Internal_GetEnumForFactor( pF, &pILE );
	if( SUCCEEDED(hr) )
	 {
	   CComObject<CFactor>* pFactor = static_cast<CComObject<CFactor>*>( pF );
       CComObject<CLingvoEnum>* pLE = static_cast<CComObject<CLingvoEnum>*>( pILE );

	   if( pFactor->m_sValue < 0 || pFactor->m_sValue >= pLE->GetCount() )
		{
		  basic_stringstream<WCHAR> strm;
		  strm << L"CMGertNet::ApplySF: bad value of factor [" << pFactor->m_sValue << L"]";
		  PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
		  return E_FAIL;  
		}

	   if( pbsDescription )
	     *pbsDescription = pLE->arrLD[ pFactor->m_sValue ].m_cbsMark.Copy();
	   if( pdVal )
	     *pdVal = pLE->arrLD[ pFactor->m_sValue ].m_dQuality;
	   

	   if( pbsTrustLvl )
		{
		  TrustLevel tlLvl = pFactor->m_tr;
		  if( tlLvl == TL_Low )
			*pbsTrustLvl = SysAllocString( L"Низкий" );
		  else if( tlLvl == TL_Normal )
			*pbsTrustLvl = SysAllocString( L"Средний" );
		  else if( tlLvl == TL_High )
			*pbsTrustLvl = SysAllocString( L"Высокий" );
		  else *pbsTrustLvl = SysAllocString( L"<Unknown>" );

		  if( *pbsTrustLvl == NULL ) hr = E_OUTOFMEMORY;
		}
	 }
	return hr;
    

	/*long lInd;
	short shQ;
	TrustLevel tlLvl;

	hr = pF->get_IDEnum( &lInd );
	if( SUCCEEDED(hr) )
	 {
	   hr = pF->get_Value( &shQ );
	   if( SUCCEEDED(hr) )
		{	
	      hr = pF->get_TrustLvl( &tlLvl );
          if( SUCCEEDED(hr) )
		   {
	         CComPtr<ILingvoEnum> spIL;
	         hr = m_spCollEnum->Item( lInd, &spIL );
             if( SUCCEEDED(hr) )
			  {
	            hr = spIL->get_Mark( shQ, pbsDescription );
				if( SUCCEEDED(hr) )
				 { 
	               hr = spIL->get_Quality( shQ, pdVal );

				   if( tlLvl == TL_Low )
					 *pbsTrustLvl = SysAllocString( L"Низкий" );
				   else if( tlLvl == TL_Normal )
					 *pbsTrustLvl = SysAllocString( L"Нормальный" );
				   else *pbsTrustLvl = SysAllocString( L"Высокий" );
				 }
			  }
		   }
		}
	 }

	if( FAILED(hr) )
	 {
       PassError( L"CMGertNet::GetFactorPresent", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	return S_OK;*/
 }

STDMETHODIMP CMGertNet::GetFactorPresentForName(BSTR bsShortName, BSTR *pbsDescription, BSTR *pbsTrustLvl, double *pdVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pbsTrustLvl || !pdVal ) return E_POINTER;

	HRESULT hr;

	CComPtr<IFactor> spF;
	hr = m_spCollFac->Item( bsShortName, &spF );
	if( FAILED(hr) ) return hr;

	return GetFactorInfo( spF, pbsDescription, pbsTrustLvl, pdVal );
}

/*STDMETHODIMP CMGertNet::ApplySF(ICollSF *pSF)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( !pSF ) return E_POINTER;

   HRESULT  hr;
   

   COleEnumIterator< ICollSF, 
	 IEnumVARIANT, CComVariant, 
	 CComPtr<ISafetyPrecaution>, 
	 VARIANTExtractor< CComPtr<ISafetyPrecaution> > > oeiSF( pSF );

   hr = oeiSF;
   if( SUCCEEDED(hr) )
	{
	  CComPtr<ISafetyPrecaution> spSP;
      for( spSP = oeiSF++; (HRESULT)oeiSF == S_OK; spSP = oeiSF++ )
	   {
	     CComPtr<IICollFChange> spFChanges;
         HRESULT  hr = spSP->get_FChanges( &spFChanges );

		 COleEnumIterator< IICollFChange, 
		   IEnumVARIANT, CComVariant, 
		   CComPtr<IFChange>, VARIANTExtractor<CComPtr<IFChange> > > oeiChanges( spFChanges );

		 hr = oeiChanges;
		 if( SUCCEEDED(hr) )
		  {
            CComPtr<IFChange> spFChange;
			for( spFChange = oeiChanges++; (HRESULT)oeiChanges == S_OK; spFChange = oeiChanges++ )
			 {
			   CComBSTR bsNF;
			   spFChange->get_NameFactor( &bsNF );
			   CComPtr<IFactor> spFac;
               hr = m_spCollFac->Item( bsNF, &spFac );
			   if( FAILED(hr) )
				{
				  basic_stringstream<WCHAR> strm;
				  strm << L"CMGertNet::ApplySF: no such factor [" << (BSTR)bsNF << L"]";
				  PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			      return E_FAIL;  
				}

			   long lIDEnum;
			   spFac->get_IDEnum( &lIDEnum );
			   CComPtr<ILingvoEnum> spLEnum;
			   hr = m_spCollEnum->Item( lIDEnum, &spLEnum );
			   if( FAILED(hr) )
				{
				  basic_stringstream<WCHAR> strm;
				  strm << L"CMGertNet::ApplySF: no such LingvoEnum [" << lIDEnum << L"]";
				  PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			      return E_FAIL;  
				}

			   spLEnum->get_Count( &lIDEnum );
			   short shVal;
			   spFac->get_Value( &shVal );
			   short shDelta;
			   spFChange->get_Delta( &shDelta );

			   shVal += shDelta;
			   if( shVal < 0 ) shVal = 0;
			   else if( shVal >= lIDEnum ) shVal = lIDEnum - 1; 

			   spFac->put_Value( shVal );
			 }
			if( (HRESULT)oeiChanges != S_FALSE && (HRESULT)oeiChanges != S_OK )
			 {
			   PassError( L"CMGertNet::ApplySF: cann't enumerate IICollFChange: ", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			   return E_FAIL;  
			 }
		  }
	   }
	  if( (HRESULT)oeiSF != S_FALSE && (HRESULT)oeiSF != S_OK )
	   {
	     PassError( L"CMGertNet::ApplySF: cann't enumerate ICollSF: ", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
         return E_FAIL;  
	   }
	}

   
   return hr;
 }*/

void __fastcall CMGertNet::ResetOverwraps()
 {
   CComObject<CCollFac>* pC = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );
   CComObject<CCollFac>::TMAP::iterator it( pC->m_coll.begin() );
   for( ; it != pC->m_coll.end(); ++it )
	{
      CComObject<CFactor>* pF = static_cast<CComObject<CFactor>*>( it->second );
	  pF->ResetOverWrap();
	}
 }

static TrustLevel ModifyTrustLvl( TrustLevel tl, TrustChange tc )
 {
   TrustLevel tlRet = tl;
   switch( tc )
	{
	  case TR_NoChange:
	    break;
	  case TR_MinusOne:
	    tlRet = TrustLevel(short(tlRet) - 1);
		break;
	  case TR_MinusTwo:
	    tlRet = TrustLevel(short(tlRet) - 2);
		break;
	  case TR_PlusOne:
	    tlRet = TrustLevel(short(tlRet) + 1);
		break;
	  case TR_PlusTwo:
	    tlRet = TrustLevel(short(tlRet) + 2);
		break;
	  case TR_SetLow:
	    tlRet = TL_Low;
		break;
	  case TR_SetNormal:
	    tlRet = TL_Normal;
		break;
	  case TR_SetHigh:
	    tlRet = TL_High;
		break;
	}

   if( tlRet < TL_Low ) tlRet = TL_Low;
   else if( tlRet > TL_High ) tlRet = TL_High;

   return tlRet;
 }

static void ApplyFCToFactor( CComObject<CFChange>* pFC, CComObject<CFactor>* pFactor, bool& bFlOverWrap )
 {
   pFactor->m_sValue += pFC->m_shDelta;
   if( pFactor->m_sValue < 0 ) 
	{
	  pFactor->AddOwerWrap();
	  pFactor->m_sValue = 0;
	  bFlOverWrap = true;
	}
   else if( pFactor->m_sValue >= CLingvoEnum::GetCount() )
	{
	  pFactor->AddOwerWrap();
	  pFactor->m_sValue = CLingvoEnum::GetCount() - 1;
	  bFlOverWrap = true;
	}
   

   pFactor->m_tr = ModifyTrustLvl( pFactor->m_tr, pFC->m_tcTrustChange );
 }


HRESULT __fastcall CMGertNet::ApplyVectorSF( TV_SF& vec, VARIANT_BOOL* pbOverWrap )
 {
   bool bFlOverWrap = false;
   ResetOverwraps();

   CComObject<CCollLingvo>* pCollEnum = static_cast<CComObject<CCollLingvo>*>( m_spCollEnum.p );
   IT_TV_SF it( vec.begin() );
   for( ; it != vec.end(); ++it )
	{
//typedef vector< CComPtr<ISafetyPrecaution> > TV_SF;
      CComObject<CSafetyPrecaution>* pSafety = static_cast<CComObject<CSafetyPrecaution>*>( it->p );
	  CComObject<CICollFChange>* pCollFC = static_cast<CComObject<CICollFChange>*>( pSafety->m_spCollFChange.p );

	  CComObject<CICollFChange>::TMAP::iterator itSta( pCollFC->m_coll.begin() );
      CComObject<CICollFChange>::TMAP::iterator itEnd( pCollFC->m_coll.end() );
	  for( ;itSta != itEnd; ++itSta )
	   {
		 CComObject<CFChange>* pFC = static_cast<CComObject<CFChange>*>( itSta->second );
		 CComObject<CCollFac>* pCollF = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );

		 CComObject<CCollFac>::TMAP::iterator itFac =
		   pCollF->m_coll.find( Chk(pFC->m_bsName) );
		 if( itFac == pCollF->m_coll.end() )
		  {
			basic_stringstream<WCHAR> strm;
			strm << L"CMGertNet::ApplyVectorSF: no such factor [" << (BSTR)Chk(pFC->m_bsName) << L"]";
			PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			return E_FAIL;  
		  }
		 CComObject<CFactor>* pFactor = static_cast<CComObject<CFactor>*>( itFac->second );


		 ApplyFCToFactor( pFC, pFactor, bFlOverWrap );		 
	   }
	}

   if( pbOverWrap ) *pbOverWrap = (bFlOverWrap ? VARIANT_TRUE:VARIANT_FALSE);

   return S_OK;
 }



STDMETHODIMP CMGertNet::ApplySF(ICollSF *pSF, VARIANT_BOOL* pbOverWrap)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( !pSF ) return E_POINTER;
   bool bFlOverWrap = false;

   HRESULT  hr;

   ResetOverwraps();

   CComObject<CCollSF>* pCollSF = static_cast<CComObject<CCollSF>*>( pSF );
   CComObject<CCollSF>::TMAP::iterator it1( pCollSF->m_coll.begin() );
   CComObject<CCollSF>::TMAP::iterator it2( pCollSF->m_coll.end() );

   CComObject<CCollLingvo>* pCollEnum = static_cast<CComObject<CCollLingvo>*>( m_spCollEnum.p );

   for( ; it1 != it2; ++it1 )
	{
	  CComObject<CSafetyPrecaution>* pSafety = static_cast<CComObject<CSafetyPrecaution>*>( it1->second );
	  CComObject<CICollFChange>* pCollFC = static_cast<CComObject<CICollFChange>*>( pSafety->m_spCollFChange.p );

	  CComObject<CICollFChange>::TMAP::iterator itSta( pCollFC->m_coll.begin() );
      CComObject<CICollFChange>::TMAP::iterator itEnd( pCollFC->m_coll.end() );
	  for( ;itSta != itEnd; ++itSta )
	   {
		 CComObject<CFChange>* pFC = static_cast<CComObject<CFChange>*>( itSta->second );
		 CComObject<CCollFac>* pCollF = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );

		 CComObject<CCollFac>::TMAP::iterator itFac =
		   pCollF->m_coll.find( Chk(pFC->m_bsName) );
		 if( itFac == pCollF->m_coll.end() )
		  {
			basic_stringstream<WCHAR> strm;
			strm << L"CMGertNet::ApplySF: no such factor [" << (BSTR)Chk(pFC->m_bsName) << L"]";
			PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			return E_FAIL;  
		  }
		 CComObject<CFactor>* pFactor = static_cast<CComObject<CFactor>*>( itFac->second );


		 /*CComObject<CCollLingvo>::TMAP::iterator itEnum =
		   pCollEnum->m_coll.find( pFactor->m_lIDEnum );
		 if( itEnum == pCollEnum->m_coll.end() )
		  {
			basic_stringstream<WCHAR> strm;
			strm << L"CMGertNet::ApplySF: no such LingvoEnum [" << pFactor->m_lIDEnum << L"]";
			PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			return E_FAIL;  
		  }

		 CComObject<CLingvoEnum>* pLE = static_cast<CComObject<CLingvoEnum>*>( itEnum.second );
		 if( pFactor->m_sValue < 0 || pFactor->m_sValue >= pLE->GetCount() )
		  {
			basic_stringstream<WCHAR> strm;
			strm << L"CMGertNet::ApplySF: bad value of factor [" << pFactor->m_sValue << L"]";
			PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
			return E_FAIL;  
		  }*/

		 //pLE->arrLD[ pFactor->m_sValue ]

		 ApplyFCToFactor( pFC, pFactor, bFlOverWrap );		 		 
	   }
	}

   if( pbOverWrap ) *pbOverWrap = (bFlOverWrap ? VARIANT_TRUE:VARIANT_FALSE);

   return S_OK;
 }



STDMETHODIMP CMGertNet::GetEnumForFactor(IFactor *pF, ILingvoEnum **ppLE)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	HRESULT hr = Internal_GetEnumForFactor( pF, ppLE );
	if( SUCCEEDED(hr) ) (*ppLE)->AddRef();
	return hr;
}

HRESULT __fastcall CMGertNet::Internal_GetEnumForFactor( IFactor *pF, ILingvoEnum **ppLE )
 {
   if( !pF || !ppLE ) return E_POINTER;

   CComObject<CFactor>* pFactor = static_cast<CComObject<CFactor>*>( pF );

   CComObject<CCollLingvo>* pCollEnum = static_cast<CComObject<CCollLingvo>*>( m_spCollEnum.p );
   CComObject<CCollLingvo>::TMAP::iterator itEnum =
	    pCollEnum->m_coll.find( pFactor->m_lIDEnum );
   if( itEnum == pCollEnum->m_coll.end() )
	{
      basic_stringstream<WCHAR> strm;
	  strm << L"Internal_GetEnumForFactor: no such LingvoEnum [" << pFactor->m_lIDEnum << L"]";
	  PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	  return E_FAIL;  
	}

   *ppLE = itEnum->second;

   return S_OK;
 }



void __fastcall CMGertNet::Reset_arrRF()
 {
   for( int i = 0; i < NUMBER_SINKS; ++i )
	 m_arrRF[ i ].Reset();

   m_smSampleK.Clear();      
   //m_smSampleIdx.Clear();
 }
HRESULT CMGertNet::Bind_arrRF()
 {
   for( int i = 0; i < NUMBER_SINKS; ++i )
	 if( m_arrRF[ i ].IsUnbound() )
	  {
	    CComPtr<IFactor> spFac;
	    HRESULT hr = m_spCollFac->Item( m_arrRF[ i ].m_bsShortName_Factor, &spFac );
		if( FAILED(hr) ) return hr;		 
        CComObject<CFactor>* pFactor = static_cast<CComObject<CFactor>*>( spFac.p );

		//double dTmp;
        hr = GetFactorInfo( spFac, NULL, NULL, &m_arrRF[ i ].m_dFacVal );
		//m_arrRF[ i ].m_dFacVal = dTmp;

		if( FAILED(hr) ) return hr;		 

		m_arrRF[ i ].m_pFactor = pFactor;
		//m_arrRF[ i ].m_sValue = pFactor->m_sValue;
		//m_arrRF[ i ].m_tr = pFactor->m_tr; 
	  }

   return S_OK;
 }


HRESULT CMGertNet::Update_arrF()
 {
   HRESULT hr;
   for( int i = 0; i < NUMBER_SINKS; ++i )
	{
	  CRndFunction& rF = m_arrRF[ i ];
	  if( rF.m_sValue != rF.m_pFactor->m_sValue ||
	      rF.m_tr != rF.m_pFactor->m_tr || 
		  memcmp(rF.arrshInd, rF.m_pFactor->arrshInd, sizeof(rF.arrshInd)) != 0 ||
           
		  rF.m_pdt != rF.m_pFactor->m_pdt ||
          rF.m_fPlacement != rF.m_pFactor->m_fPlacement ||
		  rF.m_fScale != rF.m_pFactor->m_fScale
		)
	   {
         rF.m_sValue = rF.m_pFactor->m_sValue;
		 rF.m_tr = rF.m_pFactor->m_tr;

         rF.m_pdt = rF.m_pFactor->m_pdt;
         rF.m_fPlacement = rF.m_pFactor->m_fPlacement;
		 rF.m_fScale = rF.m_pFactor->m_fScale;
		 rF.GenDistribution();

		 //double dTmp;
		 hr = GetFactorInfo( rF.m_pFactor, NULL, NULL, &rF.m_dFacVal );
		 //rF.m_dFacVal = dTmp;
		 if( FAILED(hr) ) return hr;
         
         //short* pshIdx = arrbsFacBnd[i].arrshInd;
		 //for( int j = 0; j < NUMBER_IDX && *pshIdx != IDX_TERM; ++j, ++pshIdx );
		  
		 //pshIdx = arrbsFacBnd[i].arrshInd;
		 bool bIndistinct = (m_mtModelType == MT_ImitateIndistinct || m_mtModelType == MT_AnalyticalIndistinct || m_mtModelType == MT_Analytical2) ? true:false;
		 switch( rF.m_pFactor->GetNIdx() )
		  {
		    case 2:
			  rF.Div2( rF.m_pFactor->arrshInd[0], rF.m_pFactor->arrshInd[1], bIndistinct );
			  break;
			case 3:
			  rF.Div3( rF.m_pFactor->arrshInd[0], rF.m_pFactor->arrshInd[1], rF.m_pFactor->arrshInd[2], bIndistinct );
			  break;
			case 4:
			  rF.Div4( rF.m_pFactor->arrshInd[0], rF.m_pFactor->arrshInd[1], rF.m_pFactor->arrshInd[2], rF.m_pFactor->arrshInd[3], bIndistinct );
			  break;
			default:
              PassError( L"CMGertNet::Update_arrF: bad div index", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	          return E_FAIL;
		  };
	   }
	}

   return S_OK;
 }


STDMETHODIMP CMGertNet::GetDensityForFactor( /*[in]*/BSTR bsFShortName, /*[out, retval]*/SAFEARRAY** pparrPoints )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( !pparrPoints ) return E_POINTER;

   for( int i = 0; i < NUMBER_SINKS; ++i )
	 if( m_arrRF[ i ].IsUnbound() || m_arrRF[ i ].m_vec.size() < 1 ) 
	  {
	    PassError( L"Internal chache is not initialized or not updated", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
		return E_FAIL;  
	  }

   TCmpSB_BS dta( bsFShortName );
   TSincBnd* pBnd = find_if( &arrbsFacBnd[0], &arrbsFacBnd[NUMBER_SINKS], dta );
   if( pBnd == &arrbsFacBnd[NUMBER_SINKS] )
	{
	  basic_stringstream<WCHAR> strm;
	  strm << L"No such Factor [" << (BSTR)Chk(bsFShortName) << L"]";
	  PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	  return E_FAIL;  
	}
   
   CRndFunction& rFunc = m_arrRF[ pBnd - &arrbsFacBnd[0] ]; 

   const double dStepDiscr = (rFunc.m_pdt == PDT_Uniform ? 1:0.02), 
	      dX1 = 0, dX2 = 1;
   long iNPts = floor( (dX2 - dX1)/dStepDiscr + 1.0 ); 
   //const int iNPts = dNPts;   

   SAFEARRAYBOUND bnd = { iNPts * 2, 0 };
   *pparrPoints = SafeArrayCreate( VT_R8, 1, &bnd );

   if( !*pparrPoints )
	{
	  Error( L"Cann't array allocate", IID_IMGertNet, E_FAIL );
	  return E_FAIL;
	}

   double* pDta;
   HRESULT hr = SafeArrayAccessData( *pparrPoints, (void**)&pDta );
   if( FAILED(hr) )
	{
	  SafeArrayDestroy( *pparrPoints );
	  *pparrPoints = NULL;	   
	  Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
	  return E_FAIL;
	}


   for( double dX = dX1; iNPts > 0; dX += dStepDiscr, --iNPts )
	 *pDta++ = dX,  *pDta++ = rFunc.Density( dX );

   SafeArrayUnaccessData( *pparrPoints );

   return S_OK;
 }

STDMETHODIMP CMGertNet::GetPtsForFactor( /*[in]*/BSTR bsFShortName, /*[out, retval]*/SAFEARRAY** pparrPoints )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( !pparrPoints ) return E_POINTER;

   for( int i = 0; i < NUMBER_SINKS; ++i )
	 if( m_arrRF[ i ].IsUnbound() || m_arrRF[ i ].m_vec.size() < 1 ) 
	  {
	    PassError( L"Internal chache is not initialized or not updated", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
		return E_FAIL;  
	  }

   TCmpSB_BS dta( bsFShortName );
   TSincBnd* pBnd = find_if( &arrbsFacBnd[0], &arrbsFacBnd[NUMBER_SINKS], dta );
   if( pBnd == &arrbsFacBnd[NUMBER_SINKS] )
	{
	  basic_stringstream<WCHAR> strm;
	  strm << L"No such Factor [" << Chk(bsFShortName) << L"]";
	  PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	  return E_FAIL;  
	}
   
   CRndFunction& rFunc = m_arrRF[ pBnd - &arrbsFacBnd[0] ]; 
   
   SAFEARRAYBOUND bnd = { rFunc.m_vec.size() * 2, 0 };
   *pparrPoints = SafeArrayCreate( VT_R8, 1, &bnd );

   if( !*pparrPoints )
	{
	  Error( L"Cann't array allocate", IID_IMGertNet, E_FAIL );
	  return E_FAIL;
	}

   double* pDta;
   HRESULT hr = SafeArrayAccessData( *pparrPoints, (void**)&pDta );
   if( FAILED(hr) )
	{
	  SafeArrayDestroy( *pparrPoints );
	  *pparrPoints = NULL;	   
	  Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
	  return E_FAIL;
	}


   for( i = 0; i < rFunc.m_vec.size(); ++i )
	 *pDta++ = rFunc.m_vec[ i ].dX,  *pDta++ = rFunc.m_vec[ i ].shY;

   SafeArrayUnaccessData( *pparrPoints );

   return S_OK;
 }

#define TRCalc double

static void SampleStat( TSampleStat& rS, double dInSumm )
 {
   rS.dMinVal = min( rS.dMinVal, dInSumm );
   rS.dMaxVal = max( rS.dMaxVal, dInSumm );
   rS.l64Sz++;
   rS.dAvg += dInSumm;
   rS.dM2 += dInSumm * dInSumm;
 }

void __fastcall CMGertNet::InitStateArr()
 {   
#if defined(STAT_IDX) || defined(STAT_IDX0)
   static double darrEps[ NUMBER_STATES ] = {	
	0.5,
	//1,
	 0.5,
	 0.5,
	 0.5,
	 0.5,
	0.5,
	 0.5,
	 0.5
	};
#else
    static double darrEps[ NUMBER_STATES ] = {
	0.00005,
	0.00005,
	0.00005,
	0.00002,
	0.00001,
	0.00001,
	0.00001,
	0.000005	
	}; 
#endif

   for( int i = 0; i < sizeof(m_ston)/sizeof(m_ston[0]); ++i )		  
	{
	  m_ston[ i ].clear();
#if defined(STAT_IDX) || defined(STAT_IDX0)
	  for( double dS = 0; dS < 10000; dS += darrEps[i] ) 	  
#else
	  for( double dS = 0; dS < 1; dS += darrEps[i] )
#endif
	   {
	     TOutNumber2 tnTmp(0, dS, dS + darrEps[i]); tnTmp.m_lCnt = 0;
		 pair<IT_SET_TOutNumber2,bool> pp = m_ston[ i ].insert( tnTmp );		 
	   }	  
	}
 }


void __fastcall CMGertNet::AddPVal( SET_TOutNumber2& rS, const double dPVal )
 {
   pair<IT_SET_TOutNumber2, bool> pRes = rS.insert( TOutNumber2(0, dPVal, dPVal) );
   if( pRes.second == true )
	{
	  /*int yy=1;
	  IT_SET_TOutNumber it = pRes.first;
	  it--, it--, it--, it--;
	  basic_stringstream<wchar_t> strm;
      strm << setprecision( 7 ) << L"----------------" << endl;
	  USES_CONVERSION; 
	  for( int l = 0; it != rS.end() && l < 10; ++l, ++it )
	   {
	     strm << setw(0) << L"(" << setw(7) << it->m_shNumber1 <<
		 setw(0) << L"," << setw(7) << it->m_shNumber2 << L")" << 
		 setw(12) << it->m_tpP << endl;		 		   
	   }
	  OutputDebugString( OLE2CT(strm.str().c_str()) );*/
	}
   else
	{
	  pRes.first->m_lCnt++;
	}
 }

void __fastcall CMGertNet::AddInSumm( long* pB, long* pC, long* pD, long* pE, long* pF, long* pG, long* pH, long* pI )
 {
    AddPVal( m_ston[ C_B ], double(*pB) / double(m_N) );
	AddPVal( m_ston[ C_C ], double(*pC) / double(m_N) );
	AddPVal( m_ston[ C_D ], double(*pD) / double(m_N) );
	AddPVal( m_ston[ C_E ], double(*pE) / double(m_N) );	


	AddPVal( m_ston[ C_F ], double(*pF) / double(m_N) );	
	AddPVal( m_ston[ C_G ], double(*pG) / double(m_N) );	
	AddPVal( m_ston[ C_H ], double(*pH) / double(m_N) );	
	AddPVal( m_ston[ C_I ], double(*pI) / double(m_N) );		
 }

void __fastcall CMGertNet::ModelImitate()
 {


   srand( m_lRndBase == -1 ? time(NULL):m_lRndBase );

   
   if( m_sfStatOn == TSF_Full ) InitStateArr();	
   

   memset( m_ssSampleK, 0, sizeof(m_ssSampleK) );
   for( int iJK = 0; iJK < sizeof(m_ssSampleK)/sizeof(m_ssSampleK[0]); ++iJK )
     m_ssSampleK[ iJK ].dMaxVal = DBL_MIN,
	 m_ssSampleK[ iJK ].dMinVal = DBL_MAX;	 

   double dTst1, dTst2, dTst3, dTst4;
   if( m_mtModelType == MT_Imitate )
	 dTst1 = m_lInTest1,
	 dTst2 = m_lInTest2,
	 dTst3 = m_lInTest3,
	 dTst4 =m_lInTest4;
   else
	 dTst1 = m_dInTest1D,
	 dTst2 = m_dInTest2D,
	 dTst3 = m_dInTest3D,
	 dTst4 =m_dInTest4D;

   DATE dtEnd; 

   TRCalc st2, st4, st6, st8, st10, st12, st14, st16,
     st18, st20, st22, st24, st26, st28, st30, st32,
     st34, st36, st38, st40, st42, st44, st46, st48,
     st50, st52, st54, st56, st58, st60, st62,
	 InSummA, InSummA0, InSummI, InSummC, InSummC0,
	 InSummII, InSummE, InSummE0, InSumIII, InSummG,
	 InSummG0, InSummIV;

   
   //memset( m_darrMx, 0, sizeof(m_darrMx) );
   //memset( m_darrDx, 0, sizeof(m_darrDx) );
   
   //m_smSampleIdx.SetDrt( true );	  	  

   for( iJK = 0; iJK < sizeof(m_arropSampleK)/sizeof(m_arropSampleK[0]); ++iJK )
     m_arropSampleK[ iJK ].m_dx = m_arropSampleK[ iJK ].m_mx = 0,
	 m_arropSampleK[ iJK ].m_min = 1.7976931348623158e+308,
	 m_arropSampleK[ iJK ].m_max = 2.2250738585072014e-308;
   

   short shPercent;
   long *pB = m_smSampleK[ C_B ], *pC = m_smSampleK[ C_C ],
	 *pD = m_smSampleK[ C_D ], *pE = m_smSampleK[ C_E ],
	 *pF = m_smSampleK[ C_F ], *pG = m_smSampleK[ C_G ],
	 *pH = m_smSampleK[ C_H ], *pI = m_smSampleK[ C_I ];
   
   
   TNCNT i64Count = 0, i64Total = TNCNT(m_K) * TNCNT(m_N);
   m_i64Total = i64Total, m_i64Count = 0;
   for( long lK = 0; lK < m_K; ++lK /*++pB, ++pC, ++pD, ++pE, ++pF, ++pG, ++pH, ++pI*/ )
	{	     
	  m_smSampleK.Zero();   
      m_smSampleK.SetDrt( true );

	  for( long lN = 0; lN < m_N; ++lN, m_i64Count = ++i64Count )
	   {
		 /*if( CheckCancel() )
		  {
			m_dtEnd = COleDateTime::GetCurrentTime();
			dtEnd = m_dtEnd - m_dtSta;
			(this->*m_pfnNotify)( NM_OnCancel, &dtEnd );
			//goto L_REENTER_POINT;
			return;
		  }*/
		 if( (i64Count % m_i64NotifyStepAbs) == 0 )
		  { 
			shPercent =  (TNCNT(100) * i64Count) / i64Total;
			(this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
		  }
		 if( Work_CancelHelper() ) return;

         
		 st2 = m_arrRF[ ST2 ](); 
		 st4 = m_arrRF[ ST4 ](); 

		 st6 = m_arrRF[ ST6 ](); 
		 st8 = m_arrRF[ ST8 ](); 

		 st10 = m_arrRF[ ST10 ](); 
		 st12 = m_arrRF[ ST12 ](); 

		 st14 = m_arrRF[ ST14 ](); 
		 st16 = m_arrRF[ ST16 ](); 

		 InSummA = st2 + st4 + st6 + st8 + st10 + st12 + st14 + st16;


		 st18 = m_arrRF[ ST18 ](); 
		 st20 = m_arrRF[ ST20 ](); 

		 st22 = m_arrRF[ ST22 ](); 
		 st24 = m_arrRF[ ST24 ](); 

		 st26 = m_arrRF[ ST26 ](); 
		 st28 = m_arrRF[ ST28 ](); 

		 InSummA0 = st18*st20 + st22 + st24 + st26 + st28;       
         InSummI = InSummA * InSummA0; 		 

		 if( m_bStatIn ) SampleStat( m_ssSampleK[C_B], InSummI );
		 

#if defined(STAT_IDX0)
		 AddPVal( m_ston[ C_B ], InSummI );
#endif

		 if( InSummI <= TRCalc(dTst1) )
		  {
			//m_smSampleIdx( C_B, lK ) += InSummI;
#if defined(STAT_IDX)
		    AddPVal( m_ston[ C_B ], InSummI );
#endif
		    if( !m_bStatIn ) SampleStat( m_ssSampleK[C_B], InSummI );		    
			++*pB; continue;
		  }

#if defined(STAT_IDX) || defined(STAT_IDX0)
		 AddPVal( m_ston[ C_C ], InSummI );
#endif
		  //m_smSampleIdx( C_C, lK ) += InSummI;
		  SampleStat( m_ssSampleK[C_C], InSummI );		    
		  ++*pC;
          InSummC = InSummI;

		  st30 = m_arrRF[ ST30 ](); 
		  st32 = m_arrRF[ ST32 ](); 

		  st34 = m_arrRF[ ST34 ](); 
		  st36 = m_arrRF[ ST36 ](); 

		  st38 = m_arrRF[ ST38 ](); 
		  st40 = m_arrRF[ ST40 ](); 

		  st42 = m_arrRF[ ST42 ](); 
		  st44 = m_arrRF[ ST44 ](); 

		  InSummC0 = (st30 + st32 + st34)*st36 + st38*st40 + st42 + st44; 
          InSummII = InSummC * InSummC0; 

          if( m_bStatIn ) SampleStat( m_ssSampleK[C_D], InSummII );
		  		  
#if defined(STAT_IDX0)
		  AddPVal( m_ston[ C_D ], InSummII );
#endif

		  if( InSummII <= TRCalc(dTst2) )
		   {
			 //m_smSampleIdx( C_D, lK ) += InSummII;
		     if( !m_bStatIn ) SampleStat( m_ssSampleK[C_D], InSummII );		    
#if defined(STAT_IDX)
			 AddPVal( m_ston[ C_D ], InSummII );
#endif
			 ++*pD; continue;
		   }


		  //m_smSampleIdx( C_E, lK ) += InSummII;
#if defined(STAT_IDX) || defined(STAT_IDX0)
		  AddPVal( m_ston[ C_E ], InSummII );
#endif

		  SampleStat( m_ssSampleK[C_E], InSummII );		    
		  ++*pE;
          InSummE = InSummII;

		  st46 = m_arrRF[ ST46 ](); 
		  st48 = m_arrRF[ ST48 ](); 

		  st50 = m_arrRF[ ST50 ](); 
		  st52 = m_arrRF[ ST52 ](); 

		  st54 = m_arrRF[ ST54 ](); 

		  InSummE0 = (st46 + st48)*st50 + st52*st54; 
          InSumIII = InSummE * InSummE0; 

          if( m_bStatIn ) SampleStat( m_ssSampleK[C_F], InSumIII );		    
		  
#if defined(STAT_IDX0)
		  AddPVal( m_ston[ C_F ], InSumIII );
#endif

		  if( InSumIII <= TRCalc(dTst3) )
		   {
			 //m_smSampleIdx( C_F, lK ) += InSumIII;
		     if( !m_bStatIn ) SampleStat( m_ssSampleK[C_F], InSumIII );		    
#if defined(STAT_IDX)
			 AddPVal( m_ston[ C_F ], InSumIII );
#endif
			 ++*pF; continue;
		   }

		  //m_smSampleIdx( C_G, lK ) += InSumIII;
#if defined(STAT_IDX) || defined(STAT_IDX0)
		  AddPVal( m_ston[ C_G ], InSumIII );
#endif
		  SampleStat( m_ssSampleK[C_G], InSumIII );		    
		  ++*pG;
          InSummG = InSumIII;

		  st56 = m_arrRF[ ST56 ](); 
		  st58 = m_arrRF[ ST58 ](); 

		  st60 = m_arrRF[ ST60 ](); 
		  st62 = m_arrRF[ ST62 ](); 

		  InSummG0 = st56*st58 + st60 + st62;
          InSummIV = InSumIII * (st56*st58 + st60 + st62); 


          if( m_bStatIn ) SampleStat( m_ssSampleK[C_H], InSummIV );		    
	
#if defined(STAT_IDX0)
		  AddPVal( m_ston[ C_H ], InSummIV );
#endif		  

		  if( InSummIV <= TRCalc(dTst4) )
		   {
			 //m_smSampleIdx( C_H, lK ) += InSummIV;
		     if( !m_bStatIn ) SampleStat( m_ssSampleK[C_H], InSummIV );		    
#if defined(STAT_IDX)
			 AddPVal( m_ston[ C_H ], InSummIV );
#endif
			 ++*pH; continue;
		   }
		  //m_smSampleIdx( C_I, lK ) += InSummIV;
#if defined(STAT_IDX) || defined(STAT_IDX0)
		  AddPVal( m_ston[ C_I ], InSummIV );
#endif
		  SampleStat( m_ssSampleK[C_I], InSummIV );		    
		  ++*pI;
	   }

#if !defined(STAT_IDX) && !defined(STAT_IDX0)
	  if( m_sfStatOn == TSF_Full ) AddInSumm( pB, pC, pD, pE, pF, pG, pH, pI );
#endif

      m_smSampleK.OutParams0( m_arropSampleK, m_N );
	  
	}

   m_smSampleK.OutParams0_End( m_arropSampleK, m_K );

   CalcSampleK();
   Work_CancelHelper();
 }

void __fastcall CMGertNet::CalcSampleK()
 {
   for( int i = 0; i < sizeof(m_ssSampleK) / sizeof(m_ssSampleK[0]); ++i )
	  if( m_ssSampleK[ i ].l64Sz != 0 )
	   {
	     m_ssSampleK[ i ].dAvg /= double(m_ssSampleK[ i ].l64Sz);		 

		 m_ssSampleK[ i ].dM2 /= double(m_ssSampleK[ i ].l64Sz);
		 m_ssSampleK[ i ].dM2 -= m_ssSampleK[ i ].dAvg * m_ssSampleK[ i ].dAvg;
		 if( m_ssSampleK[ i ].l64Sz > 1 )
		   m_ssSampleK[ i ].dM2 *= double(m_ssSampleK[ i ].l64Sz) / double(m_ssSampleK[ i ].l64Sz- 1);

		 m_ssSampleK[ i ].dM2 = sqrt( m_ssSampleK[ i ].dM2 );
	   }
	  else
	    m_ssSampleK[ i ].dAvg = 
		m_ssSampleK[ i ].dM2  = 
		m_ssSampleK[ i ].dMaxVal = 
		m_ssSampleK[ i ].dMinVal = 0;
 }

bool __fastcall CMGertNet::Work_CancelHelper()
 {
   DATE dtEnd;  
   if( CheckCancel() )
	{
	  if( !m_bCNotified )
	   {
		 m_dtEnd = COleDateTime::GetCurrentTime();
		 dtEnd = m_dtEnd - m_dtSta;
		 m_bCNotified = true;
		 (this->*m_pfnNotify)( NM_OnCancel, &dtEnd );	  		 
	   }
	  return true;
	}
   return false;
 }


void __fastcall CMGertNet::SetCalibrateError( const double dbl )
 {
   basic_stringstream<wchar_t> strm;
   strm << L"Cann't provide such P: " << dbl;
   m_bsError = strm.str().c_str();
   CComBSTR bs( m_bsError.c_str() );
   (this->*m_pfnNotify)( NM_OnErrorCalc, &bs );
 }

void __fastcall CMGertNet::SetGenericError( const HRESULT hr )
 {
   std::basic_stringstream<WCHAR> strm;
   strm << std::setbase(16);
   
   CComPtr<IErrorInfo> spE;
   if( SUCCEEDED(GetErrorInfo(0, &spE)) && spE.p )
	{	  
	  CComBSTR bsSrc, bsDscr;
	  spE->GetDescription( &bsDscr );
	  spE->GetSource( &bsSrc );
	  

	  if( !bsDscr ) bsDscr = L"";
	  if( !bsSrc ) bsSrc = L"";
	  
	  strm << L"COM Error: 0x" << hr << L";\nFrom: \"" << Chk(bsSrc) <<
	   L"\";\nCOM Description: \"" << Chk(bsDscr) << L"\"";
	}
   else	  
     strm << L"Error: 0x" << hr;	

   //CComBSTR bs( strm.
   (this->*m_pfnNotify)( NM_OnErrorCalc, (void*)strm.str().c_str() );
 }

void __fastcall CMGertNet::ModelAnalytic()
 {  
   dUnionThreshold = m_dUnionThreshold;
   sNDiv = m_sNDiv;
//ModelAnalytic2();
//return;

   if( Work_CancelHelper() ) return;
   

/*AP_SET_TOutNumber ap00( new SET_TOutNumber() ); 
AP_SET_TOutNumber ap01( new SET_TOutNumber() ); 
*ap00 + TOutNumber(0.2, 0, 0) + TOutNumber(0.5, 0, 1) + TOutNumber(0.3, 1, 1);
*ap01 + TOutNumber(0.3, 0, 0) + TOutNumber(0.2, 0, 2) + TOutNumber(0.5, 2, 2);
TNodeSFTree nsft00( EN_Plus, "00" );
nsft00 + ap01 + ap00;

AP_SET_TOutNumber apSet00( new SET_TOutNumber() );
nsft00.Grow( *apSet00 );

DumpSet_TOutNumber( *apSet00 );

return;*/

   double dTst1, dTst2, dTst3, dTst4;
   if( m_mtModelType == MT_Analytical )
	 dTst1 = m_lInTest1,
	 dTst2 = m_lInTest2,
	 dTst3 = m_lInTest3,
	 dTst4 =m_lInTest4;
   else
	 dTst1 = m_dInTest1D,
	 dTst2 = m_dInTest2D,
	 dTst3 = m_dInTest3D,
	 dTst4 =m_dInTest4D;

   char cCmp = (m_mtModelType == MT_Analytical ? 0:1);

   short shPercent = 0;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );

   TNodeSFTree nsft2_16( EN_Plus, cCmp, Work_CancelHelper, this, "2_16" );
   nsft2_16 + m_arrRF[ ST2 ] + m_arrRF[ ST4 ] + m_arrRF[ ST6 ] +
	 m_arrRF[ ST8 ] + m_arrRF[ ST10 ] + m_arrRF[ ST12 ] +
	 m_arrRF[ ST14 ] + m_arrRF[ ST16 ];

      
   TNodeSFTree nsft20_18( EN_Mul, cCmp, Work_CancelHelper, this, "20_18" );
   nsft20_18 + m_arrRF[ ST20 ] + m_arrRF[ ST18 ];


/*AP_SET_TOutNumber apSetAA( new SET_TOutNumber() );
nsft20_18.Grow( *apSetAA );
DumpSet_TOutNumber( *apSetAA );*/


   TNodeSFTree nsft22( EN_Plus, cCmp, Work_CancelHelper, this, "22" );
   nsft22 + m_arrRF[ ST22 ];


   TNodeSFTree nsft24( EN_Plus, cCmp, Work_CancelHelper, this, "24" );
   nsft24 + m_arrRF[ ST24 ];

   TNodeSFTree nsft28_26( EN_Plus, cCmp, Work_CancelHelper, this, "28_26" );
   nsft28_26 + m_arrRF[ ST28 ] + m_arrRF[ ST26 ];



/*AP_SET_TOutNumber apSetBB( new SET_TOutNumber() );
TNodeSFTree nsftTT( EN_Mul, "TT" );
nsftTT = nsft20_18 + nsft22 + nsft24 + nsft28_26;
nsftTT.Grow( *apSetBB );

GetStatistic( &m_vton[1], *apSetBB, m_ansStat[ 2 ], m_sfStatOn, NULL );
*/

   TNodeSFTree nsftB_C( EN_Mul, cCmp, Work_CancelHelper, this, "B_C" );

   nsftB_C = nsft2_16 * (nsft20_18 + nsft22 + nsft24 + nsft28_26);
   /*nsftB_C = nsft2_16 * nsftTT;*/

   AP_SET_TOutNumber apSetB_C( new SET_TOutNumber() );
   if( nsftB_C.Grow(*apSetB_C) ) return;

#if defined(_SIZE_STATISTIC_)
   size_t szBC = apSetB_C->size();
#endif

      
   GetStatistic( &m_vton[0], *apSetB_C, m_ansStat[ 0 ], m_sfStatOn, m_bStatIn ? NULL:&dTst1 );

   double dL, dU;
   if( m_bCalibrateAnalytic )
	{
      if( !DivideByP(*apSetB_C, m_dP1, dTst1) )
	   {
		 SetCalibrateError( m_dP1 );  
		 return;
	   }
	}
   else
     DivideByValue( *apSetB_C, dTst1, dL, dU );
   
   m_arropSampleK[ C_B ].Clear();
   m_arropSampleK[ C_B ].m_mx = dL;
   m_arropSampleK[ C_C ].Clear();
   m_arropSampleK[ C_C ].m_mx = dU;

   GetStatistic( NULL, *apSetB_C, m_ansStat[ 1 ], m_sfStatOn, NULL );
   

   shPercent = 10;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;


   nsftB_C.m_lst.clear(); nsftB_C + apSetB_C;
   

   TNodeSFTree nsft34_30( EN_Plus, cCmp, Work_CancelHelper, this, "34_30" );
   nsft34_30 + m_arrRF[ ST34 ] + m_arrRF[ ST32 ] + m_arrRF[ ST30 ];

   TNodeSFTree nsft36( EN_Mul, cCmp, Work_CancelHelper, this, "36" );
   nsft36 + m_arrRF[ ST36 ];

   TNodeSFTree nsft40_38( EN_Mul, cCmp, Work_CancelHelper, this, "40_38" );
   nsft40_38 + m_arrRF[ ST40 ] + m_arrRF[ ST38 ];

   TNodeSFTree nsft44_42( EN_Plus, cCmp, Work_CancelHelper, this, "44_42" );
   nsft44_42 + m_arrRF[ ST44 ] + m_arrRF[ ST42 ];

   TNodeSFTree nsftD_E( EN_Mul, cCmp, Work_CancelHelper, this, "D_E" );
   nsftD_E = nsftB_C * (nsft34_30*nsft36 + nsft40_38 + nsft44_42);
   
   AP_SET_TOutNumber apSetD_E( new SET_TOutNumber() );
   if( nsftD_E.Grow(*apSetD_E) ) return;

#if defined(_SIZE_STATISTIC_)
   size_t szDE = apSetD_E->size();
#endif

   
   GetStatistic( &m_vton[1], *apSetD_E, m_ansStat[ 2 ], m_sfStatOn, m_bStatIn ? NULL:&dTst2 );

   if( m_bCalibrateAnalytic )
	{
      if( !DivideByP(*apSetD_E, m_dP2, dTst2) )
	   {
		 SetCalibrateError( m_dP2 );  
		 return;
	   }
	}
   else
     DivideByValue( *apSetD_E, dTst2, dL, dU );
   
   m_arropSampleK[ C_D ].Clear();
   m_arropSampleK[ C_D ].m_mx = dL;
   m_arropSampleK[ C_E ].Clear();
   m_arropSampleK[ C_E ].m_mx = dU;

   GetStatistic( NULL, *apSetD_E, m_ansStat[ 3 ], m_sfStatOn, NULL );

   shPercent = 30;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;


   nsftD_E.m_lst.clear(); nsftD_E + apSetD_E;
   

   TNodeSFTree nsft48_46( EN_Plus, cCmp, Work_CancelHelper, this, "48_46" );
   nsft48_46 + m_arrRF[ ST48 ] + m_arrRF[ ST46 ];

   TNodeSFTree nsft50( EN_Mul, cCmp, Work_CancelHelper, this, "50" );
   nsft50 + m_arrRF[ ST50 ];

   TNodeSFTree nsft54_52( EN_Mul, cCmp, Work_CancelHelper, this, "54_52" );
   nsft54_52 + m_arrRF[ ST54 ] + m_arrRF[ ST52 ];

   TNodeSFTree nsftF_G( EN_Mul, cCmp, Work_CancelHelper, this, "F_G" );
   nsftF_G = nsftD_E * (nsft48_46*nsft50 + nsft54_52);

   AP_SET_TOutNumber apSetF_G( new SET_TOutNumber() );
   if( nsftF_G.Grow(*apSetF_G) ) return;

#if defined(_SIZE_STATISTIC_)
   size_t szFG = apSetF_G->size();
#endif

   
   GetStatistic( &m_vton[2], *apSetF_G, m_ansStat[ 4 ], m_sfStatOn, m_bStatIn ? NULL:&dTst3 );

   if( m_bCalibrateAnalytic )
	{
      if( !DivideByP(*apSetF_G, m_dP3, dTst3) )
	   {
		 SetCalibrateError( m_dP3 );  
		 return;
	   }
	}
   else
     DivideByValue( *apSetF_G, dTst3, dL, dU );
   
   m_arropSampleK[ C_F ].Clear();
   m_arropSampleK[ C_F ].m_mx = dL;
   m_arropSampleK[ C_G ].Clear();
   m_arropSampleK[ C_G ].m_mx = dU;

   GetStatistic( NULL, *apSetF_G, m_ansStat[ 5 ], m_sfStatOn, NULL );

   shPercent = 60;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;


   nsftF_G.m_lst.clear(); nsftF_G + apSetF_G;
   

   TNodeSFTree nsft58_56( EN_Mul, cCmp, Work_CancelHelper, this, "58_56" );
   nsft58_56 + m_arrRF[ ST58 ] + m_arrRF[ ST56 ];

   TNodeSFTree nsft60( EN_Plus, cCmp, Work_CancelHelper, this, "60" );
   nsft60 + m_arrRF[ ST60 ];

   TNodeSFTree nsft62( EN_Plus, cCmp, Work_CancelHelper, this, "62" );
   nsft62 + m_arrRF[ ST62 ];

   TNodeSFTree nsftH_I( EN_Mul, cCmp, Work_CancelHelper, this, "H_I" );
   nsftH_I = nsftF_G * (nsft58_56 + nsft60 + nsft62);

   AP_SET_TOutNumber apSetH_I( new SET_TOutNumber() );
   if( nsftH_I.Grow(*apSetH_I) ) return;

#if defined(_SIZE_STATISTIC_)
   size_t szHI = apSetH_I->size();
#endif
   
	/*{
      IT_SET_TOutNumber it( apSetH_I->begin() );	  

	  double dDelta, dDeltaMax = -1, dDeltaSumm = 0;
	  for( ; it != apSetH_I->end(); ++it )
	   {	     
		 if( dDeltaMax < it->m_tpP )  dDeltaMax = it->m_tpP;
		 dDeltaSumm += it->m_tpP;
	   }

	  dDeltaSumm /= double(apSetH_I->size() - 1);

      basic_stringstream<char> sssc;
      sssc << "N = " << apSetH_I->size() << "; ave = " << dDeltaSumm << "; max = " << dDeltaMax;
      MessageBox( NULL, sssc.str().c_str(), "JJ", MB_OK );
	}*/

   
   GetStatistic( &m_vton[3], *apSetH_I, m_ansStat[ 6 ], m_sfStatOn, m_bStatIn ? NULL:&dTst4 );

   if( m_bCalibrateAnalytic )
	{
      if( !DivideByP(*apSetH_I, m_dP4, dTst4) )
	   {
		 SetCalibrateError( m_dP4 );  
		 return;
	   }
	}
   else
     DivideByValue( *apSetH_I, dTst4, dL, dU );
   
   m_arropSampleK[ C_H ].Clear();
   m_arropSampleK[ C_H ].m_mx = dL;
   m_arropSampleK[ C_I ].Clear();
   m_arropSampleK[ C_I ].m_mx = dU;

   GetStatistic( NULL, *apSetH_I, m_ansStat[ 7 ], m_sfStatOn, NULL );

   

   if( m_bCalibrateAnalytic )	
	{
	  if( m_mtModelType == MT_Analytical )
		m_lInTest1 = dTst1,
		m_lInTest2 = dTst2,
		m_lInTest3 = dTst3,
		m_lInTest4 = dTst4;
	  else
		m_dInTest1D = dTst1,
		m_dInTest2D = dTst2,
		m_dInTest3D = dTst3,
		m_dInTest4D = dTst4;

	  m_bRequiresSave = true;
	}

   shPercent = 90;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;

/*#if defined(_SIZE_STATISTIC_)
	{
	  basic_stringstream<char> strm;
	  strm << szBC << "; " << szDE << "; " << szFG << "; " << szHI;
	  MessageBox( NULL, strm.str().c_str(), "TT", MB_OK );
	}
#endif*/

 }


void __fastcall CMGertNet::ModelAnalytic2()
 {  
   if( Work_CancelHelper() ) return;
   
   double dTst1, dTst2, dTst3, dTst4;
   /*if( m_mtModelType == MT_Analytical )
	 dTst1 = m_lInTest1,
	 dTst2 = m_lInTest2,
	 dTst3 = m_lInTest3,
	 dTst4 =m_lInTest4;
   else*/
	 dTst1 = m_dInTest1D,
	 dTst2 = m_dInTest2D,
	 dTst3 = m_dInTest3D,
	 dTst4 =m_dInTest4D;
   

   short shPercent = 0;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );

   TNodeATree nsft2_16( m_arrRF[ST2] );
   nsft2_16 + m_arrRF[ ST4 ] + m_arrRF[ ST6 ] +
	 m_arrRF[ ST8 ] + m_arrRF[ ST10 ] + m_arrRF[ ST12 ] +
	 m_arrRF[ ST14 ] + m_arrRF[ ST16 ];

      
   TNodeATree nsft20_18( m_arrRF[ST20] );
   nsft20_18 * m_arrRF[ ST18 ];

   
   TNodeATree nsft28_26( m_arrRF[ST28] );
   nsft28_26 + m_arrRF[ ST26 ];


   TNodeATree nsftB_C( "B_C" );

   nsftB_C = nsft2_16 * (nsft20_18 + m_arrRF[ST24] + m_arrRF[ST22] + nsft28_26);
   
   if( Work_CancelHelper() ) return;

   m_ansStat[ 0 ].dMinVal = 0;
   m_ansStat[ 0 ].dMaxVal = dTst1;
   

   TNodeATree nsftC;
   TNodeATree nsftLeft, nsftRight;
   nsftB_C.VentilCut( dTst1, nsftC, nsftLeft, nsftRight, 1, m_dIntegralAccuracy );

   m_ansStat[ 1 ].dMinVal = dTst1;
   m_ansStat[ 1 ].dMaxVal = nsftB_C.m_dMaxVal;
   m_ansStat[ 1 ].dAvg = nsftC.m_dM;
   m_ansStat[ 1 ].dDx = sqrt( nsftC.m_dD );


   m_arropSampleK[ C_B ].Clear();
   m_arropSampleK[ C_B ].m_mx = nsftLeft.m_dM;
   m_arropSampleK[ C_C ].Clear();
   m_arropSampleK[ C_C ].m_mx = nsftRight.m_dM;

      
   shPercent = 10;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;
   

   TNodeATree nsft34_30( m_arrRF[ST34] );
   nsft34_30 + m_arrRF[ ST32 ] + m_arrRF[ ST30 ];

   TNodeATree nsft36( m_arrRF[ST36] );
   

   TNodeATree nsft40_38( m_arrRF[ST40] );
   nsft40_38 + m_arrRF[ ST38 ];

   TNodeATree nsft44_42( m_arrRF[ST44] );
   nsft44_42 + m_arrRF[ ST42 ];

   TNodeATree nsftD_E( "D_E" );
   nsftD_E = nsftC * (nsft34_30*nsft36 + nsft40_38 + nsft44_42);

   if( Work_CancelHelper() ) return;

   m_ansStat[ 2 ].dMinVal = 0;
   m_ansStat[ 2 ].dMaxVal = dTst2;

   double dNorma = nsftRight.m_dM;
   nsftD_E.VentilCut( dTst2, nsftC, nsftLeft, nsftRight, dNorma, m_dIntegralAccuracy );
      
   m_ansStat[ 3 ].dMinVal = dTst2;
   m_ansStat[ 3 ].dMaxVal = nsftD_E.m_dMaxVal;
   m_ansStat[ 3 ].dAvg = nsftC.m_dM;
   m_ansStat[ 3 ].dDx = sqrt( nsftC.m_dD );
   
   m_arropSampleK[ C_D ].Clear();
   m_arropSampleK[ C_D ].m_mx = nsftLeft.m_dM;
   m_arropSampleK[ C_E ].Clear();
   m_arropSampleK[ C_E ].m_mx = nsftRight.m_dM;
   

   shPercent = 30;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;
      

   TNodeATree nsft48_46( m_arrRF[ST48] );
   nsft48_46 + m_arrRF[ ST46 ];

   TNodeATree nsft50( m_arrRF[ST50] );
   

   TNodeATree nsft54_52( m_arrRF[ST54] );
   nsft54_52 + m_arrRF[ ST52 ];

   TNodeATree nsftF_G( "F_G" );
   nsftF_G = nsftC * (nsft48_46*nsft50 + nsft54_52);


   if( Work_CancelHelper() ) return;

   m_ansStat[ 4 ].dMinVal = 0;
   m_ansStat[ 4 ].dMaxVal = dTst3; 

   dNorma = nsftRight.m_dM;
   nsftF_G.VentilCut( dTst3, nsftC, nsftLeft, nsftRight, dNorma, m_dIntegralAccuracy );
      
   m_ansStat[ 5 ].dMinVal = dTst3;
   m_ansStat[ 5 ].dMaxVal = nsftF_G.m_dMaxVal;
   m_ansStat[ 5 ].dAvg = nsftC.m_dM;
   m_ansStat[ 5 ].dDx = sqrt( nsftC.m_dD );
   
   m_arropSampleK[ C_F ].Clear();
   m_arropSampleK[ C_F ].m_mx = nsftLeft.m_dM;
   m_arropSampleK[ C_G ].Clear();
   m_arropSampleK[ C_G ].m_mx = nsftRight.m_dM;


   
   shPercent = 60;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;
      

   TNodeATree nsft58_56( m_arrRF[ST58] );
   nsft58_56 + m_arrRF[ ST56 ];

   TNodeATree nsft60( m_arrRF[ST60] );   
   TNodeATree nsft62( m_arrRF[ST62] );
   

   TNodeATree nsftH_I( "H_I" );
   nsftH_I = nsftC * (nsft58_56 + nsft60 + nsft62);

   if( Work_CancelHelper() ) return;

   m_ansStat[ 6 ].dMinVal = 0;
   m_ansStat[ 6 ].dMaxVal = dTst4; 

   dNorma = nsftRight.m_dM;
   nsftH_I.VentilCut( dTst4, nsftC, nsftLeft, nsftRight, dNorma, m_dIntegralAccuracy );
   
   m_ansStat[ 7 ].dMinVal = dTst4;
   m_ansStat[ 7 ].dMaxVal = nsftH_I.m_dMaxVal;
   m_ansStat[ 7 ].dAvg = nsftC.m_dM;
   m_ansStat[ 7 ].dDx = sqrt( nsftC.m_dD );
   
   m_arropSampleK[ C_H ].Clear();
   m_arropSampleK[ C_H ].m_mx = nsftLeft.m_dM;
   m_arropSampleK[ C_I ].Clear();
   m_arropSampleK[ C_I ].m_mx = nsftRight.m_dM;
      
   
   shPercent = 90;
   (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
   if( Work_CancelHelper() ) return;
 }


void CMGertNet::MakeWork()
 {   
   while( !m_bFinalize )
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


      m_bCNotified = false;

	  DATE dtEnd;
	  short shPercent;
	  m_bsError = L"";

	  m_dtEnd = m_dtSta = COleDateTime::GetCurrentTime();


	  for( int jc = 0; jc < sizeof(m_vton)/sizeof(m_vton[0]); ++jc )
	    m_vton[ jc ].clear();
	  for( jc = 0; jc < sizeof(m_ston)/sizeof(m_ston[0]); ++jc )
	    m_ston[ jc ].clear();

      
      if( m_mtModelType == MT_Imitate || m_mtModelType == MT_ImitateIndistinct ) 
	    ModelImitate();
	  else if( m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct )
	    ModelAnalytic();
	  else if( m_mtModelType == MT_Analytical2 )
	    ModelAnalytic2();

      m_bCalibrateAnalytic = false;

	  if( CheckCancel() ) 
	   {	     
	     Work_CancelHelper();
	     goto L_REENTER_POINT;
	   }

	  shPercent = 100;
	  (this->*m_pfnNotify)( NM_OnStepOfWork, &shPercent );
	  m_dtEnd = COleDateTime::GetCurrentTime();
	  dtEnd = m_dtEnd - m_dtSta;
	  (this->*m_pfnNotify)( NM_OnEndOfWork, &dtEnd );

L_REENTER_POINT:
	  m_evInternalEndCalc.SetEvent(); 
	  ::SuspendThread( m_tfFunc.m_hthread );
	}
 }

void __fastcall CMGertNet::Notify_PostMessage( NotifyMsg msg, void* pDta )
 {
   switch(msg)
	{
	  case NM_OnEndOfWork: case NM_OnEndOfWork2:
	  case NM_OnEndOfWork3:
	   ::PostMessage( m_hwNotify, msg, 0, 0 );
	    break;

	  case NM_OnStepOfWork: case NM_OnStepOfWork2:
	  case NM_OnStepOfWork3:
	    ::PostMessage( m_hwNotify, msg, (WPARAM)*(short*)pDta, 0 );
	    break;

	  case NM_OnCancel: case NM_OnCancel3:
	    ::PostMessage( m_hwNotify, msg, 0, 0 );
	    break;	  

	  case NM_OnErrorCalc:
	    ::PostMessage( m_hwNotify, msg, 0, 0 );
	    break;	  
	};
 }

void __fastcall CMGertNet::Notify_Callback( NotifyMsg msg, void* pDta )
 {
   switch(msg)
	{
	  case NM_OnEndOfWork: case NM_OnEndOfWork2:
	  case NM_OnEndOfWork3:
	    Fire_OnEndOfWork( *(DATE*)pDta, msg );
	    break;

	  case NM_OnStepOfWork: case NM_OnStepOfWork2:
	  case NM_OnStepOfWork3:
	    Fire_OnStepOfWork( *(short*)pDta, msg );
	    break;

	  case NM_OnCancel: case NM_OnCancel3:
	    Fire_OnCancel( *(DATE*)pDta, msg );
	    break;	  	  

	  case NM_OnErrorCalc:
	    Fire_OnErrorCalc( *(BSTR*)pDta, msg );
	    break;	  	  
	};
 }


STDMETHODIMP CMGertNet::GetInfSampleK(WCHAR cState, long *pMin, long *pMax, double *pMx, double *pDx)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_K == 0 )
	 {
	   Error( L"No runs", IID_IMGertNet, E_FAIL ); return E_FAIL;
	 }
	if( !pMin || !pMax || !pMx || !pDx ) return E_POINTER;

	short shInd = StateToIndex( cState );
	if( shInd == -1 )
	 {
	   PassError( L"IMGertNet.get_Counter: bad index of state", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet  );
       return E_INVALIDARG;
	 }	
	

	if( m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct || m_mtModelType == MT_Analytical2 )
	 {
	   /**pMin = *pMax = 0; *pDx = 0;
	   *pMx = m_arropSampleK[ shInd ].m_mx;*/

	   *pMin = double(m_arropSampleK[ shInd ].m_min);
	   *pMax = double(m_arropSampleK[ shInd ].m_max);
	   *pMx = m_arropSampleK[ shInd ].m_mx;
	   *pDx = m_arropSampleK[ shInd ].m_dx;
	 }
	else
	 {
	   /*if( m_smSampleK.GetDrt() )
		{
		  m_smSampleK.OutParams( m_arropSampleK, m_N );
		  m_smSampleK.SetDrt( false );
		}*/

	   *pMin = double(m_arropSampleK[ shInd ].m_min);// / double(m_N);
	   *pMax = double(m_arropSampleK[ shInd ].m_max);// / double(m_N);
	   *pMx = m_arropSampleK[ shInd ].m_mx;// / double(m_N);
	   *pDx = m_arropSampleK[ shInd ].m_dx;// / double(m_N);
	 }	  

	return S_OK;
}





STDMETHODIMP CMGertNet::get_IsRunning(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	*pVal = CheckRunning() ? VARIANT_TRUE:VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CMGertNet::GetInfSampleKIdx(WCHAR cState, double *pMin, double *pMax, double *pMx, double *pDx)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_K == 0 )
	 {
	   Error( L"No runs", IID_IMGertNet, E_FAIL ); return E_FAIL;
	 }
	if( !pMin || !pMax || !pMx || !pDx ) return E_POINTER;

	short shInd = StateToIndex( cState );
	if( shInd == -1 )
	 {
	   PassError( L"IMGertNet.get_Counter: bad index of state", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet  );
       return E_INVALIDARG;
	 }	


	if( (m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct || m_mtModelType == MT_Analytical2) &&
	    m_sfStatOn == TSF_No
	  )
	 {
	   Error( L"In Analytical mode cann't get GetInfSampleK - use GetInfSampleK", IID_IMGertNet, E_FAIL ); 
	   return E_FAIL;
	 }

	if( m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct || m_mtModelType == MT_Analytical2 )
	 {
	   *pMin = m_ansStat[ shInd ].dMinVal;
	   *pMax = m_ansStat[ shInd ].dMaxVal;
	   *pMx = m_ansStat[ shInd ].dAvg;
	   *pDx = m_ansStat[ shInd ].dDx;
	 }
	else
	 {
	   /*if( m_smSampleIdx.GetDrt() )
		{
		  m_smSampleIdx.OutParamsDiv( m_arropSampleIdx, &m_smSampleK );
		  m_smSampleIdx.SetDrt( false );
		}
	   
	   *pMin = double(m_arropSampleIdx[ shInd ].m_min);
	   *pMax = double(m_arropSampleIdx[ shInd ].m_max);
	   *pMx = m_arropSampleIdx[ shInd ].m_mx;
	   *pDx = m_arropSampleIdx[ shInd ].m_dx;	*/

	   *pMin = m_ssSampleK[ shInd ].dMinVal;
	   *pMax = m_ssSampleK[ shInd ].dMaxVal;
	   *pMx  = m_ssSampleK[ shInd ].dAvg;
	   *pDx  = m_ssSampleK[ shInd ].dM2;
	 }
	  

	return S_OK;
}

STDMETHODIMP CMGertNet::TestFunc(BSTR bsFac, double dX, double* dY)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

    
	if( !dY ) return E_POINTER;

	TCmpSB_BS dta( bsFac );
	TSincBnd* pBnd = find_if( &arrbsFacBnd[0], &arrbsFacBnd[NUMBER_SINKS], dta );
	if( pBnd == &arrbsFacBnd[NUMBER_SINKS] )
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"No such Factor [" << Chk(bsFac) << L"]";
	   PassError( strm.str().c_str(), E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	   return E_FAIL;  
	 }
   
	CRndFunction& rFunc = m_arrRF[ pBnd - &arrbsFacBnd[0] ]; 

    *dY = rFunc( dX );

	return S_OK;
}



STDMETHODIMP CMGertNet::ChangeDirty(VARIANT_BOOL bFlDirty)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	bool bDrt = (bFlDirty == VARIANT_TRUE) ? true:false;
	m_bRequiresSave = bDrt;

    CApplyDirty<CFactor> ad1( bDrt );
	CApplyDirty<CLingvoEnum> ad2( bDrt );

	CApplyDirty<CCollFac> ad11( bDrt );
	CApplyDirty<CCollLingvo> ad22( bDrt );

	ForEach_InColl<CCollFac, ICollFac, 
	  CFactor, CApplyDirty<CCollFac>, 
	  CApplyDirty<CFactor> >( m_spCollFac.p, ad11, ad1 );

    ForEach_InColl<CCollLingvo, ICollLingvo, 
	  CLingvoEnum, CApplyDirty<CCollLingvo>, 
	  CApplyDirty<CLingvoEnum> >( m_spCollEnum.p, ad22, ad2 );


	/*CComObject<CCollFac>* pCF = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );
	CComObject<CCollFac>::TMAP::iterator it1( pCF->m_coll.begin() );
	CComObject<CCollFac>::TMAP::iterator it2( pCF->m_coll.end() );
    for( ; it1 != it2; ++it1 )
	 {
	   CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it1->second );
	   pFac->m_bRequiresSave = bDrt;
	 }*/

	

	return S_OK;
}

STDMETHODIMP CMGertNet::CalcDeltaQ(ICollSF* pCollSF, long lK, long lN, short shMedVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pCollSF ) return E_POINTER;

	if( lK < 1 || lN < 1 ) return E_INVALIDARG;
	m_K = lK, m_N = lN;

	 if( CheckRunning() || CheckRunning2() )
	 {
       Error( L"CMGertNet::CalcDeltaQ: is already running", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	m_evCancel.ResetEvent();

	if( shMedVal < -1 || shMedVal >= CLingvoEnum::GetCount() 	    
	  ) 
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"Need: 0 <= MedVal < " << CLingvoEnum::GetCount() << L" or -1 (calc from current model)";
	   Error( strm.str().c_str(), IID_IMGertNet, E_INVALIDARG );
	   return E_INVALIDARG;
	 }

	m_shMedVal = shMedVal; 
	m_lRndBase = 1;

	if( !m_spCollFac || !m_spCollEnum )
	 {
	   PassError( L"CMGertNet::CalcDeltaQ: Object MGertNet is not initialized", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	   return E_FAIL;
	 }

	long lCnt1 = 0, lCnt2 = 0;
	m_spCollFac->get_Count( &lCnt1 ), m_spCollEnum->get_Count( &lCnt2 );
	if( lCnt1 < 1 || lCnt2 < 1 )
	 {
	   PassError( L"CMGertNet::CalcDeltaQ: Object MGertNet is not initialized", E_FAIL, GetObjectCLSID(), IID_IMGertNet );
	   return E_FAIL;
	 }
    			
	m_pfnNotify = ::IsWindow( m_hwNotify ) == TRUE ? &CMGertNet::Notify_PostMessage:&CMGertNet::Notify_Callback;
	 					

	m_spCollSFKeep = pCollSF;
	if( m_tfFuncCDQ.Start() == FALSE )
	 {
	   m_spCollSFKeep.Release();
	   Error( L"CMGertNet::CalcDeltaQ: cann't start work2", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }
	
	return S_OK;
}

typedef list<short> TL_sht;
typedef TL_sht::iterator IT_TL_sht;

void CMGertNet::MakeWork2()
 {   
   while( !m_bFinalize )
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


      m_bCNotified = false;

	  m_dtEnd2 = m_dtSta2 = COleDateTime::GetCurrentTime();
      m_lRndBase = 1;

	  m_bsError = L"";

	  DATE dtEnd; 
	  short shPercent;
	  TL_sht lstStoreFac;	

	  CComObject<CCollFac>* pObj = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );

	  CComObject<CCollFac>::TMAP::iterator it1( pObj->m_coll.begin() );
	  CComObject<CCollFac>::TMAP::iterator it2( pObj->m_coll.end() );
	  lstStoreFac.assign( pObj->m_coll.size() );
	  IT_TL_sht itKeep( lstStoreFac.begin() );
	  for( ; it1 != it2; ++it1, ++itKeep )
	   {
	     CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it1->second );
	     *itKeep = pFac->m_sValue;
		 if( m_shMedVal != -1 )
		   pFac->m_sValue = m_shMedVal;
	   }
      
      CComObject<CCollSF>* pCollSF = static_cast<CComObject<CCollSF>*>( m_spCollSFKeep.p ); 

	  CComObject<CCollSF>* pCollSFTmp;
	  HRESULT hr = CComObject<CCollSF>::CreateInstance( &pCollSFTmp );
	  CComPtr<ICollSF> spCollSFTmp;
      hr = pCollSFTmp->_InternalQueryInterface( IID_ICollSF, (void**)&spCollSFTmp );
	  /*if( FAILED(hr) )
	   {
		 delete pCollSFTmp;
		 PassError( L"_InternalQueryInterface of IID_ICollFac failed", hr, GetObjectCLSID(), IID_IMGertNet  );
		 return hr;
	   }*/

	  int iTotal = pCollSF->m_coll.size() + 1;
	  m_i64TotalRng = iTotal;
	  
	  int iNotifyStepAbs = double(iTotal) / 100.0 * double(1);
	  if( iNotifyStepAbs < 1 ) iNotifyStepAbs = 1;
	  CComObject<CCollSF>::TMAP::iterator itSF( pCollSF->m_coll.begin() );	  
	  double dQP0 = 0;
	  for( int i = 0; i < iTotal; m_i64CountRng  = ++i )
	   {
	     if( CheckCancel() )
		  {
			m_dtEnd2 = COleDateTime::GetCurrentTime();
			dtEnd = m_dtEnd2 - m_dtSta2;
			(this->*m_pfnNotify)( NM_OnCancel, &dtEnd );
			goto L_REENTER_POINT;
		  }
		 if( (i % iNotifyStepAbs) == 0 )
		  { 
			shPercent =  100 * i / iTotal;
			(this->*m_pfnNotify)( NM_OnStepOfWork2, &shPercent );
		  }

		 if( i > 0 )
		  {
		    if( m_shMedVal != -1 )
			 {
			   it1 = pObj->m_coll.begin();
			   it2 = pObj->m_coll.end();
			   for( ; it1 != it2; ++it1 )
				{
				  CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it1->second );						   
				  pFac->m_sValue = m_shMedVal;
				}
			 }
			else
			 {
			   IT_TL_sht itKeep1( lstStoreFac.begin() );
			   IT_TL_sht itKeep2( lstStoreFac.end() );
			   it1 = pObj->m_coll.begin();
			   for( ; itKeep1 != itKeep2; ++itKeep1, ++it1 )
				{
				  CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it1->second );
				  pFac->m_sValue = *itKeep1;
				}
			 }

			CComObject<CSafetyPrecaution>* pSFi = static_cast<CComObject<CSafetyPrecaution>*>( itSF->second );			
            spCollSFTmp->Clear();
			spCollSFTmp->Add( itSF->second, pSFi->m_lID, NULL );			 			
			ApplySF( spCollSFTmp, NULL );
		  }

		 m_evInternalEndCalc.ResetEvent(); 
		 //m_tfFunc.Start();
		 InternalRun( m_N, m_K, VARIANT_FALSE, m_lRndBase, VARIANT_FALSE );
		 WaitForSingleObject( m_evInternalEndCalc, INFINITE );
		 if( CheckCancel() )
		  {
		    Work_CancelHelper();
		    m_dtEnd2 = COleDateTime::GetCurrentTime();			
			goto L_REENTER_POINT;
		  }
		 WaitThrdSleeped( m_tfFunc.m_hthread );

		 /*if( m_mtModelType != MT_Analytical && m_mtModelType != MT_AnalyticalIndistinct && m_mtModelType != MT_Analytical2 && m_smSampleK.GetDrt() )
		   m_smSampleK.OutParams( m_arropSampleK, m_N ),
	       m_smSampleK.SetDrt( false );*/
	 		
		 
		 if( i > 0 )
		  {		   
		    CComObject<CSafetyPrecaution>* pSFi = static_cast<CComObject<CSafetyPrecaution>*>( itSF->second );			
			if( m_mtModelType != MT_Analytical && m_mtModelType != MT_AnalyticalIndistinct && m_mtModelType != MT_Analytical2 )
              pSFi->m_dDeltaQ = dQP0 - (m_arropSampleK[ C_I ].m_mx /*/ double(m_N)*/);
			else
			  pSFi->m_dDeltaQ = dQP0 - m_arropSampleK[ C_I ].m_mx;
		    ++itSF;	 			
		  }
		 else
		  {
		    if( m_mtModelType != MT_Analytical && m_mtModelType != MT_AnalyticalIndistinct && m_mtModelType != MT_Analytical2 )
		      dQP0 = m_arropSampleK[ C_I ].m_mx;// / double(m_N);
			else
			  dQP0 = m_arropSampleK[ C_I ].m_mx;
		  }		 
	   }
	  

	  shPercent = 100;
	  (this->*m_pfnNotify)( NM_OnStepOfWork2, &shPercent );

	  m_dtEnd2 = COleDateTime::GetCurrentTime();
	  dtEnd = m_dtEnd2 - m_dtSta2;
	  (this->*m_pfnNotify)( NM_OnEndOfWork2, &dtEnd );

L_REENTER_POINT:	  

	  IT_TL_sht itKeep1( lstStoreFac.begin() );
	  IT_TL_sht itKeep2( lstStoreFac.end() );
	  it1 = pObj->m_coll.begin();
	  for( ; itKeep1 != itKeep2; ++itKeep1, ++it1 )
	   {
	     CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it1->second );	     
		 pFac->m_sValue = *itKeep1;
	   }

	  spCollSFTmp.Release();
	  m_spCollSFKeep.Release(); 
	  ::SuspendThread( m_tfFuncCDQ.m_hthread );
	}
 }




class TNtf
 {
public:
   TNtf()
	{	  
	}
   bool operator()( int* pArr, int* parrIdx, int iSzArr, __int64 iCurrIter )
	{
	  basic_stringstream<WCHAR> strm;
	  //strm << setw(5);
	  for( int i = 0; i < iSzArr; ++i )
	    strm << pArr[ parrIdx[i] ] << ((i < iSzArr - 1) ?  L"; ":L" ");

	  strm << endl;

      bs += strm.str().c_str();

      return true;
	}
   basic_string< WCHAR > bs;   
 };

STDMETHODIMP CMGertNet::get_TimeWork2(DATE *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	*pVal = m_dtEnd2 - m_dtSta2;

	/*__int64 ii1 = Fac64MN( 8, 5 );
	__int64 ii2 = Fac64MN( 8, 1 );
	__int64 ii3 = Fac64MN( 8, 2 );
	__int64 ii4 = Fac64MN( 8, 8 );
	__int64 ii5 = Fac64MN( 50, 7 );

	__int64 ii6 = Fac64MN( 7, 4 );
	__int64 ii7 = Fac64MN( 7, 2 );
	__int64 ii8 = Fac64MN( 7, 1 );
	__int64 ii9 = Fac64MN( 7, 6 );

	__int64 ii10 = Fac64MN( 1, 1 );
	__int64 ii11 = Fac64MN( 2, 1 );
	__int64 ii12 = Fac64MN( 2, 2 );

	__int64 ii12_ = Fac64MN( 3, 2 );*/
                           //0  1  2  3  4

	/*static int iarr[ 8 ] = { 0, 1, 2, 3, 4, 5, 6, 7 };
    typedef basic_string<WCHAR> TWSTR;
	
    list<TWSTR> lst;
	__int64 ii = Fac64MN( 8, 2 );
	int iCnt = 0;
	for( int i = 0; i < 3; ++i )
	  for( int j = i + 1; j < 4; ++j )
	    for( int k = j + 1; k < 5; ++k )
		 {
		   iCnt++;
           wchar_t cc[ 512 ];
		   swprintf( cc, L"%5d, %5d, %5d", iarr[i], iarr[j], iarr[k] );
		   lst.push_back( basic_string<wchar_t>(cc) );
		 }

	list<TWSTR>::iterator it1( lst.begin() );
	list<TWSTR>::iterator it2( lst.end() );
	for( ; it1 != it2; ++it1 )
	   OutputDebugStringW( it1->c_str() ),
	   OutputDebugStringW( L"\n" );


    TNtf ntf;
    CArrayHyperIterator<int, int, __int64, TNtf&> iter(
	 iarr, sizeof(iarr)/sizeof(iarr[0]), 2, ntf );

	iter.Iterate();
    OutputDebugStringW( ntf.bs.c_str() );*/

	//kkkkkkkkkkkkkk

	/*TNodeSFTree node( EN_Plus );
	
    AP_SET_TOutNumber ap1( new SET_TOutNumber() );
    *ap1 + TOutNumber(0.417, 0) + TOutNumber(1.0 - 0.417, 1);

	AP_SET_TOutNumber ap2( new SET_TOutNumber() );
    *ap2 + TOutNumber(0.750, 1) + TOutNumber(1.0 - 0.750, 2);

	AP_SET_TOutNumber ap3( new SET_TOutNumber() );
    *ap3 + TOutNumber(0.500, 0) + TOutNumber(1.0 - 0.500, 2);

	node + ap1 + ap2 + ap3;

	SET_TOutNumber setRes;

	node.Grow( setRes );

	DumpSet_TOutNumber( setRes );*/

	return S_OK;
}



STDMETHODIMP CMGertNet::OptimizeSelectSF(OptTask otTask, ICollSF *pCollSF, CURRENCY cyMax, double dDeltaQ, double dP0, short shNRetAlternate )
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pCollSF ) return E_POINTER;

	m_d_P0 = dP0;

	if( CheckRunningOpt() )
	 {
       Error( L"CMGertNet::OptimizeSelectSF: is already running", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	m_sOptResults.clear();

	if( otTask == OT_FixMoney_MinQ && COleCurrency(0, 0) == cyMax  ||
	    otTask == OT_FixQ_MinMoney && dDeltaQ == 0
	  )
	 {
	   Error( L"CMGertNet::OptimizeSelectSF: invalid arg cyMax or dDeltaQ", IID_IMGertNet, E_INVALIDARG );   
       return E_INVALIDARG;
	 }

	m_pfnNotify = ::IsWindow( m_hwNotify ) == TRUE ? &CMGertNet::Notify_PostMessage:&CMGertNet::Notify_Callback;

	CollSFSorting csfKeep;
	pCollSF->get_Sorting( &csfKeep );
	pCollSF->put_Sorting( otTask == OT_FixMoney_MinQ ? CSFS_Price:CSFS_DeltaQ );
	CComPtr<IUnknown> spUnkEnum;
	HRESULT hr = pCollSF->get__NewEnum( &spUnkEnum );
	pCollSF->put_Sorting( csfKeep );
	if( FAILED(hr) )
	 {	   
       Error( L"CMGertNet::OptimizeSelectSF: cann't get__NewEnum", IID_IMGertNet, hr );   
       return hr;
	 }

	CComQIPtr<IEnumVARIANT> spEnumVar( spUnkEnum.p );
	if( !spEnumVar )
	 {
	   Error( L"CMGertNet::OptimizeSelectSF: QI IEnumVARIANT", IID_IMGertNet, E_FAIL );
       return E_FAIL;
	 }

	spEnumVar->Reset();
	CComVariant cvVar;
	m_iNSelectMax = 0;
	COleCurrency ocMoney( 0, 0 );
	double dDelta = 0;
	//bool bFlBreak = false;
	while( spEnumVar->Next(1, &cvVar, NULL) == S_OK )
	 {
       CComQIPtr<ISafetyPrecaution> spSP( cvVar.pdispVal );
	   cvVar.Clear();
	   CComObject<CSafetyPrecaution>* pSP = static_cast<CComObject<CSafetyPrecaution>*>( spSP.p );
	   	   
	   if( otTask == OT_FixMoney_MinQ )
		{
          ocMoney += pSP->m_ocCost ;
          //COleCurrency ocMoneyNew( ocMoney + pSP->m_ocCost );
		  if( ocMoney <= cyMax )
            ++m_iNSelectMax;
		  else 
		    break;
		}
	   else
		{
		  dDelta += pSP->m_dDeltaQ;
		  //double dDeltaNew = dDelta + pSP->m_dDeltaQ;
		  if( dDelta <= dDeltaQ )
            ++m_iNSelectMax;
		  else 
		    break;
		}
	 }

	spEnumVar.Release();

	if( !m_iNSelectMax && otTask == OT_FixMoney_MinQ )
	 {	   
	   Error( L"CMGertNet::OptimizeSelectSF: too little cyMax", IID_IMGertNet, E_FAIL );	 	     
       return E_FAIL;
	 }

	if( dDelta < dDeltaQ )
	 {
	   Error( L"CMGertNet::OptimizeSelectSF: requested DeltaQ is too large", IID_IMGertNet, E_FAIL );
	   return E_FAIL;
	 }

	if( otTask == OT_FixQ_MinMoney )
	 {
	   long lCnt;
	   pCollSF->get_Count( &lCnt );
	   m_iNSelectMax = lCnt;
	 }

	long lM, lMFac;	
	pCollSF->get_Count( &lM );	
	
	if( m_otOptimizationType == OT_Integer || 
	    m_otOptimizationType == OT_IntegerAdv ) m_i64TotalOptCycles = lM * 3;
	else
	 {
	   m_i64TotalOptCycles = 0;
	   for( int i = 1; i <= m_iNSelectMax; ++i )
		 m_i64TotalOptCycles += Fac64MN( lM, i );
	 }
	  

	m_otTask = otTask;
    m_cyMax = cyMax;
    m_dDeltaQ = dDeltaQ;
    m_shNRetAlternate = shNRetAlternate; 

    m_i64NotifyStepAbsOpt = double(m_i64TotalOptCycles) / 100.0 * double(1);	
    if( m_i64NotifyStepAbsOpt < 1 ) m_i64NotifyStepAbsOpt = 1;

	
	m_spCollSF_Opt = pCollSF;
	if( m_tfFuncOpt.Start() == FALSE )
	 {	   
	   m_spCollSF_Opt.Release();
	   Error( L"CMGertNet::OptimizeSelectSF: cann't start workOpt", IID_IMGertNet, E_FAIL );
       return E_FAIL;
	 }

	return S_OK;
 }
 

struct TIntTableSlot
 {
   TIntTableSlot()
	{
	}
   TIntTableSlot( OptTask otTask, ISafetyPrecaution* psf, char cStrategy ):
	  m_cStrategy( cStrategy ),
	  m_pSF( psf ),
	  m_otTask( otTask )
	{
	  CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( psf );

	  switch( cStrategy )
	   {
	     /*case 0:		   
		  {
			COleVariant v1, v2;
			if( otTask == OT_FixMoney_MinQ )
			  v1 = pObj->m_dDeltaQ, v2 = pObj->m_ocCost;
			else
			  v2 = pObj->m_dDeltaQ, v1 = pObj->m_ocCost;

			if( ((V_VT((LPVARIANT)v2)&VT_TYPEMASK) == VT_R8 &&
				V_R8((LPVARIANT)v2) == 0.0) ||
				((V_VT((LPVARIANT)v2)&VT_TYPEMASK) == VT_CY &&
				COleCurrency(v2) == COleCurrency(0, 0))
			  )
			  m_Alfa = COleVariant( DBL_MAX );
			else		   
			  VarDiv( (LPVARIANT)v1, (LPVARIANT)v2, (LPVARIANT)m_Alfa );
			break;
		  }

	     case 1:
		    if( otTask == OT_FixMoney_MinQ ) m_Alfa = pObj->m_dDeltaQ;
			else m_Alfa = pObj->m_ocCost;
			break;

         case 2:
		    if( otTask == OT_FixMoney_MinQ ) m_Alfa = pObj->m_ocCost;
			else m_Alfa = pObj->m_dDeltaQ;
			break;*/

	     case 0:		   
		  {
			COleVariant v1, v2;			
			v1 = pObj->m_dDeltaQ, v2 = pObj->m_ocCost;						
			VarDiv( (LPVARIANT)v1, (LPVARIANT)v2, (LPVARIANT)m_Alfa );
			break;
		  }

	     case 1:
		    m_Alfa = pObj->m_dDeltaQ;			
			break;

         case 2:
		    m_Alfa = pObj->m_ocCost;			
			break;
	   };	  	  
	}

   TIntTableSlot( const TIntTableSlot& rS )
	{
	  this->operator=( rS );
	}

   TIntTableSlot& operator=( const TIntTableSlot& rS )
	{
	  m_Alfa = rS.m_Alfa;	  
	  m_otTask = rS.m_otTask;
	  m_cStrategy = rS.m_cStrategy;
	  m_pSF = rS.m_pSF;
	  return *this;
	}

   bool operator<( const TIntTableSlot& rS ) const
	{	  
	  HRESULT hr;
	  if( m_cStrategy == 2 )
	    hr = VarCmp( (LPVARIANT)(LPCVARIANT)m_Alfa, (LPVARIANT)(LPCVARIANT)rS.m_Alfa, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );	  	  	   
	  else
	    hr = VarCmp( (LPVARIANT)(LPCVARIANT)rS.m_Alfa, (LPVARIANT)(LPCVARIANT)m_Alfa, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );	  	  	   

	  return hr == 0;
	}

   COleVariant Target()
	{
	  CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( m_pSF );
	  if( m_otTask == OT_FixMoney_MinQ ) return pObj->m_dDeltaQ;
	  else return pObj->m_ocCost;
	}

   COleVariant Restriction()
	{
	  CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( m_pSF );
	  if( m_otTask == OT_FixMoney_MinQ ) return pObj->m_ocCost;
	  else return pObj->m_dDeltaQ;
	}

   COleVariant m_Alfa;
   ISafetyPrecaution* m_pSF;

   char m_cStrategy;
   OptTask m_otTask;
 };

struct TITS_less : binary_function<TIntTableSlot, TIntTableSlot, bool> {
	bool operator()(const TIntTableSlot& _X, const TIntTableSlot& _Y) const
	 {
	   return _X.operator<( _Y );
	 }
 };

typedef multiset<TIntTableSlot, TITS_less> TS_IntTable;
typedef TS_IntTable::iterator IT_TS_IntTable;

class TOptInteger
 {
public:
   TOptInteger( CMGertNet* pmgn ): 
	  m_pmgn( pmgn ), m_pOptVec( NULL ), 
	  m_pfnInv( pmgn->m_otTask == OT_FixMoney_MinQ ? Inversion_FixMoneyMinQ:Inversion_FixQMinMoney )
	{	  
	  if( pmgn->m_otTask == OT_FixMoney_MinQ )
	    m_cvRestriction = pmgn->m_cyMax;
	  else
	    m_cvRestriction = pmgn->m_dDeltaQ;	  

	  CComObject<CCollSF>* pC = static_cast<CComObject<CCollSF>*>( pmgn->m_spCollSF_Opt.p );      
	  m_pOptVec = new TIdxItem[ (m_sz=pC->m_coll.size()) ];	  
	  CComObject<CCollSF>::TMAP::iterator it( pC->m_coll.begin() );

	  for( TIdxItem* pI = m_pOptVec; it != pC->m_coll.end(); ++it, ++pI )
	    pI->m_p = it->second, pI->m_Select = false, pI->m_bLock = false;
	}

   ~TOptInteger()
	{
	  if( m_pOptVec ) { delete m_pOptVec; m_pOptVec = NULL; }
	}

   bool Optimize();
   void AddResult( TS_IntTable& );
   void DumpMS( TS_IntTable& );
   bool ImprovementOfResult();
   bool ImprovementOfResult_Internal( IT_TS_SetRes& rIt );
   
	   
   CMGertNet* m_pmgn;   
   COleVariant m_cvRestriction;

   struct TIdxItem
	{
	  ISafetyPrecaution* m_p;
	  bool m_Select;
	  bool m_bLock;
	};
   void DumpVec( short );
   TIdxItem* Search( ISafetyPrecaution* );
   TIdxItem* SOneZer( TIdxItem*, bool );
   void Inversion_FixMoneyMinQ( TOptInteger::TIdxItem *p, COleVariant& rovR, COleVariant& rovT );
   void Inversion_FixQMinMoney( TOptInteger::TIdxItem *p, COleVariant& rovR, COleVariant& rovT );
   void UnlockAll()
	{
	  if( m_pOptVec )
	   {
	     for( long sz = m_sz - 1; sz >= 0; --sz )
		   m_pOptVec[ sz ].m_bLock = false;
	   }
	}

   void (TOptInteger::* m_pfnInv)( TIdxItem *p, COleVariant& rovR, COleVariant& rovT );

   TIdxItem* m_pOptVec; 
   long m_sz;
 };

TOptInteger::TIdxItem* TOptInteger::Search( ISafetyPrecaution* p )
 {
   if( m_pOptVec == NULL ) return NULL;
   
   for( long i = 0; i < m_sz; ++i )
	 if( (void*)m_pOptVec[ i ].m_p == (void*)p ) return m_pOptVec + i;

   return NULL;
 }

void TOptInteger::DumpVec( short sh )
 {
   if( m_pOptVec == NULL ) return;

   USES_CONVERSION;
   OutputDebugString( _T("------------\n") );

   basic_stringstream<wchar_t> strm;
   for( long i = 0; i < m_sz; ++i )	
	if( m_pOptVec[i].m_Select )
	 {
	   CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( m_pOptVec[i].m_p );
	   strm << (BSTR)pObj->m_bsName << endl;
	 }

   for( i = 0; i < m_sz; ++i )	
	  strm << (m_pOptVec[i].m_Select ? 1:0);	

   strm << endl;
   OutputDebugString( OLE2CT(strm.str().c_str()) );
 }

void TOptInteger::DumpMS( TS_IntTable& ms )
 {
   USES_CONVERSION;
   OutputDebugString( _T("------------\n") );

   basic_stringstream<wchar_t> strm;
   IT_TS_IntTable it( ms.begin() );
   for( ; it != ms.end(); ++it )
	{
	  CComBSTR bsO;
	  VarFormat( (LPVARIANT)(it->m_Alfa), L"", 0, 0, 0, &bsO );      	  
	  strm << setw(15) << ((BSTR)bsO == NULL ? L"":(BSTR)bsO) << endl;
	}

   OutputDebugString( OLE2CT(strm.str().c_str()) );
 }

bool TOptInteger::Optimize()
 {

   if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) return false;

   TS_IntTable ms1;

   CComObject<CCollSF>* pC = static_cast<CComObject<CCollSF>*>( m_pmgn->m_spCollSF_Opt.p );
   CComObject<CCollSF>::TMAP::iterator it( pC->m_coll.begin() );
   COleCurrency oc0( 0, 0 );
   for( ; it != pC->m_coll.end(); ++it ) 
	{
	  CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( it->second );

	  if( pObj->m_dDeltaQ == 0 || pObj->m_ocCost == oc0 )
	   {	     
		 CComBSTR bs( L"Обнаружены мероприятия с нулевой эффективностью или стоимостью" );   
		 m_pmgn->m_bsError = bs;
         (m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnErrorCalc, &bs );

	     return false;
	   }
	  ms1.insert( TIntTableSlot(m_pmgn->m_otTask, it->second, 0) );
	}

   //DumpMS( ms1 );
   AddResult( ms1 );

   short shPercent = 10;
   (m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnStepOfWork3, &shPercent );

   ms1.clear();
   it = pC->m_coll.begin();

   if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) return false;

   for( ; it != pC->m_coll.end(); ++it ) 
	 ms1.insert( TIntTableSlot(m_pmgn->m_otTask, it->second, 1) );
   
   //DumpMS( ms1 );
   AddResult( ms1 );

   shPercent = 20;
   (m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnStepOfWork3, &shPercent );

   ms1.clear();
   it = pC->m_coll.begin();

   if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) return false;

   for( ; it != pC->m_coll.end(); ++it ) 
	 ms1.insert( TIntTableSlot(m_pmgn->m_otTask, it->second, 2) );

   //DumpMS( ms1 );
   AddResult( ms1 );

   shPercent = 30;
   (m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnStepOfWork3, &shPercent );

   ms1.clear();

   bool bResBool; 

   if( m_pmgn->m_otOptimizationType == OT_IntegerAdv )	
	{
	  bResBool = ImprovementOfResult();
	  if( !m_pmgn->m_bCNotified )
	   {
		 shPercent = 100;
		(m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnStepOfWork3, &shPercent );
	   }

	  //DumpSetRes( m_pmgn->m_sOptResults );
	}
   else bResBool = true;

   return bResBool;
 }

TOptInteger::TIdxItem* TOptInteger::SOneZer( TIdxItem* p, bool bVal )
 {
   TIdxItem* pEnd = m_pOptVec + m_sz;
   if( (void*)p >= (void*)pEnd ) return NULL;

   for( ; p < pEnd; ++p )
	 if( p->m_Select == bVal && p->m_bLock == false ) return p;

   return NULL;
 }

void TOptInteger::Inversion_FixMoneyMinQ( TOptInteger::TIdxItem *p, COleVariant& rovR, COleVariant& rovT )
 {
   CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( p->m_p );

   p->m_bLock = !p->m_bLock;

   if( p->m_Select == true )
	{
	  p->m_Select = false;
	  COleVariant vTmp;
	  VarSub( rovR, COleVariant(pObj->m_ocCost), vTmp );
	  rovR = vTmp;

	  vTmp.Clear();
	  VarSub( rovT, COleVariant(pObj->m_dDeltaQ), vTmp );
	  rovT = vTmp;
	}
   else
	{
	  p->m_Select = true;
	  COleVariant vTmp;
	  VarAdd( rovR, COleVariant(pObj->m_ocCost), vTmp );
	  rovR = vTmp;

	  vTmp.Clear();
	  VarAdd( rovT, COleVariant(pObj->m_dDeltaQ), vTmp );
	  rovT = vTmp;
	}   
 }

void TOptInteger::Inversion_FixQMinMoney( TOptInteger::TIdxItem *p, COleVariant& rovR, COleVariant& rovT )
 {
   CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( p->m_p );
   if( p->m_Select == true )
	{
	  p->m_Select = false;
	  COleVariant vTmp;
	  VarSub( rovR, COleVariant(pObj->m_dDeltaQ), vTmp );
	  rovR = vTmp;

	  vTmp.Clear();
	  VarSub( rovT, COleVariant(pObj->m_ocCost), vTmp );
	  rovT = vTmp;
	}
   else
	{
	  p->m_Select = true;
	  COleVariant vTmp;
	  VarAdd( rovR, COleVariant(pObj->m_dDeltaQ), vTmp );
	  rovR = vTmp;

	  vTmp.Clear();
	  VarAdd( rovT, COleVariant(pObj->m_ocCost), vTmp );
	  rovT = vTmp;
	}   
 }

bool TOptInteger::ImprovementOfResult()
 {
   if( m_pmgn->m_sOptResults.size() == 0 ) return true;

   for( long l = m_sz - 1; l >= 0; --l )
     m_pOptVec[ l ].m_Select = false, m_pOptVec[ l ].m_bLock = false;

   TSampleData* psdBest = m_pmgn->m_sOptResults.rbegin()->get();
   IT_TV_SF it0( psdBest->m_vec.begin() );
   for( ; it0 != psdBest->m_vec.end(); ++it0 )
	{
	  TOptInteger::TIdxItem* ti = Search( it0->p );
	  ti->m_Select = true;
	}

   //DumpVec(0);       
   
   COleVariant cvTSumm;
   if( m_pmgn->m_otTask == OT_FixQ_MinMoney )
	 cvTSumm = psdBest->m_ocSumm;	 
   else
	 cvTSumm = psdBest->m_dSummDQ;	    

   do {
	  IT_TS_SetRes it( m_pmgn->m_sOptResults.end() );
	  UnlockAll();
	  bool bRes = ImprovementOfResult_Internal( it );
	  if( !bRes ) return false;

	  
	  if( it == m_pmgn->m_sOptResults.end() ) break;

	  IT_TS_SetRes it2( m_pmgn->m_sOptResults.end() ); --it2;
	  if( it2 != it )
	   {
	     m_pmgn->m_sOptResults.erase( it );
	     break;
	   }
	  else
	   {
	     COleVariant cvTSumm2;
		 psdBest = m_pmgn->m_sOptResults.rbegin()->get();
		 if( m_pmgn->m_otTask == OT_FixQ_MinMoney )
		   cvTSumm2 = psdBest->m_ocSumm;	 
		 else
		   cvTSumm2 = psdBest->m_dSummDQ;	    

		 HRESULT hr = 
		   VarCmp( cvTSumm, cvTSumm2, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

		 if( (m_pmgn->m_otTask == OT_FixMoney_MinQ && hr == 0) ||
		     (m_pmgn->m_otTask == OT_FixQ_MinMoney && hr == 2)
		   ) 
		  {
		    cvTSumm = cvTSumm2;
			continue;
		  }
		 else
		  {
		    m_pmgn->m_sOptResults.erase( it );
	        break;
		  }
	   }

	} while( 1 );

   return true;
 }

bool TOptInteger::ImprovementOfResult_Internal( IT_TS_SetRes& rIt )
 {      
   TSampleData* psdBest = m_pmgn->m_sOptResults.rbegin()->get();
   COleVariant cvRSumm, cvTSumm;   
   if( m_pmgn->m_otTask == OT_FixQ_MinMoney )
	 cvRSumm = psdBest->m_dSummDQ, cvTSumm = psdBest->m_ocSumm;	 
   else
	 cvTSumm = psdBest->m_dSummDQ, cvRSumm = psdBest->m_ocSumm;	    


   TOptInteger::TIdxItem *tiZero = SOneZer(m_pOptVec, true);
   
   COleVariant cvPrevious1( cvTSumm );
   bool bAccept = false;   
   while( tiZero != NULL )
	{	  
	  (this->*m_pfnInv)( tiZero, cvRSumm, cvTSumm );

	  TOptInteger::TIdxItem *tiOne = SOneZer(m_pOptVec, false);	  	  
	  bool bAccept0 = false;   
	  while( tiOne != NULL )
	   {	     
	     (this->*m_pfnInv)( tiOne, cvRSumm, cvTSumm );

		 //DumpVec(1);

		 HRESULT hr = 
		   VarCmp( cvRSumm, m_cvRestriction, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

		 if( m_pmgn->m_otTask == OT_FixMoney_MinQ )
		   bAccept = (hr == 0 || hr == 1) ? true:false;		    
		 else
		   bAccept = (hr == 2 || hr == 1) ? true:false;

		 if( bAccept )
		  {
		    hr = VarCmp( cvTSumm, cvPrevious1, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

			if( m_pmgn->m_otTask == OT_FixMoney_MinQ )
		      bAccept = (hr == 2) ? true:false;
			else
			  bAccept = (hr == 0) ? true:false;
		  }

		 if( !bAccept ) (this->*m_pfnInv)( tiOne, cvRSumm, cvTSumm );
		 else 
		  { 
		    cvPrevious1 = cvTSumm, bAccept0 = true; 
		  }
		 tiOne = SOneZer( tiOne + 1, false );

		 if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) return false;
	   }
      

      if( !bAccept0 ) (this->*m_pfnInv)( tiZero, cvRSumm, cvTSumm );
	  tiZero = SOneZer( tiZero + 1, true );
	}
   
   long szCountRes = 0;
   for( long i = 0; i < m_sz; ++i )
	 if( m_pOptVec[i].m_Select == true ) szCountRes++;

   if( szCountRes > 0 )
	{
	  AP_TSampleData ap( new TSampleData() );
	  ap->m_vec.assign( szCountRes );
	  ap->m_mode = ((m_pmgn->m_otTask == OT_FixMoney_MinQ) ? 
	   TSampleData::TC_DQ:TSampleData::TC_Money);
	   	  	  

	  for( szCountRes = i = 0; i < m_sz; ++i )
	    if( m_pOptVec[i].m_Select == true ) 
		  ap->m_vec[ szCountRes++ ] = m_pOptVec[i].m_p;

	  ap->Summ();

     /* if( !(m_pmgn->m_otTask == OT_FixQ_MinMoney && 
	      ap->m_dSummDQ < m_pmgn->m_dDeltaQ) 
		)*/

	  /*TS_SetRes sstmp;
	  sstmp.insert( ap );	  
	  DumpSetRes( sstmp );*/

      rIt = m_pmgn->m_sOptResults.insert( ap );	  	  
	}

   return true;
 }

void TOptInteger::AddResult( TS_IntTable& rMS )
 {
   if( rMS.size() == 0 ) return;

   
   IT_TS_IntTable it( rMS.begin() );
   COleVariant cvRSumm( it->Restriction() );
   set<long> setNonCompatibleSF, setInt, setAdded;

   //if( m_pmgn->m_otTask == OT_FixMoney_MinQ )
       
   HRESULT hr = VarCmp( (LPVARIANT)cvRSumm, (LPVARIANT)m_cvRestriction, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );
   if( hr == 0 )
	{	  
	  CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( it->m_pSF );
	  //setSelectedSF.insert( pObj->m_lID );
	  setNonCompatibleSF.insert( pObj->m_vecCollNC.begin(), pObj->m_vecCollNC.end() );
	  setAdded.insert( pObj->m_lID );

	  for( ++it; it != rMS.end(); )
	   {
		 pObj = static_cast<CComObject<CSafetyPrecaution>*>( it->m_pSF );
         set<long>::iterator itLng = setNonCompatibleSF.find( pObj->m_lID );			 
		 if( itLng != setNonCompatibleSF.end() ) rMS.erase( it++ ); 			 
		 else
		  {
            setInt.clear();
			set_intersection( setAdded.begin(), setAdded.end(), 
	          pObj->m_vecCollNC.begin(), pObj->m_vecCollNC.end(), inserter(setInt, setInt.begin()) );

			if( setInt.size() > 0 ) rMS.erase( it++ );
			else
			 {
			   COleVariant vtTmp;
			   VarAdd( (LPVARIANT)cvRSumm, it->Restriction(), (LPVARIANT)vtTmp );
			   hr = VarCmp( (LPVARIANT)vtTmp, (LPVARIANT)m_cvRestriction, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 ); 

			   if( hr == 0 ) { setAdded.insert(pObj->m_lID); ++it; }
			   else if( hr == 1 ) { ++it; break; }
			   else
				{				  
		          if( m_pmgn->m_otTask == OT_FixQ_MinMoney )
				    { ++it; break; }

				  rMS.erase( it++ );
				  continue;
				}

			   cvRSumm = vtTmp;
			   setNonCompatibleSF.insert( pObj->m_vecCollNC.begin(), pObj->m_vecCollNC.end() );
			 }
		  }
	   }
	}
   else if( m_pmgn->m_otTask == OT_FixQ_MinMoney ) it++;	


   ptrdiff_t sz = distance( rMS.begin(), it );

   if( sz > 0 )
	{
	  AP_TSampleData ap( new TSampleData() );
	  ap->m_vec.assign( sz );
	  ap->m_mode = ((m_pmgn->m_otTask == OT_FixMoney_MinQ) ? 
	   TSampleData::TC_DQ:TSampleData::TC_Money);
	   
	  //for( sz--, it--; it != rMS.begin(); --sz, --it )
		//ap->m_vec[ sz ] = it->m_pSF;

	  sz--, it--;
	  do {
	     ap->m_vec[ sz-- ] = it->m_pSF;
	   } while( it-- != rMS.begin() );

	  ap->Summ();

      if( !(m_pmgn->m_otTask == OT_FixQ_MinMoney && 
	      ap->m_dSummDQ < m_pmgn->m_dDeltaQ) 
		)
        m_pmgn->m_sOptResults.insert( ap );	  
	}
 }

class TOptStep
 {
public:
  TOptStep( CMGertNet* pmgn ): m_pmgn( pmgn )
   {
     m_dP0 = pmgn->m_d_P0;
   }

  bool operator()( CComObject<CSafetyPrecaution>** parr, int* parrIdx, int iSz, __int64 i64CurrIter );  
  

CMGertNet* m_pmgn;
double m_dP0;
 };

bool __fastcall CMGertNet::MakeOnePass( TSampleData& rDta, const double dP0, VARIANT_BOOL* pbOver )
 {
   HRESULT hr = ApplyVectorSF( rDta.m_vec, pbOver );
   if( FAILED(hr) )
	{
	  SetGenericError( hr );	  
      return false;
	}

   m_evInternalEndCalc.ResetEvent(); 

   InternalRun( m_N, m_K, VARIANT_FALSE, m_lRndBase, VARIANT_FALSE );
   WaitForSingleObject( m_evInternalEndCalc, INFINITE );
   if( CheckCancel() ) return false;
	
   WaitThrdSleeped( m_tfFunc.m_hthread );

   /*if( m_mtModelType != MT_Analytical && m_mtModelType != MT_AnalyticalIndistinct && m_mtModelType != MT_Analytical2 && m_smSampleK.GetDrt() )
	 m_smSampleK.OutParams( m_arropSampleK, m_N ),
	 m_smSampleK.SetDrt( false );*/

      
   if( m_mtModelType != MT_Analytical && m_mtModelType != MT_AnalyticalIndistinct && m_mtModelType != MT_Analytical2 )
     rDta.m_dSummDQ = dP0 - (m_arropSampleK[ C_I ].m_mx /*/ double(m_N)*/);
   else
	 rDta.m_dSummDQ = dP0 - m_arropSampleK[ C_I ].m_mx;
   

   return true;
 }

bool __fastcall CMGertNet::Work_CancelHelper3()
 {
   DATE dtEnd;  
   if( CheckCancel() )
	{
	  if( !m_bC3Notified )
	   {
		 m_dtEnd3 = COleDateTime::GetCurrentTime();
		 dtEnd = m_dtEnd3 - m_dtSta3;
		 m_bC3Notified = true;
		 (this->*m_pfnNotify)( NM_OnCancel3, &dtEnd );	  	  
	   }
	  return true;
	}
   return false;
 }

bool TOptStep::operator()( CComObject<CSafetyPrecaution>** parr, int* parrIdx, int iSz, __int64 i64CurrIter )
 {
   DATE dtEnd; 
   short shPercent;

   m_pmgn->Revert();

   if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) 
	{
	  return false;
	}

   /*if( m_pmgn->CheckCancel() )
	{
	  m_pmgn->m_dtEnd3 = COleDateTime::GetCurrentTime();
	  dtEnd = m_pmgn->m_dtEnd3 - m_pmgn->m_dtSta3;
	  (m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnCancel3, &dtEnd );
	  return false;
	}*/
   m_pmgn->m_i64CurrIterOpt = i64CurrIter;
   if( (i64CurrIter % m_pmgn->m_i64NotifyStepAbsOpt) == 0 )
	{ 
	  shPercent =  __int64(100) * i64CurrIter / m_pmgn->m_i64TotalOptCycles;
	  (m_pmgn->*(m_pmgn->m_pfnNotify))( NM_OnStepOfWork3, &shPercent );
	}

   set<long> setSelectedSF, setNonCompatibleSF, setInt;
   for( int k = 0; k < iSz; ++k )
	{
	  CComObject<CSafetyPrecaution>& rSF = *parr[ parrIdx[k] ];
	  setSelectedSF.insert( rSF.m_lID );
	  setNonCompatibleSF.insert( rSF.m_vecCollNC.begin(), rSF.m_vecCollNC.end() );
	}

   set_intersection( setSelectedSF.begin(), setSelectedSF.end(), 
	  setNonCompatibleSF.begin(), setNonCompatibleSF.end(), inserter(setInt, setInt.begin()) );

   if( setInt.size() != 0 )  return true;
	

   //m_spCollNC

   AP_TSampleData ap( new TSampleData() );
   ap->m_vec.assign( iSz );
   ap->m_mode = ((m_pmgn->m_otTask == OT_FixMoney_MinQ) ? 
	TSampleData::TC_DQ:TSampleData::TC_Money);
	
   for( parrIdx += --iSz; iSz >= 0; --iSz, --parrIdx )
	 ap->m_vec[ iSz ] = static_cast<ISafetyPrecaution*>( parr[ *parrIdx ] );

   if( m_pmgn->m_otOptimizationType == OT_Full )
	{	  
	  VARIANT_BOOL bOver;
	  if( !(m_pmgn->*CMGertNet::MakeOnePass)(*ap, m_dP0, &bOver) ) return false;
	  if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) return false;
	  ap->SummMoney();
	  ap->m_shOverWrap = (bOver == VARIANT_TRUE ? 1:0);
	}
   else
	{
	  //if( (m_pmgn->*CMGertNet::Work_CancelHelper3)() ) return false;
      ap->Summ();
	}

   if( m_pmgn->m_otTask == OT_FixMoney_MinQ && ap->m_ocSumm > m_pmgn->m_cyMax ||
	   m_pmgn->m_otTask == OT_FixQ_MinMoney && ap->m_dSummDQ < m_pmgn->m_dDeltaQ
	 ) 
	 return true;



   if( m_pmgn->m_sOptResults.size() < m_pmgn->m_shNRetAlternate )
	 m_pmgn->m_sOptResults.insert( ap );
   else
    {
	  if( m_pmgn->m_sOptResults.size() < 1 )
	     m_pmgn->m_sOptResults.insert( ap );
	  else
	   {
		 //TS_SetRes::iterator it1( m_pmgn->m_sOptResults.begin() );
		 //TS_SetRes::iterator it2( m_pmgn->m_sOptResults.end() );

		 TS_SetRes::iterator it1 = m_pmgn->m_sOptResults.end();

		 if( m_pmgn->m_otTask == OT_FixMoney_MinQ )
		  {
			if( (*m_pmgn->m_sOptResults.begin())->m_dSummDQ < ap->m_dSummDQ )
			  it1 = m_pmgn->m_sOptResults.begin();
		   //for( it1 = m_pmgn->m_sOptResults.begin(); it1 != m_pmgn->m_sOptResults.end(); ++it1 )
			 //if( (*it1)->m_dSummDQ < ap->m_dSummDQ ) break;
		  }
		 else if( m_pmgn->m_otTask == OT_FixQ_MinMoney )
		  {
		   //for( ; it1 != it2; ++it1 )
			 //if( (*it1)->m_ocSumm > ap->m_ocSumm ) break;
		    if( (*m_pmgn->m_sOptResults.begin())->m_ocSumm > ap->m_ocSumm )
			   it1 = m_pmgn->m_sOptResults.begin();
		  }

		 if( it1 != m_pmgn->m_sOptResults.end() )
		  {
            /*static int ss = 1;
			if( ss == 1 ) 
			 {
               DumpSetRes( m_pmgn->m_sOptResults );
			   ss = 0;
			 }*/

			m_pmgn->m_sOptResults.erase( it1 );
			m_pmgn->m_sOptResults.insert( ap );

			//DumpSetRes( m_pmgn->m_sOptResults );
		  }
	   }
	}


   return true;
 }

void DumpSetRes( TS_SetRes& rS )
 {
   USES_CONVERSION; 

   OutputDebugString( _T("------------\n") );

   basic_stringstream<wchar_t> strm;
   IT_TS_SetRes it( rS.begin() );
   for( ; it != rS.end(); ++it )
	{
	  strm << setw(15) << (*it)->m_dSummDQ << L"; " << setw(1) << 
	     (BSTR)CComBSTR((LPCTSTR)((*it)->m_ocSumm).Format()) << endl;

	  OutputDebugString( OLE2CT(strm.str().c_str()) );
	  strm.str( L"" );
      
	  IT_TV_SF it2( (*it)->m_vec.begin() );
	  for( ; it2 != (*it)->m_vec.end(); ++it2 )
	   {
	     //CComObject<CSafetyPrecaution>* pObj = static_cast<CComObject<CSafetyPrecaution>*>( it2->p );
	     CComBSTR bsTmp;
	     it2->p->get_Name( &bsTmp );
		 strm << L"\t" << (BSTR)bsTmp << endl;
	   }

	  OutputDebugString( OLE2CT(strm.str().c_str()) );
	  strm.str( L"" );
	}
   
 }//hhhhhhhhhh

typedef list<AP_TSampleData> TL_SetRes;
typedef TL_SetRes::iterator IT_TL_SetRes;

void __fastcall CMGertNet::IncreaseAccuracyOfResult()
 {
   m_i64TotalOptCycles = m_sOptResults.size();   
   m_i64NotifyStepAbsOpt = double(m_i64TotalOptCycles) / 100.0 * double(1);	
   if( m_i64NotifyStepAbsOpt < 1 ) m_i64NotifyStepAbsOpt = 1;
   
   short shPercent;

   TL_SetRes lstTmp;
   copy( m_sOptResults.begin(), m_sOptResults.end(), back_inserter(lstTmp) );
   m_sOptResults.clear();

   IT_TL_SetRes it( lstTmp.begin() );
   for( m_i64CurrIterOpt = 0; it != lstTmp.end(); ++it, ++m_i64CurrIterOpt )
	{
      if( (m_i64CurrIterOpt % m_i64NotifyStepAbsOpt) == 0 )
	   { 
		 shPercent =  __int64(100) * m_i64CurrIterOpt / m_i64TotalOptCycles;
		 (this->*m_pfnNotify)( NM_OnStepOfWork3, &shPercent );
	   } 

	  Revert();

	  VARIANT_BOOL bOver;
	  if( !MakeOnePass(**it, m_d_P0, &bOver) ) return;
	  if( Work_CancelHelper3() ) return;
	  //(*it)->SummMoney();
	  (*it)->m_shOverWrap = (bOver == VARIANT_TRUE ? 1:0);
	}

   for( it = lstTmp.begin() ; it != lstTmp.end(); ++it )
	 m_sOptResults.insert( *it );
 }


void CMGertNet::MakeWorkOpt()
 {   
   while( !m_bFinalize )
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


      //m_otOptimizationType = OT_Integer;

	  m_bCNotified = false;

	  m_dtEnd3 = m_dtSta3 = COleDateTime::GetCurrentTime();      

	  m_bsError = L"";
	  m_i64CurrIterOpt = 0;

	  DATE dtEnd; 
	  short shPercent;

	  
	  CComObject<CCollSF>* pCollSF = static_cast<CComObject<CCollSF>*>( m_spCollSF_Opt.p );
	  int iSz = pCollSF->m_coll.size();
	  CComObject<CSafetyPrecaution>** arrPSP = new CComObject<CSafetyPrecaution>*[ iSz ];

	  CComObject<CCollSF>::TMAP::iterator it1( pCollSF->m_coll.begin() );
	  CComObject<CCollSF>::TMAP::iterator it2( pCollSF->m_coll.end() );
	  int i = 0;
	  for( ; it1 != it2; ++it1, ++i )
	    arrPSP[ i ] = static_cast<CComObject<CSafetyPrecaution>*>( it1->second );
	  
      Snapshot( VARIANT_TRUE );

	  shPercent = 0;
	  (this->*m_pfnNotify)( NM_OnStepOfWork3, &shPercent );

	  if( m_otOptimizationType == OT_Integer || m_otOptimizationType == OT_IntegerAdv )
	   {
	     TOptInteger oi( this );
	     oi.Optimize();
	   }
	  else
	   {
		 TOptStep osStepper( this );
		 CArrayHyperIterator<CComObject<CSafetyPrecaution>*, int, __int64, TOptStep>
		   ahiIterator( arrPSP, iSz, m_iNSelectMax, osStepper, m_i64TotalOptCycles );

		 m_bC3Notified = false;
		 ahiIterator.Iterate();

		 if( m_otOptimizationType == OT_Quick2 && !CheckCancel() ) 
	       IncreaseAccuracyOfResult();
	   }

	  
	  Revert();
	  
      if( CheckCancel() ) 
	   {
	     if( !m_bC3Notified ) Work_CancelHelper3();		  
	     m_sOptResults.clear();		 
	     goto L_REENTER_POINT;
	   }
	  	  

	  shPercent = 100;
	  (this->*m_pfnNotify)( NM_OnStepOfWork3, &shPercent );

	  m_dtEnd3 = COleDateTime::GetCurrentTime();
	  dtEnd = m_dtEnd3 - m_dtSta3;
	  (this->*m_pfnNotify)( NM_OnEndOfWork3, &dtEnd );

L_REENTER_POINT:	  	  

	  delete[] arrPSP;
	  m_spCollSF_Opt.Release(); 
	  ::SuspendThread( m_tfFuncOpt.m_hthread );
	}
 }


STDMETHODIMP CMGertNet::get_IsRunning2(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	*pVal = CheckRunning2() ? VARIANT_TRUE:VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CMGertNet::get_IsRunningOpt(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	*pVal = CheckRunningOpt() ? VARIANT_TRUE:VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CMGertNet::get_OptimizResultsGetAndClear(SAFEARRAY** ppVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !ppVal ) return E_POINTER;
	if( m_sOptResults.size() < 1 )
	 {
       Error( L"CMGertNet::get_OptimizResults_GetAndClear: no results", IID_IMGertNet, E_FAIL );
       return E_FAIL;
	 }

    SAFEARRAYBOUND bnd = { m_sOptResults.size(), 0 };
	*ppVal = SafeArrayCreate( VT_DISPATCH, 1, &bnd );
	
	if( !*ppVal )
	 {
	   Error( L"Cann't array allocate", IID_IMGertNet, E_FAIL );
	   return E_FAIL;
	 }	

	IDispatch* pDta;
	HRESULT hr = SafeArrayAccessData( *ppVal, (void**)&pDta );
	if( FAILED(hr) )
	 {
	   SafeArrayDestroy( *ppVal );
	   *ppVal = NULL;	   
	   Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
	   return E_FAIL;
	 }
	memset( pDta, 0, sizeof(IDispatch*) * m_sOptResults.size() );


	TS_SetRes::reverse_iterator it1( m_sOptResults.rbegin() );
	TS_SetRes::reverse_iterator it2( m_sOptResults.rend() );
	for( ; it1 != it2; ++it1, ++pDta )
	 {
       CComPtr<ICollSF> spColl;
	   HRESULT hr = (*it1)->GetCollection( &spColl );
	   if( SUCCEEDED(hr) )
		 hr = spColl.QueryInterface( (IDispatch**)pDta ),
		 spColl.Release();

       if( FAILED(hr) )
		{
		  SafeArrayUnaccessData( *ppVal );
		  SafeArrayDestroy( *ppVal );
		  *ppVal = NULL;

		  Error( L"Cann't GetCollection", IID_IMGertNet, E_FAIL );
	      return E_FAIL;
		}	   
	 }
	

	SafeArrayUnaccessData( *ppVal );
	m_sOptResults.clear();

	return S_OK;
}

void TSampleData::SummMoney()
 {
   IT_TV_SF it1( m_vec.begin() );
   IT_TV_SF it2( m_vec.end() );
   for( m_ocSumm.SetCurrency(0, 0); it1 != it2; ++it1 )
	{
      CComObject<CSafetyPrecaution>* pSP = static_cast<CComObject<CSafetyPrecaution>*>( it1->p );
      m_ocSumm += pSP->m_ocCost;
	}
 }
void TSampleData::SummDQ()
 {
   IT_TV_SF it1( m_vec.begin() );
   IT_TV_SF it2( m_vec.end() );
   for( m_dSummDQ = 0; it1 != it2; ++it1 )
	{
      CComObject<CSafetyPrecaution>* pSP = static_cast<CComObject<CSafetyPrecaution>*>( it1->p );
      m_dSummDQ += pSP->m_dDeltaQ;
	}
 }

void TSampleData::Summ()
 {
   IT_TV_SF it1( m_vec.begin() );
   IT_TV_SF it2( m_vec.end() );
   for( m_dSummDQ = 0; it1 != it2; ++it1 )
	{
      CComObject<CSafetyPrecaution>* pSP = static_cast<CComObject<CSafetyPrecaution>*>( it1->p );
      m_dSummDQ += pSP->m_dDeltaQ;
	  m_ocSumm += pSP->m_ocCost;
	}
 }

HRESULT TSampleData::GetCollection( ICollSF** ppColl )
 {
   if( !ppColl ) return E_POINTER;

   CComObject<CCollSF>* pObj;
   HRESULT hr = CComObject<CCollSF>::CreateInstance( &pObj );
   if( FAILED(hr) )
	{
	  PassError( L"TSampleData::GetCollection: CComObject<>::CreateInstance", hr, CLSID_NULL, IID_IMGertNet );
	  return hr;
	}
   
   CComPtr<ICollSF> spTmp;
   hr = pObj->_InternalQueryInterface( IID_ICollSF, (void**)&spTmp );
   if( FAILED(hr) )
	{	  
	  delete pObj;
	  PassError( L"TSampleData::GetCollection: CComObject<>::_InternalQueryInterface", hr, CLSID_NULL, IID_IMGertNet );
	  return hr;
	}

   IT_TV_SF it1( m_vec.begin() );
   IT_TV_SF it2( m_vec.end() );
   for( ; it1 != it2; ++it1 )
	{
	  CComObject<CSafetyPrecaution>* pSP = static_cast<CComObject<CSafetyPrecaution>*>( it1->p );
      hr = spTmp->Add( it1->p, pSP->m_lID, NULL );
	  if( FAILED(hr) )
	   {	  
		 spTmp.Release();
		 PassError( L"TSampleData::GetCollection: Add error: ", hr, CLSID_NULL, IID_IMGertNet );
		 return hr;
	   }  
	}

   *ppColl = NULL;
   spTmp.CopyTo( ppColl );

   spTmp->put_DeltaQ( m_dSummDQ );
   spTmp->put_Cost( m_ocSumm );
   
   

   return S_OK;
 }


STDMETHODIMP CMGertNet::get_TimeWork3(DATE *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	*pVal = m_dtEnd3 - m_dtSta3;

	return S_OK;
}

STDMETHODIMP CMGertNet::CalibrateModel(double dP1, double dP2, double dP3, double dP4)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	if( CheckRunning() || CheckRunning2() )
	 {
       Error( L"CMGertNet::CalibrateModel: is already running", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	if( m_mtModelType != MT_Analytical && m_mtModelType != MT_AnalyticalIndistinct )
	 {
	   Error( L"CMGertNet::CalibrateModel: before using CalibrateModel set please RunMode=MT_Analytical", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }
    

	m_dP1 = dP1; m_dP2 = dP2; m_dP3 = dP3; m_dP4 = dP4;
	m_bCalibrateAnalytic = true;

	return InternalRun( 10, 100, VARIANT_FALSE, -1, VARIANT_FALSE );	
}

STDMETHODIMP CMGertNet::get_LastCalcError(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	if( m_bsError.empty() )
	  *pVal = SysAllocString( L"" );
	else
	  *pVal = SysAllocString( m_bsError.c_str() );
	m_bsError = L"";

	return *pVal == NULL ? E_OUTOFMEMORY:S_OK;
}

void MyLRCut( SET_TOutNumber2& rS, long lFreq, IT_SET_TOutNumber2& it1, IT_SET_TOutNumber2& it2 )
 {
   it1 = rS.begin(), it2 = rS.end(); it2--;
   //long ll = rS.size();
   for( ; it1 != it2 && it1->m_lCnt <= lFreq; ++it1 );
	/*{
	  int yy=1;
	}*/
   for( ; it1 != it2 && it2 != rS.begin() && it2->m_lCnt <= lFreq; --it2 );
	/*{
	  int yy=1;
	}*/

   if( it2 != rS.end() ) it2++;
   if( it1 == rS.end() ) it1 = rS.begin();	 
   if( it2 == rS.begin() ) it2 = rS.end();	 

   
   if( it1 == it2 ) it1 = rS.begin(), it2 = rS.end();	 
 }

//jjjjjjjjjjjjj
STDMETHODIMP CMGertNet::GetModelStat(short shIdx, TypANodesStat* pNS, SAFEARRAY** pparrSmpl, VARIANT_BOOL bFullGet )
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pNS || !pparrSmpl ) return E_POINTER;
	if( shIdx < 0 || shIdx >= sizeof(m_ansStat) / sizeof(m_ansStat[0]) )
	 {
	   Error( L"CMGertNet::GetModelStat: bad index", IID_IMGertNet, E_INVALIDARG );   
       return E_INVALIDARG;
	 }

	if( m_mtModelType == MT_Analytical2 )
	 {
	   Error( L"CMGertNet::GetModelStat: for MT_Analytical2 statistics is not accessible", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	if( (m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct) && 
	    !m_ansStat[ shIdx ].bFlInit )
	 {
	   Error( L"CMGertNet::GetModelStat: this statistic isn't initialized", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	if( m_sfStatOn == TSF_No )
	 {
	   Error( L"CMGertNet::GetModelStat: statistic is turned off", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	if( m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct )
	 {
	   pNS->dAvgP   = m_ansStat[ shIdx ].dAvgP;
	   pNS->dSummP  = m_ansStat[ shIdx ].dSummP;
	   pNS->dMinVal = m_ansStat[ shIdx ].dMinVal;
	   pNS->dMaxVal = m_ansStat[ shIdx ].dMaxVal;
	   pNS->cySz     = m_ansStat[ shIdx ].cySz;
	   pNS->dAvg    = m_ansStat[ shIdx ].dAvg;	
	   pNS->dDx     = m_ansStat[ shIdx ].dDx;	
	 }
	else
	 {
#if !defined(STAT_IDX0) && !defined(STAT_IDX)
	   long lMin, lMax;
	   double dMx, dDx;
	   GetInfSampleK( IndexToState(shIdx), &lMin, &lMax, &dMx, &dDx );


       pNS->dAvgP   = 0;
	   pNS->dSummP  = 0;
	   pNS->dMinVal = lMin;
	   pNS->dMaxVal = lMax;	   
	   pNS->dAvg    = dMx;
	   pNS->dDx     = dDx;	 

	   pNS->cySz.int64 = 0;
	   for( long l = 0; l < m_smSampleK.m_lClms; ++l )
	     pNS->cySz.int64 += m_smSampleK[ shIdx ][ l ];
#else
	   pNS->dAvgP   = 0;
	   pNS->dSummP  = 0;
	   pNS->dMinVal = m_ssSampleK[ shIdx ].dMinVal;
	   pNS->dMaxVal = m_ssSampleK[ shIdx ].dMaxVal;
	   pNS->cySz.int64 = m_ssSampleK[ shIdx ].l64Sz;
	   pNS->dAvg    = m_ssSampleK[ shIdx ].dAvg;	
	   pNS->dDx     = m_ssSampleK[ shIdx ].dM2;	 
#endif
	 }


	if( m_mtModelType == MT_Analytical || m_mtModelType == MT_AnalyticalIndistinct )
	 {

	   if( m_sfStatOn == TSF_Full && shIdx % 2 == 0 && bFullGet == VARIANT_TRUE )
		{
		  shIdx /= 2;

		  SAFEARRAYBOUND bnd = { m_vton[ shIdx ].size() * 4, 0 };
		  *pparrSmpl = SafeArrayCreate( VT_R8, 1, &bnd );

		  if( !*pparrSmpl )
		   {
			 Error( L"Cann't array allocate", IID_IMGertNet, E_FAIL );
			 return E_FAIL;
		   }

		  double* pDta;
		  HRESULT hr = SafeArrayAccessData( *pparrSmpl, (void**)&pDta );
		  if( FAILED(hr) )
		   {
			 SafeArrayDestroy( *pparrSmpl );
			 *pparrSmpl = NULL;	   
			 Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
			 return E_FAIL;
		   }


	  /*TPtyp m_tpP;
		 double m_shNumber1, m_shNumber2;
		 long m_lCnt;*/

		  for( int i = 0; i < m_vton[ shIdx ].size(); ++i )
			*pDta++ = m_vton[ shIdx ][ i ].m_tpP,
			*pDta++ = m_vton[ shIdx ][ i ].m_shNumber1,
			*pDta++ = m_vton[ shIdx ][ i ].m_shNumber2,
			//*pDta++ = m_vton[ shIdx ][ i ].m_shNumber1,
			*pDta++ = m_vton[ shIdx ][ i ].m_lCnt;

		  SafeArrayUnaccessData( *pparrSmpl );
		}
	   else
		{
		  //SAFEARRAYBOUND bnd = { 0, 0 };
		  //*pparrSmpl = SafeArrayCreate( VT_R8, 1, &bnd );
		}
	 }
	else
	 {
	   if( bFullGet == VARIANT_TRUE && m_sfStatOn == TSF_Full )
		{
		  IT_SET_TOutNumber2 itLeft, itRight;
		  MyLRCut( m_ston[shIdx], 1, itLeft, itRight );
		  long lSz = distance( itLeft, itRight );

		  SAFEARRAYBOUND bnd = { 4*lSz, 0 };
		  *pparrSmpl = SafeArrayCreate( VT_R8, 1, &bnd );

		  if( !*pparrSmpl )
		   {
			 Error( L"Cann't array allocate", IID_IMGertNet, E_FAIL );
			 return E_FAIL;
		   }

		  double* pDta;
		  HRESULT hr = SafeArrayAccessData( *pparrSmpl, (void**)&pDta );
		  if( FAILED(hr) )
		   {
			 SafeArrayDestroy( *pparrSmpl );
			 *pparrSmpl = NULL;	   
			 Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
			 return E_FAIL;
		   }
	  

		  for( ; itLeft != itRight; ++itLeft )
			*pDta++ = itLeft->m_lCnt,
			*pDta++ = itLeft->m_shNumber1,
			*pDta++ = itLeft->m_shNumber2,
			//*pDta++ = itLeft->m_shNumber1,
			*pDta++ = itLeft->m_tpP;

		  SafeArrayUnaccessData( *pparrSmpl );
		} 
	   else
		{
		}
	 }


	return S_OK;
}

STDMETHODIMP CMGertNet::get_IsDirty(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = (IsDirty() == S_OK ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CMGertNet::get_CalcSpeed(ThrdPriority *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	int iSpd = GetThreadPriority( m_tfFunc.m_hthread );
	switch( iSpd )
	 {
		case THREAD_PRIORITY_ABOVE_NORMAL:
		   *pVal = TP_ABOVE_NORMAL;
		   break;
		case THREAD_PRIORITY_BELOW_NORMAL:
		   *pVal = TP_BELOW_NORMAL;
		   break;
		case THREAD_PRIORITY_HIGHEST:
		   *pVal = TP_HIGHEST;
		   break;
		case THREAD_PRIORITY_LOWEST:
		   *pVal = TP_LOWEST;
		   break;
		case THREAD_PRIORITY_NORMAL:
		   *pVal = TP_NORMAL;
		   break;
		default:
		   return E_FAIL;
	 };
	

	return S_OK;
}

STDMETHODIMP CMGertNet::put_CalcSpeed(ThrdPriority newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	int iSpd;
	switch( newVal )
	 {
		case TP_ABOVE_NORMAL:
		   iSpd = THREAD_PRIORITY_ABOVE_NORMAL;
		   break;
		case TP_BELOW_NORMAL:
		   iSpd = THREAD_PRIORITY_BELOW_NORMAL;
		   break;
		case TP_HIGHEST:
		   iSpd = THREAD_PRIORITY_HIGHEST;
		   break;
		case TP_LOWEST:
		   iSpd = THREAD_PRIORITY_LOWEST;
		   break;
		case TP_NORMAL:
		   iSpd = THREAD_PRIORITY_NORMAL;
		   break;
		default:
		   return E_FAIL;
	 }; 

	SetThreadPriority( m_tfFunc.m_hthread, iSpd );
	SetThreadPriority( m_tfFuncCDQ.m_hthread, iSpd );
	SetThreadPriority( m_tfFuncOpt.m_hthread, iSpd );

	return S_OK;
}




STDMETHODIMP CMGertNet::get_OptimizationType(OptType *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( !pVal ) return E_POINTER;
	*pVal = m_otOptimizationType;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_OptimizationType(OptType newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	MPUT_PROPERTY( m_otOptimizationType, newVal );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_RemovLingvoOverwrap(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = (m_bRemovLingvoOverwrap ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CMGertNet::put_RemovLingvoOverwrap(VARIANT_BOOL newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    const bool bFl = (newVal == VARIANT_TRUE ? true:false);
	MPUT_PROPERTY( m_bRemovLingvoOverwrap, bFl );    

	return S_OK;
}

STDMETHODIMP CMGertNet::get_AverageDamage(CURRENCY *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	
	*pVal = m_cyAverageDamage;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_AverageDamage(CURRENCY newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

    try {
	   m_cyAverageDamage = newVal;
	 }
	catch( CException* pE )
	 {
	   pE->Delete();
       PassError( L"CMGertNet::put_MoneyForSF: ", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet );   
	   return E_INVALIDARG;
	 }

	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CMGertNet::get_MoneyForSF(CURRENCY *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_cyMoneyForSF;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_MoneyForSF(CURRENCY newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

 
    try {
	  m_cyMoneyForSF = newVal;
	 }
	catch( CException* pE )
	 {
	   pE->Delete();
       PassError( L"CMGertNet::put_MoneyForSF: ", E_INVALIDARG, GetObjectCLSID(), IID_IMGertNet );   
	   return E_INVALIDARG;
	 }

	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CMGertNet::Snapshot(VARIANT_BOOL bMake)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( bMake == VARIANT_FALSE )
	  m_vtfs.clear();
	else
	 {
       CComObject<CCollFac>* pCF = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );
	   CComObject<CCollFac>::TMAP::iterator it( pCF->m_coll.begin() );
	   
	   if( pCF->m_coll.size() < 1 )
		{
		  Error( L"CMGertNet::Snapshot: factor's collection is empty", IID_IMGertNet, E_FAIL );   
          return E_FAIL;
		}

	   if( m_vtfs.size() != pCF->m_coll.size() )
	     m_vtfs.clear(),
	     m_vtfs.assign( pCF->m_coll.size() );

	   for( int i = 0; it != pCF->m_coll.end(); ++it, ++i )
		{
		  CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it->second );
		  m_vtfs[ i ] = *pFac;
		}
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::Revert()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_vtfs.size() < 1 )
	 {
	   Error( L"CMGertNet::Revert: no snapshot present", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	CComObject<CCollFac>* pCF = static_cast<CComObject<CCollFac>*>( m_spCollFac.p );
	CComObject<CCollFac>::TMAP::iterator it( pCF->m_coll.begin() );
	if( pCF->m_coll.size() != m_vtfs.size() )
	 {
	   Error( L"CMGertNet::Revert: invalid shapshot size", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }
	
	for( int i = 0; it != pCF->m_coll.end(); ++it, ++i )
	 {
	   CComObject<CFactor>* pFac = static_cast<CComObject<CFactor>*>( it->second );
	   m_vtfs[ i ].LoadTo( *pFac );
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_UserProp(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_lUser;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_UserProp(long newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_lUser = newVal;

	return S_OK;
}

STDMETHODIMP CMGertNet::get_RndBase(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_lRndBase;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_RndBase(long newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    MPUT_PROPERTY( m_lRndBase, newVal );        

	return S_OK;
}

STDMETHODIMP CMGertNet::get_CurrIterStr(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	wchar_t wcBuff[ 1024 ];
	swprintf( wcBuff, L"%I64d", m_i64CurrIterOpt );
    *pVal = ::SysAllocString( wcBuff );

	return *pVal == NULL ? E_OUTOFMEMORY:S_OK;
}

STDMETHODIMP CMGertNet::get_TotalIterStr(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    wchar_t wcBuff[ 1024 ];
	swprintf( wcBuff, L"%I64d", m_i64TotalOptCycles	);
    *pVal = ::SysAllocString( wcBuff );

	return *pVal == NULL ? E_OUTOFMEMORY:S_OK;
}

STDMETHODIMP CMGertNet::get_StatOn(TStatFlag *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    *pVal = m_sfStatOn;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_StatOn(TStatFlag newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    m_sfStatOn = newVal;

	
	for( int i = 0; i < sizeof(m_vton) / sizeof(m_vton[0]); ++i )
	   m_vton[i].clear();

	for( i = 0; i < sizeof(m_ston) / sizeof(m_ston[0]); ++i )
	   m_ston[i].clear();
	 

	return S_OK;
}

STDMETHODIMP CMGertNet::get_StatIn(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    *pVal = (m_bStatIn ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CMGertNet::put_StatIn(VARIANT_BOOL newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    m_bStatIn = (newVal == VARIANT_TRUE ? true:false);

	return S_OK;
}

STDMETHODIMP CMGertNet::TstApplyTCH(IFactor *pFac, IFChange *pFC, TrustLevel *pLvl)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	//static void ApplyFCToFactor( CComObject<CFChange>* pFC, CComObject<CFactor>* pFactor, bool& bFlOverWrap )
    if( !pFac || !pFC || !pLvl ) return E_POINTER;
	
	*pLvl = ModifyTrustLvl( static_cast<CComObject<CFactor>*>(pFac)->m_tr, static_cast<CComObject<CFChange>*>(pFC)->m_tcTrustChange );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_ImitateCount( DATE* pVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	if( pVal )  
	 {
	   if( m_i64Count )	     	     
	     *pVal = double(m_i64Total)/double(m_i64Count) * (double)(COleDateTime::GetCurrentTime() - m_dtSta);
	   else *pVal = 0;
	 }
	else return E_POINTER;

	return S_OK;
 }


STDMETHODIMP CMGertNet::get_OptCount( DATE *pVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( pVal )  
	 {
	   if( m_i64CurrIterOpt )	     	     
	     *pVal = double(m_i64TotalOptCycles)/double(m_i64CurrIterOpt) * (double)(COleDateTime::GetCurrentTime() - m_dtSta3);
	   else *pVal = 0;
	 }
	else return E_POINTER;

   return S_OK;
 }


STDMETHODIMP CMGertNet::get_RngCount(DATE *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal )  
	 {
	   if( m_i64CountRng )	     	     
	     *pVal = double(m_i64TotalRng)/double(m_i64CountRng) * (double)(COleDateTime::GetCurrentTime() - m_dtSta2);
	   else *pVal = 0;
	 }
	else return E_POINTER;

	return S_OK;
}

STDMETHODIMP CMGertNet::get_ModuleProgID(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

    
	return m_bsModuleProgID.CopyTo( pVal );	
}

STDMETHODIMP CMGertNet::put_ModuleProgID(BSTR newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    m_bsModuleProgID = newVal;
	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CMGertNet::get_NDiv(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_sNDiv;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NDiv(short newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( newVal < 1 )
	 {
	   Error( L"CMGertNet::put_NDiv: invalid NDiv value, need >= 1", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	MPUT_PROPERTY( m_sNDiv, newVal );	

	return S_OK;
}

STDMETHODIMP CMGertNet::get_UnionThreshold(double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_dUnionThreshold;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_UnionThreshold(double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    
	MPUT_PROPERTY( m_dUnionThreshold, newVal );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_IntegralAccuracy(double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_dIntegralAccuracy;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_IntegralAccuracy(double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( newVal <= 0 )
	 {
	   Error( L"CMGertNet::put_IntegralAccuracy: invalid step value, need > 0", IID_IMGertNet, E_FAIL );   
       return E_FAIL;
	 }

	MPUT_PROPERTY( m_dIntegralAccuracy, newVal );	

	return S_OK;
}



STDMETHODIMP CMGertNet::get_NFormatPr(NumberFormat *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_nfFormatPr;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NFormatPr(NumberFormat newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_nfFormatPr != newVal )
	 {
	   m_nfFormatPr = newVal;
	   HRESULT hr = CreFormat( m_nfFormatPr, m_sNDigitsPr, m_bsFmtPr );
	   if( FAILED(hr) ) return hr;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_NFormatOther(NumberFormat *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_nfFormatOther;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NFormatOther(NumberFormat newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_nfFormatOther != newVal )
	 {
	   m_nfFormatOther = newVal;
	   HRESULT hr = CreFormat( m_nfFormatOther, m_sNDigitsOther, m_bsFmtOther );
	   if( FAILED(hr) ) return hr;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_NDigitsPr(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_sNDigitsPr;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NDigitsPr(short newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_sNDigitsPr != newVal )
	 {
	   m_sNDigitsPr = newVal;
	   HRESULT hr = CreFormat( m_nfFormatPr, m_sNDigitsPr, m_bsFmtPr );
	   if( FAILED(hr) ) return hr;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_NDigitsOther(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_sNDigitsOther;

	return S_OK;
}

STDMETHODIMP CMGertNet::put_NDigitsOther(short newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( m_sNDigitsOther != newVal )
	 {
	   m_sNDigitsOther = newVal;
	   HRESULT hr = CreFormat( m_nfFormatOther, m_sNDigitsOther, m_bsFmtOther );
	   if( FAILED(hr) ) return hr;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}


void __fastcall CMGertNet::InitFmts()
 {
   CreFormat( m_nfFormatPr, m_sNDigitsPr, m_bsFmtPr );
   CreFormat( m_nfFormatOther, m_sNDigitsOther, m_bsFmtOther );
 }


STDMETHODIMP CMGertNet::get_NFormat(NFormatTyp nf, BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	switch( nf )
	 {
	   case NFT_Pr:
		 return m_bsFmtPr.CopyTo( pVal );
	   case NFT_Other:
		 return m_bsFmtOther.CopyTo( pVal );
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_VParam(short Ventil, long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( Ventil < 1 || Ventil > 4 )
	 {
	   Error( L"CMGertNet::get_VParam: invalid Ventil number, need 1 - 4", IID_IMGertNet, E_FAIL );   
       return E_INVALIDARG;
	 }

	switch( Ventil )
	 {
	   case 1:
		 *pVal = m_lInTest1;
		 break;
	   case 2:
         *pVal = m_lInTest2;
		 break;
	   case 3:
		 *pVal = m_lInTest3;
		 break;
	   case 4:
		 *pVal = m_lInTest4;
		 break;
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::put_VParam(short Ventil, long newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( Ventil < 1 || Ventil > 4 )
	 {
	   Error( L"CMGertNet::put_VParam: invalid Ventil number, need 1 - 4", IID_IMGertNet, E_FAIL );   
       return E_INVALIDARG;
	 }

	switch( Ventil )
	 {
	   case 1:		 
		 MPUT_PROPERTY( m_lInTest1, newVal );
		 break;
	   case 2:
         MPUT_PROPERTY( m_lInTest2, newVal );
		 break;
	   case 3:
		 MPUT_PROPERTY( m_lInTest3, newVal );
		 break;
	   case 4:
		 MPUT_PROPERTY( m_lInTest4, newVal );
		 break;
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_VParamIndistinct(short Ventil, double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( Ventil < 1 || Ventil > 4 )
	 {
	   Error( L"CMGertNet::get_VParamIndistinct: invalid Ventil number, need 1 - 4", IID_IMGertNet, E_FAIL );   
       return E_INVALIDARG;
	 }

	switch( Ventil )
	 {
	   case 1:
		 *pVal = m_dInTest1D;
		 break;
	   case 2:
         *pVal = m_dInTest2D;
		 break;
	   case 3:
		 *pVal = m_dInTest3D;
		 break;
	   case 4:
		 *pVal = m_dInTest4D;
		 break;
	 }

	return S_OK;
}



STDMETHODIMP CMGertNet::put_VParamIndistinct(short Ventil, double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( Ventil < 1 || Ventil > 4 )
	 {
	   Error( L"CMGertNet::put_VParamIndistinct: invalid Ventil number, need 1 - 4", IID_IMGertNet, E_FAIL );   
       return E_INVALIDARG;
	 }

	switch( Ventil )
	 {
	   case 1:
		 MPUT_PROPERTY( m_dInTest1D, newVal );		 
		 break;
	   case 2:
         MPUT_PROPERTY( m_dInTest2D, newVal );
		 break;
	   case 3:
		 MPUT_PROPERTY( m_dInTest3D, newVal );
		 break;
	   case 4:
		 MPUT_PROPERTY( m_dInTest4D, newVal );
		 break;
	 }

	return S_OK;
}

STDMETHODIMP CMGertNet::get_VProbability(short Ventil, double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( Ventil < 1 || Ventil > 4 )
	 {
	   Error( L"CMGertNet::get_VProbability: invalid Ventil number, need 1 - 4", IID_IMGertNet, E_FAIL );   
       return E_INVALIDARG;
	 }

	*pVal = m_dArrPVent[ Ventil - 1 ];

	return S_OK;
}

STDMETHODIMP CMGertNet::put_VProbability(short Ventil, double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( Ventil < 1 || Ventil > 4 )
	 {
	   Error( L"CMGertNet::put_VProbability: invalid Ventil number, need 1 - 4", IID_IMGertNet, E_FAIL );   
       return E_INVALIDARG;
	 }	

	MPUT_PROPERTY( m_dArrPVent[ Ventil - 1 ], newVal );

	return S_OK;
}

STDMETHODIMP CMGertNet::get_ModuleConfig(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	
	return m_bsModuleCnf.CopyTo( pVal );
}

STDMETHODIMP CMGertNet::put_ModuleConfig(BSTR newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	
	m_bsModuleCnf = newVal;
	m_bRequiresSave = true;

	return S_OK;
}
