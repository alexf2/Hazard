// IntCreator.cpp : Implementation of CIntCreator
#include "stdafx.h"
#include "AWPlugin.h"
#include "IntCreator.h"

#include "FactorTable.h"
#include "GostTable.h"
#include "CollGosts.h"

#include "IntCreatorHelper.h"

/*LPCWSTR FactorTable_CName = L"CFactorTable";
LPCWSTR FactorTable_IName = L"IFactorTable";

LPCWSTR GostTable_CName = L"CGostTable";
LPCWSTR GostTable_IName = L"IGostTable";*/


typedef  CPNCreator< CIntCreator, IFactorTable, CFactorTable, &IID_IIntCreator > CreIFactorTable;
typedef  CPNCreator< CIntCreator, IGostTable, CGostTable, &IID_IIntCreator > CreIGostTable;
typedef  CPNCreatorStg< CIntCreator, ICollGosts, CCollGosts, &IID_ICollGosts > CreICollGosts;


/////////////////////////////////////////////////////////////////////////////
// CIntCreator

STDMETHODIMP CIntCreator::InterfaceSupportsErrorInfo(REFIID riid) throw()
{
	static const IID* arr[] = 
	{
		&IID_IIntCreator
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}



STDMETHODIMP CIntCreator::NewIFactorTable(BSTR Name,/*[out, retval]*/ IFactorTable** FactorTable) throw()
 {
   static int iCnt = 0;

   CreIFactorTable creator;
   HRESULT hr = creator.NewInst( this, FactorTable, L"CFactorTable", L"IFactorTable" ); 
   if( SUCCEEDED(hr) )
	{	  
	  CComObject<CFactorTable>* pFT = static_cast< CComObject<CFactorTable>* >( *FactorTable );
	  pFT->m_lKey = iCnt++;

	  //basic_stringstream<WCHAR> stm;
	  //stm << L"Name " << pFT->m_lKey;
	  pFT->m_bsName = Name;
	}
   return hr;
 }

STDMETHODIMP CIntCreator::NewIGostTable(BSTR Name, IGostTable **GostTable) throw()
 {
    static int iCnt = 0;

	CreIGostTable creator;
    HRESULT hr = creator.NewInst( this, GostTable, L"CGostTable", L"IGostTable" ); 
	if( SUCCEEDED(hr) )
	{	  
	  CComObject<CGostTable>* pFT = static_cast< CComObject<CGostTable>* >( *GostTable );
	  pFT->m_lKey = iCnt++;

	  //basic_stringstream<WCHAR> stm;
	  //stm << L"Name " << pFT->m_lKey;
	  pFT->m_bsName = Name;
	}
   return hr;
 }

//CPNCreatorStg
STDMETHODIMP CIntCreator::NewICollGosts(BSTR Name, IStorage *Stg, ICollGosts **pCollGosts)
{
	static int iCnt = 0;

	CreICollGosts creator;
    HRESULT hr = creator.NewInst( this, Stg, pCollGosts, L"CCollGosts", L"ICollGosts" );
	if( SUCCEEDED(hr) )
	{	  
	  CComObject<CCollGosts>* pFT = static_cast< CComObject<CCollGosts>* >( *pCollGosts );
	  pFT->m_lKey = iCnt++;

	  //basic_stringstream<WCHAR> stm;
	  //stm << L"Name " << pFT->m_lKey;
	  pFT->m_bsName = Name;
	}
   return hr;
}

STDMETHODIMP CIntCreator::FetchGostTable( /*[in]*/IStorage* StmRoot, /*[in]*/SAFEARRAY** Path, /*[in]*/long OFThrough, /*[in,optional]*/VARIANT* OFEnd, /*[in,optional]*/VARIANT* CreCached, /*[out,retval]*/ IGostTable** pGT ) throw()
 {
   if( StmRoot == NULL || OFEnd == NULL || pGT == NULL || CreCached == NULL || Path == NULL ) return E_POINTER;

     
   DWORD dwOFEnd;
   if( (V_VT(OFEnd)&VT_TYPEMASK) == VT_ERROR ) dwOFEnd = OFThrough;
   else
	{
	  long lTmp;
	  if( !LongFromVariant(OFEnd, lTmp) ) return E_INVALIDARG;
	  dwOFEnd = lTmp;
	}  

   VARIANT_BOOL vbCreCached;
   if( (V_VT(CreCached)&VT_TYPEMASK) == VT_ERROR ) vbCreCached = VARIANT_TRUE;
   else
	 if( !VBoolFromVariant(CreCached, vbCreCached) ) return E_INVALIDARG;
	  	     

   long lSz;
   HRESULT hr;
   hr = SafeArrayLock( *Path );
   if( SUCCEEDED(hr) ) lSz = (*Path)->rgsabound[0].cElements + 1;
   else return hr;
   SafeArrayUnlock( *Path );

   CComPtr<IStream>* spTmp = new (nothrow) CComPtr<IStream>[ lSz ];

   
   if( spTmp )
	{   
	  hr = Fetch( StmRoot, OFThrough, dwOFEnd, Path, true, (CComPtr<IUnknown>*)spTmp );
	  if( SUCCEEDED(hr) )
	   {
		 CreIGostTable creator;
		 hr = creator.NewLoad( this, *(spTmp + (lSz - 1)), vbCreCached == VARIANT_TRUE ? true:false, pGT,  L"CGostTable", L"IGostTable" ); 
	   }

	  if( FAILED(hr) )
		ReportStgError( hr, L"IIntCreator.FetchFactorTable", GetObjectCLSID(), IID_IIntCreator );

	  delete[] spTmp;
	}
   else hr = E_OUTOFMEMORY;

   return hr;
 }

STDMETHODIMP CIntCreator::FetchFactorTable( /*[in]*/IStorage* StmRoot, /*[in]*/SAFEARRAY** Path, /*[in]*/long OFThrough, /*[in,optional]*/VARIANT* OFEnd, /*[in,optional]*/VARIANT* CreCached, /*[out,retval]*/ IFactorTable** pFT ) throw()
 {
   if( StmRoot == NULL || OFEnd == NULL || pFT == NULL || CreCached == NULL || Path == NULL ) return E_POINTER;

     
   DWORD dwOFEnd;
   if( (V_VT(OFEnd)&VT_TYPEMASK) == VT_ERROR ) dwOFEnd = OFThrough;
   else
	{
	  long lTmp;
	  if( !LongFromVariant(OFEnd, lTmp) ) return E_INVALIDARG;
	  dwOFEnd = lTmp;
	}  

   VARIANT_BOOL vbCreCached;
   if( (V_VT(CreCached)&VT_TYPEMASK) == VT_ERROR ) vbCreCached = VARIANT_TRUE;
   else
	 if( !VBoolFromVariant(CreCached, vbCreCached) ) return E_INVALIDARG;

   long lSz;
   HRESULT hr;
   hr = SafeArrayLock( *Path );
   if( SUCCEEDED(hr) ) lSz = (*Path)->rgsabound[0].cElements + 1;
   else return hr;
   SafeArrayUnlock( *Path );

   CComPtr<IStream>* spTmp = new (nothrow) CComPtr<IStream>[ lSz ];

   
   if( spTmp )
	{
	  hr = Fetch( StmRoot, OFThrough, dwOFEnd, Path, true, (CComPtr<IUnknown>*)spTmp );
	  if( SUCCEEDED(hr) )
	   {
		 CreIFactorTable creator;
		 hr = creator.NewLoad( this, *(spTmp + (lSz - 1)), vbCreCached == VARIANT_TRUE ? true:false, pFT,  L"CFactorTable", L"IFactorTable" ); 
	   }

	  if( FAILED(hr) )
		ReportStgError( hr, L"IIntCreator.FetchFactorTable", GetObjectCLSID(), IID_IIntCreator );

	  delete[] spTmp;
	}
   else hr = E_OUTOFMEMORY;

   return hr;
 }

HRESULT __fastcall CIntCreator::Fetch( IStorage* StmRoot, DWORD OpenFlagsT, DWORD OpenFlagsE, SAFEARRAY** Path, bool bLastIsStream, CComPtr<IUnknown>* pRes ) throw()
 {
   if( StmRoot == NULL || Path == NULL || pRes == NULL ) return E_POINTER;
   
   CComPtr<IUnknown>* pKeep = pRes;

   HRESULT hr;
   bool bLocked = false;
   hr = SafeArrayLock( *Path );
   if( FAILED(hr) )
	 Error( L"Cann't array access", IID_IIntCreator, hr );
   else
	{
	  bLocked = true;
	  if( (*Path)->cDims != 1 )	   
		Error( L"Array must have only 1 dimension", IID_IIntCreator, hr );	   
	  else
	   {
	     if( ((*Path)->fFeatures&FADF_BSTR) != FADF_BSTR )		
		   Error( L"Bad array type of: Array must be a VARIANTs array", IID_IIntCreator, hr );		  
		 else
		  {
		    if( (*Path)->rgsabound[0].cElements < 1 )		
			  Error( L"Array must have even one element or more", IID_IIntCreator, hr );			 
			else
			 {				   
			   BSTR* pDta = (BSTR*)(*Path)->pvData;
			   //CComPtr<IUnknown> spStorage( (IUnknown*)StmRoot );
			   pRes->Attach( (IUnknown*)StmRoot );
			   for( int i = (*Path)->rgsabound[0].cElements; i > 0; --i, ++pDta )
				if( i == 1 && bLastIsStream )
				 {
				   //CComPtr<IStream> spTmpStream;
				   hr = reinterpret_cast<IStorage*>((pRes-1)->p)->OpenStream( *pDta, NULL, OpenFlagsE, 0, (IStream**)&((++pRes)->p) );
				   if( FAILED(hr) ) break;					
				 }
				else
				 {
				   //CComPtr<IStorage> spTmpStorage;
				   hr = reinterpret_cast<IStorage*>((pRes-1)->p)->OpenStorage( *pDta, NULL, i == 1 ? OpenFlagsE:OpenFlagsT, 0, 0, (IStorage**)&((++pRes)->p) );
				   if( FAILED(hr) ) break;
				 }
			 }
		  }
	   }
	}

   pKeep->Detach(); 

   if( bLocked ) SafeArrayUnlock( *Path );
   return hr;
 }

