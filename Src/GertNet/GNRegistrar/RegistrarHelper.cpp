// RegistrarHelper.cpp : Implementation of CRegistrarHelper
#include "stdafx.h"
#include <statreg.h>
#include "GNRegistrar.h"
#include "RegistrarHelper.h"

namespace ATL {


//AWPlugin
#define IDR_COLLTOPICS                  103
#define IDR_COLLFTABLES                 106
#define IDR_INTCREATOR                  107
#define IDR_COLLFACTORS                 108

//MGertNet
#define IDR_LINGVOENUM                  102
#define IDR_COLLLINGVO                  103
#define IDR_COMPAUNDFILES               104
#define IDR_FACTOR                      105
#define IDR_COLLFAC                     106
#define IDR_MGERTNET                    107
#define IDR_SAFETYPRECAUTION            108
#define IDR_FCHANGE                     109
#define IDR_ICOLLFCHANGE                110
#define IDR_COLLSF                      111
#define IDR_CATREGISTRAR                112


//#AngleFont
#define IDR_ANGLEDFONT                  101
#define IDR_FACEDITSINK                 102


/////////////////////////////////////////////////////////////////////////////
// CRegistrarHelper

STDMETHODIMP CRegistrarHelper::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IRegistrarHelper
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

HRESULT CRegistrarHelper::FinalConstruct()
 {
   //HRESUL hr = CComObject<CDLLRegObject>::CreateInstance( &m_pReg );
   //return hr;
 return S_OK;
 }

struct lessBSTRNoCase: binary_function<CComBSTR, CComBSTR, bool> 
 {
   bool __fastcall operator()(const CComBSTR& _X, const CComBSTR& _Y) const throw()
	{
	  if( _Y.m_str == NULL && _X.m_str == NULL )
		  return false;
	  if( _Y.m_str != NULL && _X.m_str != NULL )
		  return _wcsicmp( _X.m_str, _Y.m_str ) < 0;
	  return _X.m_str == NULL;	
	}
 };


typedef auto_ptr< set<DWORD> > AP_SET;
typedef map<CComBSTR, AP_SET, lessBSTRNoCase > TMAP_DIR;
typedef TMAP_DIR::iterator IT_TMAP_DIR;
typedef TMAP_DIR::const_iterator CIT_TMAP_DIR;
typedef TMAP_DIR::value_type MVT;

AP_SET& operator+( AP_SET& rS, DWORD dw )
 {
   rS->insert( dw );
   return rS;
 }

TMAP_DIR& operator+( TMAP_DIR& rM, MVT& rI )
 {
   rM.insert( rI );
   return rM;
 }

static TMAP_DIR mDir;

struct TInitializer
 {
   TInitializer()
	{
	  mDir + MVT( L"AWMonitor", AP_SET( new set<DWORD>() ) + 
			   IDR_COLLTOPICS + IDR_COLLFTABLES + IDR_INTCREATOR + IDR_COLLFACTORS );

	  mDir + MVT( L"MGertNet", AP_SET( new set<DWORD>() ) + 
			   IDR_LINGVOENUM + IDR_COLLLINGVO + IDR_COMPAUNDFILES +
			   IDR_FACTOR + IDR_COLLFAC + IDR_MGERTNET + IDR_SAFETYPRECAUTION +
			   IDR_FCHANGE + IDR_ICOLLFCHANGE + IDR_COLLSF + IDR_CATREGISTRAR );

	  mDir + MVT( L"ANgleFont", AP_SET( new set<DWORD>() ) + 
				IDR_ANGLEDFONT + IDR_FACEDITSINK );
	}

   ~TInitializer()
	{
	  mDir.clear(); 
	}
 } initMe;

static void __fastcall GetErrInfo( std::basic_string<wchar_t>& rStr, DWORD dwErr )
 { 
   USES_CONVERSION;

   LPVOID lpMsgBuf;   
   DWORD dw = FormatMessage( 
	 FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
	 NULL,
	 dwErr,
	 MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
	 (LPTSTR) &lpMsgBuf,
	 0,
	 NULL 
   );

   rStr = dw ? A2W((char*)lpMsgBuf):L"<Ошибка извлечения информации об ошибке>";   
   
   LocalFree( lpMsgBuf );   
 }


STDMETHODIMP CRegistrarHelper::Register( BSTR ObjName, BSTR Path, short RegFlag, short TypeLibFlag, BSTR PathHelp )
 {
    HRESULT hr;
    
	
	if( Path == NULL ) 
	  hr = E_POINTER,
	  SetErr( hr, L"Нулевой указатель" );
	else
	 {
	   //hr = m_pReg->AddReplacement( L"Module", Path );
	   hr = mReg.AddReplacement( L"Module", Path );
	   if( SUCCEEDED(hr) )
		{
		  CIT_TMAP_DIR it = mDir.find( ObjName );
		  if( it == mDir.end() )
		   {
		     hr = E_FAIL;
			 basic_stringstream<wchar_t> strm;
			 strm << L"Нельзя зарегистрировать/удалить: объект '" << ObjName << L"' не найден";
			 //Error( strm.str().c_str(), IID_IRegistrarHelper, hr );
			 SetErr( hr, strm.str().c_str() );
		   }
		  else
		   {
		     set<DWORD>::const_iterator it2( it->second->begin() );
			 for( ; it2 != it->second->end(); ++it2 )
			  {
			    if( RegFlag ) 
			      hr = mReg.ResourceRegister( Path, *it2, L"REGISTRY" );
				else
				  hr = mReg.ResourceUnregister( Path, *it2, L"REGISTRY" );

				if( FAILED(hr) )
				  {
					basic_stringstream<wchar_t> strm;
					strm << L"Ошибка регистрации/удаления '" << ObjName << L"' [" << (long)*it2 << L"]";
					//Error( strm.str().c_str(), IID_IRegistrarHelper, hr );
					SetErr( hr, strm.str().c_str() );
					break;
				  }

                if( TypeLibFlag )
				 {
				   CComPtr<ITypeLib> spTLib;
				   hr = LoadTypeLib( Path, &spTLib );
				   if( SUCCEEDED(hr) )
					{
					  if( RegFlag )
					    hr = RegisterTypeLib( spTLib, Path, PathHelp ); 
					  else
					   {
					     TLIBATTR* ptla;
		                 hr = spTLib->GetLibAttr( &ptla );
						 if( SUCCEEDED(hr) )						  
					        hr = UnRegisterTypeLib( ptla->guid, ptla->wMajorVerNum, ptla->wMinorVerNum, ptla->lcid, ptla->syskind );
						  
						 spTLib->ReleaseTLibAttr( ptla );
					   }
					}
				   if( FAILED(hr) )
					{
					  std::basic_string<wchar_t> str;
					  GetErrInfo( str, hr & 0x0000FFFF );
					  SetErr( hr, str.c_str() );
					}
				 }
			  }
		   }

		  HRESULT hr2 = mReg.ClearReplacements();	      
		  if( SUCCEEDED(hr) && FAILED(hr2) ) hr = hr2;
		}
	 }

	ClearError();
	return hr;
 }

}

STDMETHODIMP CRegistrarHelper::get_Hresult(long *pVal)
{
	*pVal = m_hr;

	return S_OK;
}

STDMETHODIMP CRegistrarHelper::get_LastError(BSTR *pVal)
{	
	return m_bsErr.CopyTo( pVal );
}
