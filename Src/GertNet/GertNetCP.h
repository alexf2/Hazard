#ifndef _GERTNETCP_H_
#define _GERTNETCP_H_

#include <map>
//#include "GertNet.h"
//#include <otlbase.h>

#include "MyGIT.h"

typedef std::map<DWORD, DWORD> TMapCook;

template <class T>
class CProxy_IGertNetEvents : public IConnectionPointImpl<T, &DIID__IGertNetEvents, CComDynamicUnkArray>
{
	//Warning this class may be recreated by the wizard.
public:

    typedef IConnectionPointImpl<T, &DIID__IGertNetEvents, CComDynamicUnkArray> TCPIBase;

	struct TNM_DispID
	 {
	   NotifyMsg nm;
       DISPID id;
	 };
	static TNM_DispID nmdArr[ 9 ];
	 

	DISPID __fastcall FindID( NotifyMsg nm )
	 {
	   for( int i = sizeof(nmdArr) / sizeof(nmdArr[0]) - 1; i >= 0; --i )
		 if( nm == nmdArr[ i ].nm ) return nmdArr[ i ].id;
	   return 0;
	 }

	HRESULT __fastcall Fire_OnEndOfWork(DATE dtTime, NotifyMsg nmMsg )
	{
		CComVariant varResult;
		T* pT = static_cast<T*>(this);
		//int nConnectionIndex;
		CComVariant* pvars = new CComVariant[1];
		pvars[0].date = dtTime;
		pvars[0].vt = VT_DATE;
		DISPPARAMS disp = { pvars, NULL, 1, 0 };
		//int nConnections = m_vec.GetSize();
		TMapCook::iterator it1( m_mc.begin() );
		TMapCook::iterator it2( m_mc.end() );		 
		
		//for (nConnectionIndex = 0; nConnectionIndex < nConnections; nConnectionIndex++)
		for( ; it1 != it2; ++it1 )
		{
			pT->Lock();
			//CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
			CComPtr<_IGertNetEvents> spGNE;
			HRESULT hr = m_ogGIT.GetInterface( it1->second, DIID__IGertNetEvents, (void**)&spGNE );
			if( FAILED(hr) )
			 {
			   pT->Unlock();
			   continue;
			 }
			pT->Unlock();
			IDispatch* pDispatch = reinterpret_cast<IDispatch*>(spGNE.p);
			if (pDispatch != NULL)
			{
				varResult.Clear();
				pDispatch->Invoke( FindID(nmMsg), IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
			}
		}
		delete[] pvars;		
		return varResult.vt == VT_EMPTY ? S_OK:varResult.scode;	
	}

	HRESULT __fastcall Fire_OnStepOfWork( short shPercent, NotifyMsg nmMsg )
	{
		CComVariant varResult;
		T* pT = static_cast<T*>(this);
		//int nConnectionIndex;
		CComVariant* pvars = new CComVariant[1];
		pvars[0].iVal = shPercent;
		pvars[0].vt = VT_I2;
		DISPPARAMS disp = { pvars, NULL, 1, 0 };
		//int nConnections = m_vec.GetSize();
		TMapCook::iterator it1( m_mc.begin() );
		TMapCook::iterator it2( m_mc.end() );		 
		
		//for (nConnectionIndex = 0; nConnectionIndex < nConnections; nConnectionIndex++)
		for( ; it1 != it2; ++it1 )
		{
			pT->Lock();
			//CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
			CComPtr<_IGertNetEvents> spGNE;
			HRESULT hr = m_ogGIT.GetInterface( it1->second, DIID__IGertNetEvents, (void**)&spGNE );
			if( FAILED(hr) )
			 {
			   pT->Unlock();
			   continue;
			 }
			pT->Unlock();
			IDispatch* pDispatch = reinterpret_cast<IDispatch*>(spGNE.p);
			if (pDispatch != NULL)
			 {
				varResult.Clear();
				pDispatch->Invoke( FindID(nmMsg), IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
			 }
		}
		delete[] pvars;		
		return varResult.vt == VT_EMPTY ? S_OK:varResult.scode;	
	}

	HRESULT __fastcall Fire_OnCancel(DATE dtTime, NotifyMsg nmMsg)
	{
		CComVariant varResult;
		T* pT = static_cast<T*>(this);
		//int nConnectionIndex;
		CComVariant* pvars = new CComVariant[1];
		pvars[0].date = dtTime;
		pvars[0].vt = VT_DATE;
		DISPPARAMS disp = { pvars, NULL, 1, 0 };
		//int nConnections = m_vec.GetSize();
		TMapCook::iterator it1( m_mc.begin() );
		TMapCook::iterator it2( m_mc.end() );		 
		
		//for (nConnectionIndex = 0; nConnectionIndex < nConnections; nConnectionIndex++)
		for( ; it1 != it2; ++it1 )
		{
			pT->Lock();
			//CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
			CComPtr<_IGertNetEvents> spGNE;
			HRESULT hr = m_ogGIT.GetInterface( it1->second, DIID__IGertNetEvents, (void**)&spGNE );
			if( FAILED(hr) )
			 {
			   pT->Unlock();
			   continue;
			 }
			pT->Unlock();
			IDispatch* pDispatch = reinterpret_cast<IDispatch*>(spGNE.p);
			if (pDispatch != NULL)
			{
				varResult.Clear();
				pDispatch->Invoke(FindID(nmMsg), IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
			}
		}
		delete[] pvars;		
		return varResult.vt == VT_EMPTY ? S_OK:varResult.scode;	
	}

	HRESULT __fastcall Fire_OnErrorCalc( BSTR bsErrMsg, NotifyMsg nmMsg )
	{
		CComVariant varResult;
		T* pT = static_cast<T*>(this);
		//int nConnectionIndex;				
		CComVariant* pvars = new CComVariant[1];
		pvars[0].bstrVal = bsErrMsg;
		pvars[0].vt = VT_BSTR;
		DISPPARAMS disp = { pvars, NULL, 1, 0 };
		//int nConnections = m_vec.GetSize();
		TMapCook::iterator it1( m_mc.begin() );
		TMapCook::iterator it2( m_mc.end() );		 
		
		//for (nConnectionIndex = 0; nConnectionIndex < nConnections; nConnectionIndex++)
		for( ; it1 != it2; ++it1 )
		{
			pT->Lock();
			//CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
			CComPtr<_IGertNetEvents> spGNE;
			HRESULT hr = m_ogGIT.GetInterface( it1->second, DIID__IGertNetEvents, (void**)&spGNE );
			if( FAILED(hr) )
			 {
			   pT->Unlock();
			   continue;
			 }
			pT->Unlock();
			IDispatch* pDispatch = reinterpret_cast<IDispatch*>(spGNE.p);
			if (pDispatch != NULL)
			{
				varResult.Clear();
				pDispatch->Invoke(FindID(nmMsg), IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
			}
		}
		pvars[0].bstrVal = NULL;
		pvars[0].vt = VT_EMPTY; 
		delete[] pvars;		
		return varResult.vt == VT_EMPTY ? S_OK:varResult.scode;	
	}

	STDMETHOD(Advise)(IUnknown* pUnkSink, DWORD* pdwCookie)
	 {
	   T* pT = static_cast<T*>(this);
	   HRESULT hr = TCPIBase::Advise( pUnkSink, pdwCookie );
	   pT->Lock();

       if( SUCCEEDED(hr) )
		{
		  pair<TMapCook::iterator, bool> pRes = 
		    m_mc.insert( TMapCook::value_type(*pdwCookie, 0) );
		  if( pRes.second == false ) hr = E_FAIL;

		  IUnknown* pEvent = CComDynamicUnkArray::GetUnknown( *pdwCookie );
		  DWORD dwCook2;
		  hr = m_ogGIT.RegisterInterface( pEvent, DIID__IGertNetEvents, &dwCook2 );
		  if( FAILED(hr) )
		   {
		     m_mc.erase( pRes.first );
			 pT->Unlock();
			 return hr;
		   }

		  pRes.first->second = dwCook2;
		}
	   pT->Unlock();
	   return hr;
	 }

	STDMETHOD(Unadvise)(DWORD dwCookie)
	 {
	   T* pT = static_cast<T*>(this);	   
	   HRESULT hr = TCPIBase::Unadvise( dwCookie );
	   pT->Lock();
       if( SUCCEEDED(hr) )
		{
		  TMapCook::iterator it = m_mc.find( dwCookie );
		  if( it != m_mc.end() )	
		   {
		     hr = m_ogGIT.RevokeInterface( it->second );
			 m_mc.erase( it );
		   }
		  else hr = E_FAIL;			  		   
		}
	   pT->Unlock();
	   return hr;
	 }

private:
  TMapCook m_mc;  
};

template<class T>
CProxy_IGertNetEvents<T>::TNM_DispID CProxy_IGertNetEvents<T>::nmdArr[ 9 ] =
 {
   { NM_OnEndOfWork,   1 },
   { NM_OnStepOfWork,  2 },
   { NM_OnCancel,      3 },

   { NM_OnEndOfWork2,  5 },
   { NM_OnStepOfWork2, 6 },

   { NM_OnEndOfWork3,  7 },
   { NM_OnStepOfWork3, 8 },
   { NM_OnCancel3,     9 },
   { NM_OnErrorCalc,   4 }
 };

#endif
