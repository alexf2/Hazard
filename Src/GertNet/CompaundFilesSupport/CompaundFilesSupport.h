
#ifdef COMPAUNDFILESSUPPORT_EXPORTS
  #define COMPAUNDFILESSUPPORT_API __declspec(dllexport)
#else
  #define COMPAUNDFILESSUPPORT_API __declspec(dllimport)
#endif


extern "C" {

COMPAUNDFILESSUPPORT_API void  __fastcall GetStgErrDescr( HRESULT hr, LPCOLESTR wcPref, LPCOLESTR bsName, CComBSTR& rBs ) throw();
COMPAUNDFILESSUPPORT_API HRESULT  __fastcall WIN32ERR_TO_COMERR( std::basic_string<TCHAR>& rStr, DWORD dwErr ) throw();
COMPAUNDFILESSUPPORT_API HRESULT __fastcall ReportStgError( HRESULT h, LPCOLESTR bsNameFunc, const CLSID& ridCls, const IID& iid, LPCOLESTR bsNameFile = NULL ) throw();


COMPAUNDFILESSUPPORT_API bool __fastcall FloatIntoVariant( VARIANT* pVar, const float fArg ) throw();
COMPAUNDFILESSUPPORT_API bool __fastcall ShortIntoVariant( VARIANT* pVar, const short sArg ) throw();
COMPAUNDFILESSUPPORT_API HRESULT  __fastcall WSTRIntoVariant( VARIANT* pVar, LPCWSTR bsArg ) throw();

COMPAUNDFILESSUPPORT_API bool __fastcall LongFromVariant( const VARIANT* pVar, long& lArg ) throw();
COMPAUNDFILESSUPPORT_API bool __fastcall VBoolFromVariant( const VARIANT* pVar, VARIANT_BOOL& vbArg ) throw();
COMPAUNDFILESSUPPORT_API bool __fastcall FloatFromVariant( const VARIANT* pVar, float& lArg ) throw();

COMPAUNDFILESSUPPORT_API BSTR* __fastcall GetBSTRDst( VARIANT* pVar, HRESULT& rHr ) throw();
COMPAUNDFILESSUPPORT_API float* __fastcall GetFloatDst( VARIANT* pVar, HRESULT& rHr ) throw();
COMPAUNDFILESSUPPORT_API double* __fastcall GetDoubleDst( VARIANT* pVar, HRESULT& rHr ) throw();

COMPAUNDFILESSUPPORT_API bool __fastcall IsEmpty( IStorage* ) throw();

COMPAUNDFILESSUPPORT_API HRESULT __fastcall EndTransaction( IStorage*, HRESULT, STGC stgc = (STGC)(STGC_DEFAULT|STGC_ONLYIFCURRENT) ) throw();
COMPAUNDFILESSUPPORT_API HRESULT __fastcall ObjExists( IStorage*, DWORD, STGTY ) throw();

 };
