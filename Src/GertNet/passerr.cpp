#include "stdafx.h"
#include "PassErr.h"

#include <tchar.h>

#include <sstream>
#include <string>
#include <iomanip>
using namespace std;



void PassError( LPCOLESTR lp, HRESULT hr, const CLSID& clsid, const IID& iid )
 {
   //USES_CONVERSION;
   std::basic_stringstream<WCHAR> strm;
   strm << std::setbase(16);
   if( lp ) strm << lp << L"\n";

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

   //Error( strm.str().c_str(), IID_ILingvoEnum, hr );   
   AtlReportError( clsid, strm.str().c_str(), iid, hr );
 }
