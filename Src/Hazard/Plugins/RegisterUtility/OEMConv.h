#if !defined(OEM_Conv)
#define OEM_Conv

#include <windows.h>
#include <crtdbg.h>
#include <wincon.h>
#include <malloc.h>
#include <string.h>

namespace OEMConvNS {

inline LPSTR AnsiToOEMHelper( LPCSTR lpSrc, LPSTR lpTrg, int iNChars )
 {
   _ASSERTE( lpSrc != NULL );
   _ASSERTE( lpTrg != NULL );
   _ASSERTE( iNChars >= 0 );
   
   CharToOemBuff( lpSrc, lpTrg, iNChars );
   lpTrg[ iNChars ] = 0;
   
   return lpTrg;
 }

inline LPSTR OEMToAnsiHelper( LPCSTR lpSrc, LPSTR lpTrg, int iNChars )
 {
   _ASSERTE( lpSrc != NULL );
   _ASSERTE( lpTrg != NULL );
   _ASSERTE( iNChars >= 0 );
   
   OemToCharBuff( lpSrc, lpTrg, iNChars );
   lpTrg[ iNChars ] = 0;
   
   return lpTrg;
 }


inline LPSTR WToAHelper( LPCWSTR lpwSrc, LPSTR lpTrg, int iNBytesInBuff, UINT uiACP )
 {
   _ASSERTE( lpwSrc != NULL );
   _ASSERTE( lpTrg != NULL );
   _ASSERTE( iNBytesInBuff >= 2 );
   
   int iRes = WideCharToMultiByte( uiACP, 0, lpwSrc, -1, lpTrg, iNBytesInBuff, NULL, NULL );
   if( !iRes && iNBytesInBuff > 2 )
	{
	  *lpTrg = 0;
	  switch( GetLastError() )
	   {
	     case ERROR_INSUFFICIENT_BUFFER:
		   _RPTF1( _CRT_ERROR, "Выходной буфер [%u] слишком мал", iNBytesInBuff );
		   break;
         case ERROR_INVALID_FLAGS:
		   _RPTF0( _CRT_ERROR, "Ошибочные флаги у WideCharToMultiByte" );
		   break;
		 case ERROR_INVALID_PARAMETER:
		   _RPTF0( _CRT_ERROR, "Ошибочный параметр у WideCharToMultiByte" );
		   break;
		 default:
		   _RPTF0( _CRT_ERROR, "Неизвестная ошибка" );
	   };
	}
   else lpTrg[ iRes ] = 0;
   
   return lpTrg;
 }

 }; //namespace OEMConvNS


#define USES_OEM_CONVERSION\
   int _iConvertSize; _iConvertSize

#define USES_OEM_X_CONVERSION\
   int _iBuffSize; _iBuffSize;\
   UINT _uiACP = CP_ACP; _uiACP

#define USES_OEM_XCONSOLE_CONVERSION\
   int _iBuffSize; _iBuffSize;\
   UINT _uiACP = GetConsoleOutputCP(); _uiACP



#define A2O( lpStr ) (\
  (lpStr == NULL) ? NULL: (\
    _iConvertSize = strlen(lpStr),\
    OEMConvNS::AnsiToOEMHelper( lpStr, (LPSTR)_alloca(_iConvertSize + 1), _iConvertSize )))

#define O2A( lpStr ) (\
  (lpStr == NULL) ? NULL: (\
    _iConvertSize = strlen(lpStr),\
    OEMConvNS::OEMToAnsiHelper( lpStr, (LPSTR)_alloca(_iConvertSize + 1), _iConvertSize )))

#define W2X( lpwStr ) (\
  (lpwStr == NULL) ? NULL: (\
    _iBuffSize = (wcslen(lpwStr) + 1)*2,\
    OEMConvNS::WToAHelper( lpwStr, (LPSTR)_alloca(_iBuffSize), _iBuffSize, _uiACP )))

#define W2XP( lpwStr, uiPage ) (\
  (lpwStr == NULL) ? NULL: (\
    _iBuffSize = (wcslen(lpwStr) + 1)*2,\
    OEMConvNS::WToAHelper( lpwStr, (LPSTR)_alloca(_iBuffSize), _iBuffSize, uiPage )))

#endif //OEM_Conv
