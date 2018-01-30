#ifndef _FI_H_
#define _FI_H_


#if !defined(FI_API_DLL)

  #if defined(FI_API_EXPORT)
    #define FI_API_DLL  __declspec(dllexport) __stdcall
  #else
    #define FI_API_DLL  __declspec(dllimport) __stdcall
  #endif

#endif

namespace FI_API {

 enum FIAPI_Errors {
    FIAPI_Ok = 0,
    FIAPI_CantFontFind = 1,
	FIAPI_CantFontLoad = 2,
	FIAPI_CantFontSize = 3,
	FIAPI_CantFontLock = 4,
	FIAPI_NoMemory = 5,
	FIAPI_ErrWrite = 6,
	FIAPI_CantFontRegister = 7,
	FIAPI_CantUpdateRegistry = 8,
	FIAPI_CantFontUnregister = 9,
	FIAPI_SetPendingOp = 10,
	FIAPI_CantFonQueuetUnregister = 11,
	FIAPI_IOError = 12
  };

#ifndef UNDEF_FI_API
extern "C" {
  LONG  FI_API_DLL InstallFonts( LPCTSTR lpcPathTo );  
  LONG  FI_API_DLL LastFIError( LPTSTR lp, LONG lSizeBuf );
  LONG  FI_API_DLL PatchExample( LPTSTR lpPathCfg, LPTSTR lpStreamName, LPTSTR lpGOSTStgPath );
//return: 0 - Ok, 
//        1 - error font-resource finding
 }
#endif

 }//FI_API


#endif
