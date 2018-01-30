#ifndef _FIU_H_
#define _FIU_H_


#if !defined(FIU_API_DLL)

  #if defined(FIU_API_EXPORT)
    #define FIU_API_DLL  __declspec(dllexport) __stdcall
  #else
    #define FIU_API_DLL  __declspec(dllimport) __stdcall
  #endif

#endif

namespace FIU_API {

 
extern "C" {  
  LONG  FIU_API_DLL UnInstallFonts( LPCTSTR lpcPathFrom );
  LONG  FIU_API_DLL LastFIUError( LPTSTR lp, LONG lSizeBuf );  
//return: 0 - Ok, 
//        1 - error font-resource finding
 }

 }//FIU_API

#endif
