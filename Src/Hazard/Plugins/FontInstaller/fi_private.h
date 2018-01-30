#ifndef _FI_PRIVATE_H_
#define _FI_PRIVATE_H_

#define FILEDOP_EQ( H_Name, CmpVal, ErrCode ) \
  if( H_Name == CmpVal ) return ErrRpt( ErrCode )

#define FILEDOP_NEQ( H_Name, CmpVal, ErrCode ) \
  if( H_Name != CmpVal ) return ErrRpt( ErrCode )

#include "fi.h"

struct TInitialize
 {
   TInitialize()
	{
	  lpLastError = NULL;
	  hKernel = LoadLibrary( _T("KERNEL32.DLL") );
	}
   ~TInitialize()
	{
	  if( hKernel != NULL ) FreeLibrary( hKernel ), hKernel = NULL;
	  Clear();
	}

   LONG  SizeBuff()
	{
	  if( lpLastError == NULL ) return 0L;
	  return _tcslen( lpLastError ) + sizeof(_TCHAR);
	}

   LPCTSTR  GetSaveError()
	{
	  return lpLastError == NULL ? _T(""):lpLastError;
	}

   void  Clear()
	{
	  if( lpLastError != NULL )
	     free( lpLastError ),
		 lpLastError = NULL;
	}

   LONG  __fastcall SetError( LPCTSTR lpc )
	{
	  Clear();

	  if( lpc == NULL ) return 1;
      lpLastError = _tcsdup( lpc );	  
	  if( lpLastError == NULL ) return 0;      

	  return 1;
	}

   LPTSTR lpLastError;
   HINSTANCE hKernel;
 };

typedef WINBASEAPI BOOL (WINAPI *PFNMoveFileEx)(
  LPCTSTR lpExistingFileName,  // pointer to the name of the existing file
  LPCTSTR lpNewFileName,       // pointer to the new name for the file
  DWORD dwFlags                // flag that specifies how to move file
);


struct TFntDescr
 {
   LPCTSTR lpcFaceName, lpcFileName, lpcFileExt, lpcRegistrySlot;
   ULONG ulRcID;
   BYTE lfCharSet, lfPitchAndFamily;
   int iFontType; //TRUETYPE_FONTTYPE, RASTER_FONTTYPE, DEVICE_FONTTYPE
   bool bInstall;
 };

extern  TFntDescr tfdFonts[];
extern size_t sz_tfdFonts;
extern TInitialize tiInit;
extern LPTSTR lpCurrFile;

void __fastcall GetErrInfo( std::basic_string<_TCHAR>& rStr, DWORD dwErr );
LONG __fastcall ErrRpt( FI_API::FIAPI_Errors err );


#endif
