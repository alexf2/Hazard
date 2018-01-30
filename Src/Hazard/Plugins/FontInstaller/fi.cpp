#include "stdafx.h"
#include "fi.h"
#include "fi_private.h"

extern HMODULE hThisModule;


using namespace FI_API;
using namespace std;



extern "C" LONG  FI_API_DLL LastFIError( LPTSTR lp, LONG lSizeBuf )
 {
   LONG sz = tiInit.SizeBuff();
   if( lp == NULL || lSizeBuf < sz ) return sz;

   _tcsncpy( lp, tiInit.GetSaveError(), lSizeBuf );
   lp[ lSizeBuf - 1 ] = 0;

   return _tcslen( lp );
 }



int CALLBACK FnEnum(
  CONST LOGFONTA *lpelfe_,    // pointer to logical-font data
  CONST TEXTMETRICA *lpntme_,  // pointer to physical-font data
  DWORD FontType,             // type of font
  LPARAM lParam             // application-defined data
)
 {

   ENUMLOGFONTEX *lpelfe = (ENUMLOGFONTEX*)lpelfe_;
   NEWTEXTMETRICEX *lpntme = (NEWTEXTMETRICEX*)lpntme_;

   TFntDescr* ptfd = (TFntDescr*)lParam;
   if( ptfd->iFontType == FontType && 
	   _tcsicmp( lpelfe->elfLogFont.lfFaceName, ptfd->lpcFaceName ) == 0 )
	{
	  ptfd->bInstall = false;
	  return 0;
	}
   else return 1;	
 }

 
extern "C" LONG  FI_API_DLL InstallFonts( LPCTSTR lpcPathTo )
 {
   HWND hw = GetDesktopWindow();
   HDC hdc = GetWindowDC( hw );
   LOGFONT lf;
   
   memset( &lf, 0, sizeof lf );

   _TCHAR drive[ _MAX_DRIVE ];
   _TCHAR dir[ _MAX_DIR ];

    _tsplitpath( lpcPathTo, drive, dir, NULL, NULL );
      
   TFntDescr * pDs = tfdFonts;
   const TFntDescr * const pDs2 = pDs +  sz_tfdFonts;
   for( ; pDs != pDs2; ++pDs )
	{
	  lf.lfCharSet = pDs->lfCharSet;
	  lf.lfPitchAndFamily = pDs->lfPitchAndFamily;
	  _tcsncpy( lf.lfFaceName, pDs->lpcFaceName, (sizeof lf.lfFaceName) / (sizeof lf.lfFaceName[0]) - 1 );
	  lf.lfFaceName[ (sizeof lf.lfFaceName) - 1 ] = 0;

      EnumFontFamiliesEx( hdc, &lf, FnEnum, (LPARAM)pDs, 0 );

	  if( pDs->bInstall == false ) continue;

      HRSRC hr = FindResource( hThisModule, MAKEINTRESOURCE(pDs->ulRcID), _T("FONT_RC") );	  
	  FILEDOP_EQ( hr, NULL, FIAPI_CantFontFind );
	   
	  HGLOBAL hrc = LoadResource( hThisModule, hr );	  
	  FILEDOP_EQ( hrc, NULL, FIAPI_CantFontLoad );

	  DWORD dwFontSize = SizeofResource( hThisModule, hr );	  
	  FILEDOP_EQ( dwFontSize, 0L, FIAPI_CantFontSize );

	  LPBYTE lpFnt = (LPBYTE)LockResource( hrc );	  
	  FILEDOP_EQ( lpFnt, NULL, FIAPI_CantFontLock );


	  size_t szFBuff = _tcslen(pDs->lpcFileName)  +  _tcslen(lpcPathTo) + _MAX_PATH + 1;
	  std::auto_ptr<_TCHAR> spPath( (LPTSTR)malloc(szFBuff * sizeof(_TCHAR)) );
	  FILEDOP_EQ( spPath.get(), NULL, FIAPI_NoMemory );
	  	  
	  _tmakepath( spPath.get(), drive, dir, pDs->lpcFileName, pDs->lpcFileExt );

      HANDLE hFile = ::CreateFile( spPath.get(), GENERIC_WRITE, FILE_SHARE_READ, 
	    NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL );
	  FILEDOP_EQ( hFile, INVALID_HANDLE_VALUE, FIAPI_CantFontLock );

	  DWORD dwWritten;
	  BOOL bResW = ::WriteFile( hFile, lpFnt, dwFontSize, &dwWritten, NULL );
	  
	  if( bResW == FALSE )
	   {
	     CloseHandle( hFile );
		 FILEDOP_EQ( bResW, FALSE, FIAPI_ErrWrite );
	   }
	  CloseHandle( hFile );

	  
	  int iRes = AddFontResource( spPath.get() );
	  if( iRes == 0 )
	   {
	     if( GetLastError() == 0 )
		  {
			std::basic_stringstream<TCHAR> strm;
			strm << _T("Ошибка регистрации шрифта '") << spPath.get() << _T("'");
			tiInit.SetError( strm.str().c_str() );
			return FIAPI_CantFontRegister;
		  }
		 else
		   FILEDOP_EQ( iRes, 0, FIAPI_CantFontRegister );
	   }
	  

	  HKEY hKey = NULL;
	  LONG lRes = RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts"), 0, KEY_SET_VALUE, &hKey );
	  if( lRes != ERROR_SUCCESS )
	    lRes = RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Fonts"), 0, KEY_SET_VALUE, &hKey );

	  lRes = RegSetValueEx( hKey, pDs->lpcRegistrySlot, 0, REG_SZ, (LPBYTE)spPath.get(), _tcslen( spPath.get() ) * sizeof(_TCHAR) ); 
	  if( lRes != ERROR_SUCCESS )
	   {
	     RegCloseKey( hKey );
		 FILEDOP_NEQ( lRes, ERROR_SUCCESS, FIAPI_CantUpdateRegistry );
	   }
	  RegCloseKey( hKey );

	}

   ReleaseDC( hw, hdc );

   return FIAPI_Ok;
 }

struct TComInit
 {
   TComInit()
	{
	  CoInitialize( NULL );
	}
   ~TComInit()
	{
	  CoUninitialize();
	}
 };

extern "C" LONG  FI_API_DLL PatchExample( LPTSTR lpPathCfg, LPTSTR lpStreamName, LPTSTR lpGOSTStgPath )
 {
   USES_CONVERSION;
   TComInit tci;
   
   lpCurrFile = lpPathCfg;

   CComPtr<IStorage> spStm;
   HRESULT hr = StgOpenStorage( T2COLE(lpPathCfg), NULL, STGM_DIRECT|STGM_READWRITE|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStm );
   FILEDOP_NEQ( hr, S_OK, FIAPI_IOError );

   CComPtr<IStream> spStream;
   hr = spStm->OpenStream( T2COLE(lpStreamName), 0, STGM_DIRECT|STGM_READWRITE|STGM_SHARE_EXCLUSIVE, 0, &spStream );
   if( FAILED(hr) ) return FIAPI_Ok;   
   
   ULARGE_INTEGER ulSz; memset( &ulSz, 0, sizeof ulSz );   
   hr = spStream->SetSize( ulSz );
   FILEDOP_NEQ( hr, S_OK, FIAPI_IOError );   

   CComBSTR bs( T2COLE(lpGOSTStgPath) ); 
   hr = bs.WriteToStream( spStream ); 
   FILEDOP_NEQ( hr, S_OK, FIAPI_IOError );   

   return FIAPI_Ok;
 }

