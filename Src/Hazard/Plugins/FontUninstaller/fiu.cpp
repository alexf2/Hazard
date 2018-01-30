#include "stdafx.h"
#include "fiu.h"
#define UNDEF_FI_API
#include "..\\FontInstaller\\fi.h"
#include "..\\FontInstaller\\fi_private.h"



extern HMODULE hThisModule;
extern TInitialize tiInit;

using namespace FIU_API;
using namespace std;


extern "C" LONG  FIU_API_DLL LastFIUError( LPTSTR lp, LONG lSizeBuf )
 {
   LONG sz = tiInit.SizeBuff();
   if( lp == NULL || lSizeBuf < sz ) return sz;

   _tcsncpy( lp, tiInit.GetSaveError(), lSizeBuf );
   lp[ lSizeBuf - 1 ] = 0;

   return _tcslen( lp );
 }


extern "C" LONG  FIU_API_DLL UnInstallFonts( LPCTSTR lpcPathFrom )
 {
   _TCHAR drive[ _MAX_DRIVE ];
   _TCHAR dir[ _MAX_DIR ];

    _tsplitpath( lpcPathFrom, drive, dir, NULL, NULL );

   TFntDescr * pDs = tfdFonts;
   const TFntDescr * const pDs2 = pDs + sz_tfdFonts;	
   for( ; pDs != pDs2; ++pDs )
	{
	  size_t szFBuff = _tcslen(pDs->lpcFileName)  +  _tcslen(lpcPathFrom) + _MAX_PATH + 4;
	  std::auto_ptr<_TCHAR> spPath( (LPTSTR)malloc(szFBuff * sizeof(_TCHAR)) );
	  FILEDOP_EQ( spPath.get(), NULL, FI_API::FIAPI_NoMemory );
	  
	  memset( spPath.get(), 0, szFBuff * sizeof(_TCHAR) );
	  _tmakepath( spPath.get(), drive, dir, pDs->lpcFileName, pDs->lpcFileExt );


	  HKEY hKey;
	  LONG lRes = RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts"), 0, KEY_QUERY_VALUE|KEY_SET_VALUE, &hKey );
	  if( lRes != ERROR_SUCCESS )
	    lRes = RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Fonts"), 0, KEY_QUERY_VALUE|KEY_SET_VALUE, &hKey );
	  FILEDOP_NEQ( lRes, ERROR_SUCCESS, FI_API::FIAPI_CantUpdateRegistry );

	  std::auto_ptr<_TCHAR> apTmp( (LPTSTR)malloc(szFBuff) );
	  if( apTmp.get() == NULL )
	   {
	     RegCloseKey( hKey );
		 FILEDOP_EQ( apTmp.get(), NULL, FI_API::FIAPI_NoMemory );
	   }

	  DWORD dwSzRet = szFBuff, dwTyp = REG_SZ;
	  lRes = RegQueryValueEx( hKey, pDs->lpcRegistrySlot, NULL, &dwTyp, (LPBYTE)apTmp.get(), &dwSzRet );
	  if( lRes != ERROR_SUCCESS )	    		
	   {
	     RegCloseKey( hKey );
		 continue;
	   }
	  
	  if( _tcsicmp(apTmp.get(), spPath.get()) != 0 )		
	   {
	     RegCloseKey( hKey );
		 continue;
	   }
	  

	  lRes = RegDeleteValue( hKey, pDs->lpcRegistrySlot );
	  if( lRes != ERROR_SUCCESS )
	   {
	     RegCloseKey( hKey );
		 FILEDOP_NEQ( lRes, ERROR_SUCCESS, FI_API::FIAPI_CantUpdateRegistry );
	   }
	  RegCloseKey( hKey );

	  
	  BOOL bRes = RemoveFontResource( spPath.get() );
	  if( bRes == FALSE )
	   {
	     std::basic_stringstream<TCHAR> strm;
		 strm << _T("Ошибка удаления шрифта '") << spPath.get() << _T("'");
	     tiInit.SetError( strm.str().c_str() );
	     return FI_API::FIAPI_CantFontUnregister;
	   }
	  

	  bRes = DeleteFile( spPath.get() );
	  if( bRes == FALSE )
	   {
         PFNMoveFileEx pfn = (PFNMoveFileEx)GetProcAddress( tiInit.hKernel, _T("MoveFileEx") );
		 if( pfn )
		  {
		    bRes = pfn( spPath.get(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT );
			FILEDOP_EQ( bRes, FALSE, FI_API::FIAPI_CantFonQueuetUnregister );
		  }
		 else
		  {
		    /*HKEY hKey;	     
			lRes = RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("SYSTEM\\CurrentControlSet\\Control\\Session Manager\\FileRenameOperations\\"), 0, KEY_WRITE, &hKey );
			FILEDOP_NEQ( lRes, ERROR_SUCCESS, FI_API::FIAPI_SetPendingOp );

			lRes = RegSetValueEx( hKey, spPath.get(), 0, REG_MULTI_SZ, (LPBYTE)spPath.get(), _tcslen(spPath.get()) * sizeof(_TCHAR) ); 
			if( lRes != ERROR_SUCCESS )
			 {
			   RegCloseKey( hKey );
			   FILEDOP_NEQ( lRes, ERROR_SUCCESS, FI_API::FIAPI_SetPendingOp );
			 }
			RegCloseKey( hKey );*/
		  }	                       
	   }
	}

   return FI_API::FIAPI_Ok;
 }
