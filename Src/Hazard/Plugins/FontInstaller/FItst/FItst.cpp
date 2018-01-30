// FItst.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <malloc.h>
#include <windows.h>
#include "..\\fi.h"

int main(int argc, char* argv[])
 {

    FI_API::PatchExample( "e:\\hazard 2000\\examples\\def_ammonia.cfg", "StgGostsName", "e:\\hazard 2000\\examples\\ammonia.htg" );
	return 0;

    //long lRes = FI_API::InstallFonts( "g:\\tstFI\\" );
    long lRes = 0;

	if( lRes != 0 )
	 {
	   lRes = FI_API::LastFIError( NULL, 0 );
	   LPSTR lp = (LPSTR)_alloca( lRes );
	   *lp = 0;
	   lRes = FI_API::LastFIError( lp, lRes );
	   printf( "%s\r\n", lp );
	 }

    //lRes = FI_API::UnInstallFonts( "g:\\tstFI\\" );

	if( lRes != 0 )
	 {
	   lRes = FI_API::LastFIError( NULL, 0 );
	   LPSTR lp = (LPSTR)_alloca( lRes );
	   *lp = 0;
	   lRes = FI_API::LastFIError( lp, lRes );
	   printf( "%s\r\n", lp );
	 }

	return 0;
 }

