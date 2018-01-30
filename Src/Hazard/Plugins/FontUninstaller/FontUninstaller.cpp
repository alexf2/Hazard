// FontUninstaller.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"

HMODULE hThisModule = NULL;

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    if( ul_reason_for_call == DLL_PROCESS_ATTACH && !hThisModule )
	 {
	   hThisModule = (HMODULE)hModule;
	   DisableThreadLibraryCalls( (HINSTANCE)hModule );
	 }

    return TRUE;
}

