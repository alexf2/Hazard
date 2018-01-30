// GNRTst.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <windows.h>

#include "..\\GNRegistrar.h"

const IID IID_IRegistrarHelper = {0x56955A41,0xB40D,0x11D4,{0x90,0x5F,0x00,0x50,0x4E,0x02,0xC3,0x9D}};
const CLSID CLSID_RegistrarHelper = {0x56955A42,0xB40D,0x11D4,{0x90,0x5F,0x00,0x50,0x4E,0x02,0xC3,0x9D}};

int main(int argc, char* argv[])
{
    CoInitialize( 0 );
	IRegistrarHelper *pI;
	HRESULT hr = CoCreateInstance( CLSID_RegistrarHelper, 0, CLSCTX_ALL, IID_IRegistrarHelper, (void**)&pI );

    hr = pI->Register( L"MGertNet", L"g:\\Work\\Mag\\Alexf\\GertNet\\ReleaseMinSize\\GertNet.dll", 1, 1, L"" ) ;

	pI->Release();

	CoUninitialize();

	return 0;
}

