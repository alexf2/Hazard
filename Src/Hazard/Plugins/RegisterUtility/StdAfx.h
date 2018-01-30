// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__840E76BA_6ADC_11D4_8FBA_00504E02C39D__INCLUDED_)
#define AFX_STDAFX_H__840E76BA_6ADC_11D4_8FBA_00504E02C39D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED


#include <stdio.h>
#include <conio.h>
#include <locale.h>
#include <atlbase.h>
#include <atlconv.h>
#include <malloc.h>

#include "OEMConv.h"
#include "FHolder.h"
#include "TCOMInit.h"


#define WAIT_AND_EXIT( iCode ) getch(); return iRet

// TODO: reference additional headers your program requires here

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__840E76BA_6ADC_11D4_8FBA_00504E02C39D__INCLUDED_)
