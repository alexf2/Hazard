// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__7C4BDA55_98C6_11D4_900E_00504E02C39D__INCLUDED_)
#define AFX_STDAFX_H__7C4BDA55_98C6_11D4_900E_00504E02C39D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

// Insert your headers here
#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

#include <windows.h>

#include <atlbase.h>
extern CComModule _Module;
#include <atlcom.h>

//#include <malloc.h>
#include <sstream>
#include <memory>
//#include <list>
#include <map>
#include <new>
#include <algorithm>
using namespace std;

extern CComBSTR _bsEmpty;
inline const BSTR Chk( const CComBSTR& rBs )
 {
   return rBs.m_str == NULL ? _bsEmpty:rBs;
 }

// TODO: reference additional headers your program requires here

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__7C4BDA55_98C6_11D4_900E_00504E02C39D__INCLUDED_)
