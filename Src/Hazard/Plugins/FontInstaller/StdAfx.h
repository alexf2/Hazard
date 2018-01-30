// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__E6E9F583_B276_11D4_904F_00504E02C39D__INCLUDED_)
#define AFX_STDAFX_H__E6E9F583_B276_11D4_904F_00504E02C39D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// Insert your headers here
#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//#include <atlcom.h>

#include <windows.h>
//#include <objbase.h>

#include <stdlib.h>
#include <tchar.h>


#include <sstream>
#include <memory>

#include "resource.h"

#define FI_API_EXPORT


// TODO: reference additional headers your program requires here

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__E6E9F583_B276_11D4_904F_00504E02C39D__INCLUDED_)
