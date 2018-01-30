// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__F39A15D6_9890_11D4_900E_00504E02C39D__INCLUDED_)
#define AFX_STDAFX_H__F39A15D6_9890_11D4_900E_00504E02C39D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

#define MPUT_PROPERTY( OldV, NewV ) \
  if( OldV != NewV ) \
    OldV = NewV, Modify()

#define MPUT_PROPERTY_NM( OldV, NewV ) \
  if( OldV != NewV ) \
    OldV = NewV

#define MGET_PROPERTY( To, From ) \
  if( To == NULL ) return E_POINTER; \
  else *To = From


inline bool CmpBStrNoCase( const BSTR bstr1, const BSTR bstr2 )
 {
	if( bstr1 == NULL && bstr2 == NULL ) return true;
	if( bstr1 != NULL && bstr2 != NULL )
		return _wcsicmp(bstr1, bstr2) == 0;
	return false;
 }

extern CComBSTR _bsEmpty;
inline const BSTR Chk( const CComBSTR& rBs )
 {
   return rBs.m_str == NULL ? _bsEmpty:rBs;
 }

inline  const BSTR  Chk( const BSTR bs )
 {
   return bs == NULL ? _bsEmpty:bs;
 }

#pragma warning( push )
#pragma warning( disable: 4166 4503 )

#include <sstream>
#include <memory>
#include <list>
#include <set>
#include <vector>
#include <map>
#include <algorithm>
#include <exception>
#include <comutil.h>
#include <comdef.h>
using namespace std;

#include "..\..\..\GertNet\CompaundFilesSupport\CompaundFilesSupport.h"
  
#pragma warning( pop )

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__F39A15D6_9890_11D4_900E_00504E02C39D__INCLUDED)
