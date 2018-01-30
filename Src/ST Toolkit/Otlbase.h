// Otlbase.h
// Copyright (C) 1999 Stingray Software Inc.
// All Rights Reserved

#ifndef __OTLBASE_H__
#define __OTLBASE_H__

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


#ifndef _ATL_NO_DEBUG_CRT
	#include <crtdbg.h>
#endif

#define STINGRAY_OTLVER 0x0100 // Objective Toolkit for ATL version 1.0

#ifndef OTLASSERT
#define OTLASSERT(expr) _ASSERTE(expr)
#endif

namespace StingrayOTL
{
	inline int OtlRectWidth(const RECT& rect) { return rect.right - rect.left; }
	inline int OtlRectHeight(const RECT& rect) { return rect.bottom - rect.top; }
	inline int OtlRectWidth(LPCRECT prect) 
	{ 
		OTLASSERT(prect);
		return prect->right - prect->left; 
	}
	inline int OtlRectHeight(LPCRECT prect) 
	{ 
		OTLASSERT(prect);
		return prect->bottom - prect->top; 
	}


};	// namespace StingrayOTL

using namespace StingrayOTL;

#endif // __OTLBASE_H__
