// AngledFont.cpp : Implementation of CAngledFont
#include "stdafx.h"

//#include <tchar.h>

#include "AnglFont.h"
#include "AngledFont.h"


//#include "memory"
//using namespace std;



/////////////////////////////////////////////////////////////////////////////
// CAngledFont

STDMETHODIMP CAngledFont::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IAngledFont
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CAngledFont::get_Font(IDispatch **pVal)
 {
   USES_CONVERSION; 

   if( !pVal ) return E_POINTER;

   if( !m_lFontSize ) m_lFontSize = 90000;

   //CComBSTR bsN( m_lf.lfFaceName );
   FONTDESC fd = {
	  sizeof(fd), T2OLE(m_lf.lfFaceName), {m_lFontSize}, 
	  m_lf.lfWeight, m_lf.lfCharSet,m_lf.lfItalic, 
	  m_lf.lfUnderline, m_lf.lfStrikeOut 
	};

   
   return OleCreateFontIndirect( &fd, IID_IDispatch, (void**)pVal );

   /*CComPtr<IFont> spFnt;
   HRESULT hr = spFnt.CoCreateInstance( CLSID_StdFont, NULL );

   if( FAILED(hr) )
	{
	  Error( L"Font creation failed", IID_IAngledFont, hr );
	  return hr;
	}

   hr = spFnt->put_Name( m_lf.lfFaceName );
   if( SUCCEEDED(hr) )
	{
	  //spFnt->SetRatio( 13400, 1500 );
      spFnt->put_Bold( m_lf.lfWeight >= FW_BOLD );
	  spFnt->put_Italic( m_lf.lfItalic );
	  spFnt->put_Underline( m_lf.lfUnderline );
	  spFnt->put_Strikethrough( m_lf.lfStrikeOut );
	  spFnt->put_Weight( m_lf.lfWeight );
	  spFnt->put_Charset( m_lf.lfCharSet );


	  CY cy; cy.int64 = m_lf.lfHeight * 60000;

	  
	  spFnt->put_Size( cy );	  

	  hr = spFnt.QueryInterface( pVal );     
	}
      

   return hr;*/
 }

STDMETHODIMP CAngledFont::putref_Font(IDispatch *newVal)
 {
    USES_CONVERSION; 

    if( !newVal ) return E_POINTER;
	CComQIPtr<IFont> spFnt( newVal );

	if( !spFnt )
	 {
	   Error( L"Cann't QI IFont", IID_IAngledFont, E_FAIL );
	   return E_FAIL;
	 }

	CComBSTR bsName;
	HRESULT hr = spFnt->get_Name( &bsName );
	if( SUCCEEDED(hr) )
	 {
	   //CUTT memset( m_lf.lfFaceName, 0, sizeof(m_lf.lfFaceName) );       
	   ZeroMemory(  m_lf.lfFaceName, sizeof(m_lf.lfFaceName) );
	   //CUTT _tcsncpy( m_lf.lfFaceName, OLE2CT((BSTR)bsName), sizeof(m_lf.lfFaceName)/sizeof(m_lf.lfFaceName[0]) );
	   lstrcpyn( m_lf.lfFaceName, OLE2CT(Chk(bsName)), sizeof m_lf.lfFaceName );

	   int iTmp;
	   spFnt->get_Italic( &iTmp );
	   m_lf.lfItalic = iTmp;

	   spFnt->get_Underline( &iTmp );
	   m_lf.lfUnderline = iTmp;

	   spFnt->get_Strikethrough( &iTmp );
	   m_lf.lfStrikeOut = iTmp;

	   short shTmp;
	   spFnt->get_Weight( &shTmp );
	   m_lf.lfWeight = shTmp;

	   spFnt->get_Charset( &shTmp );
	   m_lf.lfCharSet = shTmp;


	   TEXTMETRICOLE tmOle;
	   spFnt->QueryTextMetrics( &tmOle );
	   m_lf.lfHeight = tmOle.tmHeight;

	   
	   CY cy;
	   spFnt->get_Size( &cy );
	   m_lFontSize = cy.int64;

	   ClearFont();
	 }

	return hr;
 }

STDMETHODIMP CAngledFont::put_Font(IDispatch *newVal)
{
	// TODO: Add your implementation code here

	return putref_Font( newVal );
}

STDMETHODIMP CAngledFont::get_Angle(long *pVal)
{
	if( !pVal ) return E_POINTER;
	*pVal = m_lAngle;

	return S_OK;
}

STDMETHODIMP CAngledFont::put_Angle(long newVal)
{	
	if( m_lAngle != newVal ) 
	 {
	   ClearFont();
	   m_lAngle = newVal;
	   Fire_Changed( DS_Angle );
	 }

	return S_OK;
}

STDMETHODIMP CAngledFont::get_hFont(long *pVal)
{
    HRESULT hr = S_OK;
	if( !pVal ) hr = E_POINTER;
	else
	 {
	   if( !m_hFont ) hr = RecreateFont();
	   if( SUCCEEDED(hr) ) *pVal = (long)m_hFont;
	 }

	return hr;
}

HRESULT CAngledFont::RecreateFont()
 {
   ClearFont();
   m_lf.lfOrientation = m_lf.lfEscapement = m_lAngle * 10;
   m_hFont = CreateFontIndirect( &m_lf );

   TEXTMETRIC tm;
   HWND hw = GetDesktopWindow();
   HDC hdc = GetWindowDC( hw );
   HFONT hfOld = (HFONT)SelectObject( hdc, m_hFont );
   GetTextMetrics( hdc, &tm );
   SelectObject( hdc, hfOld );
   ReleaseDC( hw, hdc );
   m_lFH = tm.tmHeight;

   if( !m_hFont )
	{
      Error( L"Cann't create font with current parameters", IID_IAngledFont, E_FAIL );
	  return E_FAIL;
	}
   return S_OK;
 }

STDMETHODIMP CAngledFont::InitNew()
 {
   static const LOGFONT lf = { -14, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
      OUT_TT_ONLY_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, 
	  FF_DONTCARE|VARIABLE_PITCH, _T("Arial")
	};

   m_lAngle = 0;
   m_lFontSize = 90000;
   //CUTT memcpy( &m_lf, &lf, sizeof(m_lf) );
   CopyMemory( &m_lf, &lf, sizeof m_lf );

   ClearFont();

   return S_OK;
 }
STDMETHODIMP CAngledFont::Load( IStream *pStm )
 {
   if( !pStm ) return E_POINTER;

   HRESULT hr = pStm->Read( &m_lAngle, sizeof(m_lAngle), 0 );
   if( SUCCEEDED(hr) )
	{ 
	  hr = pStm->Read( &m_lFontSize, sizeof(m_lFontSize), 0 );
	  if( SUCCEEDED(hr) )
	    hr = pStm->Read( &m_lf, sizeof(m_lf), 0 );
	}
   
   return hr;
 }
STDMETHODIMP CAngledFont::Save( IStream *pStm,  BOOL fClearDirty )
 {
   if( !pStm ) return E_POINTER;

   HRESULT hr = pStm->Write( &m_lAngle, sizeof(m_lAngle), 0 );
   if( SUCCEEDED(hr) )
	{ 
	  hr = pStm->Write( &m_lFontSize, sizeof(m_lFontSize), 0 );
	  if( SUCCEEDED(hr) )
	    hr = pStm->Write( &m_lf, sizeof(m_lf), 0 );
	}
   
   return hr;   
 }

/*void Stuff( LPBYTE pB, long sz )
 {
   LPBYTE lp2 = pB + sz - 1;
   LPBYTE lp1 = pB + sz/2 - 1;
   for( ; lp2 > pB; --lp1, lp2 -= 2 )
	{	  
	  BYTE bTmp = *lp1;
	  if( bTmp )
        *(lp2 - 1) = 1, *lp2 = bTmp;
	  else
	    *(lp2 - 1) = 2, *lp2 = 1;	  
	}
 }

void Unstuff( LPBYTE pB, long sz )
 {
   LPBYTE lp2 = pB + 1;
   LPBYTE lp1 = pB;
   pB += sz;
   for( ; lp2 < pB; ++lp1, lp2 += 2 )	
	 *lp1 = (*(lp2 - 1) == 2 ? 0:*lp2);
 }

STDMETHODIMP CAngledFont::get_Ctx(BSTR *pVal)
{
    if( !pVal ) return E_POINTER;

    long lSz = sizeof(m_lAngle) + sizeof(m_lf) + sizeof(m_lFontSize);
	lSz *= 2;
	auto_ptr<BYTE> ap( new BYTE[lSz] );

	LPBYTE lpB = ap.get();

	memcpy( lpB, &m_lAngle, sizeof(m_lAngle) );
	lpB += sizeof(m_lAngle);
	memcpy( lpB, &m_lFontSize, sizeof(m_lFontSize) );
	lpB += sizeof(m_lFontSize);
	memcpy( lpB, &m_lf, sizeof(m_lf) );
		

	Stuff( ap.get(), lSz );
	*pVal = SysAllocStringByteLen( NULL, lSz );
	memcpy( *pVal, ap.get(), lSz );	

	return S_OK;
}

STDMETHODIMP CAngledFont::put_Ctx(BSTR newVal)
{
	if( !newVal ) return E_POINTER;
	long lSz = (long)SysStringByteLen( newVal );
	
	if( lSz != 2*(sizeof(m_lAngle) + sizeof(m_lf) + sizeof(m_lFontSize)) )
	 {
       Error( L"Illegal data size", IID_IAngledFont, E_FAIL );
	   return E_FAIL;
	 }
	

	LPBYTE lpB = ((LPBYTE)newVal);

	Unstuff( lpB, lSz );

	memcpy( &m_lAngle, lpB, sizeof(m_lAngle) );
	lpB += sizeof(m_lAngle);
	memcpy( &m_lFontSize, lpB, sizeof(m_lFontSize) );
	lpB += sizeof(m_lFontSize);
	memcpy( &m_lf, lpB, sizeof(m_lf) );

	ClearFont();

	return S_OK;
}*/

STDMETHODIMP CAngledFont::get_FHeight(long *pVal)
{
    if( !pVal ) return E_POINTER;
	HRESULT hr = S_OK;
    if( !m_hFont ) hr = RecreateFont();
	if( SUCCEEDED(hr) ) *pVal = m_lFH;

	return hr;
}

STDMETHODIMP CAngledFont::CalcExtent(BSTR bs, long *pX, long *pY)
 {
    USES_CONVERSION;

	if( !bs || !pX || !pY ) return E_POINTER;

	HRESULT hr = S_OK;
	if( !m_hFont ) hr = RecreateFont();
	if( SUCCEEDED(hr) )
	 {
	   HWND hw = GetDesktopWindow();
	   HDC hdc = GetWindowDC( hw );
	   HFONT hfOld = (HFONT)SelectObject( hdc, m_hFont );
	   
	   SIZE sz;
       GetTextExtentPoint32( hdc, OLE2CT(bs), SysStringLen(bs), &sz );

	   SelectObject( hdc, hfOld );
	   ReleaseDC( hw, hdc );

	   *pX = sz.cx; *pY = sz.cy;
	 }

	return hr;
 }


STDMETHODIMP CAngledFont::get_Bold(VARIANT_BOOL *pVal)
{
	if( !pVal ) return E_POINTER;

    *pVal = (m_lf.lfWeight == FW_BOLD ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CAngledFont::put_Bold(VARIANT_BOOL newVal)
{
    long lW = (newVal == VARIANT_TRUE ? FW_BOLD:FW_NORMAL);
	if( m_lf.lfWeight != lW ) 
	 {
	   ClearFont();
	   m_lf.lfWeight = lW;
	   Fire_Changed( DS_Bold );
	 }

	return S_OK;
}
