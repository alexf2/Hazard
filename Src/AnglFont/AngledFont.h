// AngledFont.h : Declaration of the CAngledFont

#ifndef __ANGLEDFONT_H_
#define __ANGLEDFONT_H_

#include "resource.h"       // main symbols
#include "AnglFontCP.h"

/////////////////////////////////////////////////////////////////////////////
// CAngledFont
class ATL_NO_VTABLE CAngledFont : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAngledFont, &CLSID_AngledFont>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CAngledFont>,
	public IDispatchImpl<IAngledFont, &IID_IAngledFont, &LIBID_ANGLFONTLib>,
	public CProxy_IAngledFontEvents< CAngledFont >,
	//public IPersistPropertyBag,	
	public IPersistStreamInit,	
	public IProvideClassInfo2Impl<&CLSID_AngledFont, &DIID__IAngledFontEvents, &LIBID_ANGLFONTLib>
	
{
public:
	CAngledFont()
	{	  
	}

HRESULT FinalConstruct( )
 {
   m_hFont = NULL; m_lFH = 0; m_lFontSize = 0;
   return S_OK;
 }

void FinalRelease()
 {
   ClearFont();
 }

DECLARE_REGISTRY_RESOURCEID(IDR_ANGLEDFONT)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CAngledFont)
	COM_INTERFACE_ENTRY(IAngledFont)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
	//COM_INTERFACE_ENTRY(IPersistPropertyBag)
	//COM_INTERFACE_ENTRY2(IPersist, IPersistPropertyBag)	
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)	
	COM_INTERFACE_ENTRY2(IPersistStream, IPersistStreamInit)	
	
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)	
	
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CAngledFont)
    CONNECTION_POINT_ENTRY(DIID__IAngledFontEvents)
END_CONNECTION_POINT_MAP()


// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IAngledFont
public:
	STDMETHOD(get_Bold)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_Bold)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(CalcExtent)(/*[in]*/ BSTR bs, /*[out]*/long* pX, /*[out]*/long* pY);
	STDMETHOD(get_FHeight)(/*[out, retval]*/ long *pVal);	
	STDMETHOD(get_hFont)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_Angle)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_Angle)(/*[in]*/ long newVal);
	STDMETHOD(get_Font)(/*[out, retval]*/ IDispatch* *pVal);
	STDMETHOD(putref_Font)(/*[in]*/ IDispatch* newVal);
	STDMETHOD(put_Font)(/*[in]*/ IDispatch* newVal);

	STDMETHOD(GetClassID)(CLSID *pClassID)
	{
		ATLTRACE2(atlTraceCOM, 0, _T("IPersistPropertyBagImpl::GetClassID\n"));
		*pClassID = GetObjectCLSID();
		return S_OK;
	}

	// IPersistPropertyBag
	//
	STDMETHOD(InitNew)();	
	//STDMETHOD(Load)(LPPROPERTYBAG pPropBag, LPERRORLOG pErrorLog);	
	//STDMETHOD(Save)(LPPROPERTYBAG pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties);	

	STDMETHOD(IsDirty)()
	 {
	   return S_OK;
	 }
	STDMETHOD(GetSizeMax)( ULARGE_INTEGER*  )
	 {
	   return E_NOTIMPL;
	 }
                           
    STDMETHOD(Load)( IStream *pStm );	                 
	STDMETHOD(Save)( IStream *pStm,  BOOL fClearDirty );	 
 
	
protected:
  long m_lAngle;
  long m_lFH;
  long m_lFontSize;
  LOGFONT m_lf;

  HFONT m_hFont;

  HRESULT RecreateFont();
  void ClearFont()
   {
     if( m_hFont )
	  {
        DeleteObject( m_hFont );
		m_hFont = NULL;
	  }
   }
};

#endif //__ANGLEDFONT_H_
