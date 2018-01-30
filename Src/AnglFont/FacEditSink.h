// FacEditSink.h : Declaration of the CFacEditSink

#ifndef __FACEDITSINK_H_
#define __FACEDITSINK_H_
/*
#include "resource.h"       // main symbols
#include "GNControlsCP.h"

#include <set>
using namespace std;

class ATL_NO_VTABLE CFacEditSink;

class ATL_NO_VTABLE CSnkHelp: 
    public CComObjectRootEx<CComSingleThreadModel>,
    public IDispatchImpl<__FacEdit, &DIID___FacEdit, &LIBID_FacEditLib>
 {
public:
  CSnkHelp()
   {
   }

   void Bind( CFacEditSink* pS )
	{
	  m_pS = pS;
	}

   DECLARE_NOT_AGGREGATABLE(CSnkHelp)

   BEGIN_COM_MAP(CSnkHelp)	 
	 COM_INTERFACE_ENTRY(IDispatch)	 
   END_COM_MAP()


   STDMETHOD(ClickAssEnum)(_FacEdit* feSender);
   STDMETHOD(ClickSetupValue)(_FacEdit* feSender);
   STDMETHOD(FactorChanged)(IFactor* fac);

private:
   CFacEditSink* m_pS;
 };

/////////////////////////////////////////////////////////////////////////////
// CFacEditSink
class ATL_NO_VTABLE CFacEditSink : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFacEditSink, &CLSID_FacEditSink>,
	public IConnectionPointContainerImpl<CFacEditSink>,
	public IDispatchImpl<IFacEditSink, &IID_IFacEditSink, &LIBID_ANGLFONTLib>,
	public CProxy__FacEdit< CFacEditSink >,
	public ISupportErrorInfo,
	public IProvideClassInfo2Impl<&CLSID_AngledFont, &DIID___FacEdit, &LIBID_ANGLFONTLib>
{
public:
	CFacEditSink() 
	{
	  m_pUnkHelp = NULL;	  
	}

	HRESULT FinalConstruct()
	 {
	   m_pShlp = new CComObject<CSnkHelp>();
	   m_pShlp->Bind( this );
       HRESULT hr = m_pShlp->_InternalQueryInterface( IID_IUnknown, (void**)&m_pUnkHelp );
	   if( FAILED(hr) ) delete m_pShlp, m_pShlp = NULL;

	   return hr;
	 }
	void FinalRelease();


DECLARE_REGISTRY_RESOURCEID(IDR_FACEDITSINK)
DECLARE_NOT_AGGREGATABLE(CFacEditSink)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CFacEditSink)
	COM_INTERFACE_ENTRY(IFacEditSink)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)	
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CFacEditSink)
CONNECTION_POINT_ENTRY(DIID___FacEdit)
END_CONNECTION_POINT_MAP()


// IFacEditSink
public:
	STDMETHOD(Remove)(IDispatch* pF);
	STDMETHOD(Add)(IDispatch* pF);
	STDMETHOD(RemoveAll)(void);

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);	
	

private:
  struct TCookItem
   {
	 IDispatch* m_pDsp;
	 DWORD m_dw;	   


   };	
  struct Myless : public binary_function<TCookItem, TCookItem, bool> {	  
	bool operator()(const TCookItem& _X, const TCookItem& _Y) const
	 { return (DWORD)_X.m_pDsp < (DWORD)_Y.m_pDsp; }
  };

  typedef set<TCookItem, Myless > TS_Dsp;
  typedef TS_Dsp::iterator IT_TS_Dsp;
    

  TS_Dsp m_set;
  CComObject<CSnkHelp>* m_pShlp;
  IUnknown* m_pUnkHelp;  
};

*/

#endif //__FACEDITSINK_H_
