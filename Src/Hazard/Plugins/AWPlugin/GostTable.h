// GostTable.h : Declaration of the CGostTable

#ifndef __GOSTTABLE_H_
#define __GOSTTABLE_H_

#include "resource.h"       // main symbols

struct CGostSlot
 {
   CGostSlot() throw()
	{
	  m_fVal = 0;
	}

   CGostSlot( int ) throw()
	{
	  m_fVal = 0;
	  m_bsValDescr = L"<Нет>";	  
	  m_bsLingvoVal = L"<Нет>";
	  m_bsComment = L"";
	}

   CGostSlot( const CGostSlot& rObj ) throw()
	{
	  m_bsValDescr.m_str = 0;	  
	  m_bsLingvoVal.m_str = 0;
	  m_bsComment.m_str = 0;
	  this->operator=( rObj );
	}
   
   CGostSlot& operator=( const CGostSlot& rObj ) throw()
	{
	  m_bsValDescr = rObj.m_bsValDescr;
	  m_fVal = rObj.m_fVal;
	  m_bsLingvoVal = rObj.m_bsLingvoVal;
	  m_bsComment = rObj.m_bsComment;	  
	  return *this;
	}

   bool operator==( const CGostSlot& rObj ) throw()
	{
	  return CmpBStrNoCase( m_bsValDescr, rObj.m_bsValDescr );
	         
	}

   HRESULT WriteToStream( LPSTREAM ) throw();
   HRESULT ReadFromStream( LPSTREAM ) throw();
   

   CComBSTR m_bsValDescr;   
   float m_fVal;
   CComBSTR m_bsLingvoVal, m_bsComment;
 };


inline bool operator<( const CGostSlot& rObj1, const CGostSlot& rObj2 ) throw()
 {
   return rObj1.m_bsValDescr.operator<( (BSTR)rObj2.m_bsValDescr );
 }

typedef vector<CGostSlot> VEC_CGostSlot;
typedef VEC_CGostSlot::iterator IT_VEC_CGostSlot;
typedef VEC_CGostSlot::const_iterator CIT_VEC_CGostSlot;

/////////////////////////////////////////////////////////////////////////////
// CGostTable
class ATL_NO_VTABLE CGostTable : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CGostTable, &CLSID_GostTable>,
	public ISupportErrorInfo,
	public IDispatchImpl<IGostTable, &IID_IGostTable, &LIBID_AWPLUGINLib>,
    public IPersistStreamInit,
	public IStCollItem
{

friend class CIntCreator;

public:

//DECLARE_REGISTRY_RESOURCEID(IDR_GOSTTABLE)
DECLARE_NOT_AGGREGATABLE(CGostTable)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CGostTable)
	COM_INTERFACE_ENTRY(IGostTable)
	COM_INTERFACE_ENTRY(IStCollItem)	
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)

	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)  throw();

// IGostTable
public:		
					
   CGostTable() throw()
	{
	}

   HRESULT FinalConstruct() throw();
   void Modify() throw()
	{
	  if( m_os == OS_None ) m_os = OS_Updated;
	}
    
//IStCollItem 
	STDMETHOD(get_Key)(/*[out, retval]*/ long *pVal) throw();	
	STDMETHOD(put_Key)(/*[out, retval]*/ long NewVal) throw();	
	STDMETHOD(get_Status)(/*[out, retval]*/ ObjStatus *pVal) throw();
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_Name)(/*[out, retval]*/ BSTR NewVal) throw();
	STDMETHOD(get_Sign)(/*[out, retval]*/ long *pVal) throw();	
	STDMETHOD(put_Sign)(/*[out, retval]*/ long NewVal) throw();	
	STDMETHOD(get_Storage)( /*[out, retval]*/ IStorage** Stg ) throw();
	STDMETHOD(put_DefaultCU)( /*[in]*/VARIANT_BOOL bMode ) throw();
    STDMETHOD(get_DefaultCU)( /*[out,retval]*/VARIANT_BOOL* pbMode ) throw();
	/*[local]*/BSTR STDMETHODCALLTYPE NameByRef( void );

//IGostTable
	STDMETHOD(get_LastErrStrZeroIdx)(/*[out, retval]*/ short *pVal);
	STDMETHOD(Check)(); 
	STDMETHOD(RemoveSlots)(/*[in]*/long Number, /*[in, optional, defaultvalue(-1)]*/long IdxStart) throw();
	STDMETHOD(InsertSlots)(/*[in]*/long Number, /*[in, optional, defaultvalue(-1)]*/long IdxStart) throw();
	STDMETHOD(get_NumberSlots)(/*[out, retval]*/ long *pVal) throw();

	STDMETHOD(get_NValue)(/*[in]*/long ZeroIndex, /*[out, retval]*/ float *pVal) throw();
	STDMETHOD(put_NValue)(/*[in]*/long ZeroIndex, /*[in]*/ float newVal) throw();
	STDMETHOD(get_NComment)(/*[in]*/long ZeroIndex, /*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_NComment)(/*[in]*/long ZeroIndex, /*[in]*/ BSTR newVal) throw();
	STDMETHOD(get_NLingvoVal)(/*[in]*/long ZeroIndex, /*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_NLingvoVal)(/*[in]*/long ZeroIndex, /*[in]*/ BSTR newVal) throw();
	STDMETHOD(get_NDescr)(/*[in]*/long ZeroIndex, /*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_NDescr)(/*[in]*/long ZeroIndex, /*[in]*/ BSTR newVal) throw();
	
	STDMETHOD(SetUniformGrad)(/*[in]*/VARIANT_BOOL Revers ) throw();
	STDMETHOD(get_ClmName)(/*[out, retval]*/ BSTR *pVal) throw();
	STDMETHOD(put_ClmName)(/*[in]*/ BSTR newVal) throw();	
	

// IPersistStreamInit
	STDMETHOD(GetClassID)(CLSID *pClassID) throw()
	 {
	   /*ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::GetClassID\n"));
	   *pClassID = GetObjectCLSID();
	   return S_OK;*/
	   return E_NOTIMPL;
	 }	
	STDMETHOD(IsDirty)() throw()
	 {
	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::IsDirty\n"));		
	   return m_os != OS_None ? S_OK : S_FALSE;
	 }

	STDMETHOD(Load)(LPSTREAM pStm) throw();	
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty) throw();
	STDMETHOD(InitNew)() throw();
	
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER FAR* /* pcbSize */) throw()
	 {
	   ATLTRACENOTIMPL(_T("IPersistStreamInit::GetSizeMax"));
	   return E_NOTIMPL;
	 }


	bool operator==( const CGostTable& rObj ) throw()
	 {
	   return CmpBStrNoCase( m_bsName, rObj.m_bsName );			  
	 }
		
protected:   
   long m_lKey;
   ObjStatus m_os;
   long m_lSign;
   bool m_bDefaultCU;
   short m_shLastErr;

   CComBSTR m_bsName;
   CComBSTR m_bsValClmName;
   VEC_CGostSlot m_vec;

   HRESULT CheckArgs( long Number, long IdxStart ) throw();
   HRESULT CheckIdx( long lIdx ) throw();
};

#endif //__GOSTTABLE_H_
