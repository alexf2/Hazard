// CompaundFiles.h : Declaration of the CCompaundFiles

#ifndef __COMPAUNDFILES_H_
#define __COMPAUNDFILES_H_

#include "resource.h"       // main symbols

#include "CompaundFilesSupport\\CompaundFilesSupport.h"

/////////////////////////////////////////////////////////////////////////////
// CCompaundFiles
class ATL_NO_VTABLE CCompaundFiles : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCompaundFiles, &CLSID_CompaundFiles>,
	public ISupportErrorInfo,
	public IDispatchImpl<ICompaundFiles, &IID_ICompaundFiles, &LIBID_GERTNETLib>
{
public:
	CCompaundFiles() throw()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COMPAUNDFILES)
DECLARE_NOT_AGGREGATABLE(CCompaundFiles)
DECLARE_CLASSFACTORY_SINGLETON(CCompaundFiles)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCompaundFiles)
	COM_INTERFACE_ENTRY(ICompaundFiles)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)  throw();

// ICompaundFiles
public:
	STDMETHOD(MakeLong)(/*[in]*/long l1, /*[in]*/long l2, /*[out,retval]*/long* lRes) throw();
	STDMETHOD(IsEmptyStg)(/*[in]*/IStorage* Stg, /*[out,retval]*/VARIANT_BOOL* bEmpty) throw();
	STDMETHOD(DestroyCF)(IStorage* Stg, /*[in]*/BSTR Name) throw();
	STDMETHOD(Revert)(/*[in]*/IStorage *Stg)  throw();
	STDMETHOD(Commit)(/*[in]*/STGC1 Flags, /*[in]*/IStorage *Stg)  throw();
	STDMETHOD(CutPath)(/*[in]*/BSTR FullPath, /*[out, optional]*/VARIANT* Dir, /*[out, optional]*/VARIANT* Name,  /*[out, optional]*/VARIANT* Ext)  throw();
	STDMETHOD(Win32ErrInfo)(/*[in]*/long ErrCode, /*[out,retval]*/BSTR* ResultStr)  throw();
	STDMETHOD(DirOfStorage)(/*[in]*/IStorage* pStg, /*[out, retval]*/ TDRFlags shFlags, SAFEARRAY** ppArrNamesStg, SAFEARRAY** ppArrNamesStrm)  throw();
	STDMETHOD(LShiftLong)(/*[in]*/long lArg, /*[in]*/long lCnt, /*[out,retval]*/long* pRes)  throw();
	STDMETHOD(RShiftLong)(/*[in]*/long lArg, /*[in]*/long lCnt, /*[out,retval]*/long* pRes)  throw();
	STDMETHOD(LoLong)(/*[in]*/long lArg, /*[out, retval]*/long* pRes)  throw();
	STDMETHOD(HiLong)(/*[in]*/long lArg, /*[out, retval]*/long* pRes)  throw();
	STDMETHOD(Sprintf)(/*[in]*/BSTR bsFormatString, long lBuffSize, /*[in]*/SAFEARRAY** parrArgs, /*[out,retval]*/BSTR* bsResult)  throw();
	STDMETHOD(OpenCF)(/*[in]*/BSTR bsName, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStorage** ppIStorage)  throw();
	STDMETHOD(CreateCF)(/*[in]*/BSTR bsName, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStorage** ppIStorage)  throw();
	STDMETHOD(CreateStorageInt)(/*[in]*/IStorage *pStg, /*[in]*/BSTR bsName, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStorage** ppIStorage)  throw();
	STDMETHOD(OpenStorageInt)(/*[in]*/IStorage *pStg, /*[in]*/BSTR bsName, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStorage** ppIStorage)  throw();
	STDMETHOD(OpenStreamInt)(/*[in]*/IStorage *pStg, /*[in]*/BSTR NameStream, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStream** ppIStrm) throw();
	STDMETHOD(CreateStreamInt)(IStorage *pStg, BSTR NameStream, EnumCompConsts lAttrOpen, IStream** ppIStrm) throw();

	STDMETHOD(PutString)( /*[in]*/IStream* Strm, /*[in]*/BSTR Str ) throw();
	STDMETHOD(GetString)( /*[in]*/IStream* Strm, /*[out,retval]*/BSTR* pStr ) throw();
};



#endif //__COMPAUNDFILES_H_
