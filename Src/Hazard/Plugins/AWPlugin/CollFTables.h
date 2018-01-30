// CollFTables.h : Declaration of the CCollFTables

#ifndef __COLLFTABLES_H_
#define __COLLFTABLES_H_

#include "resource.h"       // main symbols
#include "StColl.h"

class CCollFTables;

typedef CStColl<
  CCollFTables,
  IFactorTable,  
  IDispatchImpl<ICollFTables, &IID_ICollFTables, &LIBID_AWPLUGINLib>,
  &IID_ICollFTables,
  &CLSID_CollFTables, 
  CPersistToStreamImpl<CCollFTables, CFactorTable, &IID_ICollFTables, &CLSID_CollFTables > > TStColl_FTables;

/////////////////////////////////////////////////////////////////////////////
// CCollFTables
class ATL_NO_VTABLE CCollFTables : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollFTables, &CLSID_CollFTables>,
	public ISupportErrorInfo,	
	public TStColl_FTables
	
{
public:
	CCollFTables()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COLLFTABLES)
DECLARE_NOT_AGGREGATABLE(CCollFTables)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCollFTables)
	COM_INTERFACE_ENTRY(ICollFTables)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IStCollItem)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)

	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid) throw();

// ICollFTables
public:
   STDMETHOD(DetectRoot)( /*[out,retval]*/long* pKey ) throw();
};

#endif //__COLLFTABLES_H_
