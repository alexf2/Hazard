// CollTopics.h : Declaration of the CCollTopics

#ifndef __COLLTOPICS_H_
#define __COLLTOPICS_H_

#include "resource.h"       // main symbols
#include "StColl.h"

class CCollTopics;

typedef CStColl<
  CCollTopics,
  ICollGosts,  
  IDispatchImpl<ICollTopics, &IID_ICollTopics, &LIBID_AWPLUGINLib>,
  &IID_ICollTopics,
  &CLSID_CollTopics, 
  CPersistToStorageImpl<CCollTopics, CCollGosts, &IID_ICollTopics, &CLSID_CollTopics > > TStColl_Topics;

/////////////////////////////////////////////////////////////////////////////
// CCollTopics
class ATL_NO_VTABLE CCollTopics : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollTopics, &CLSID_CollTopics>,
	public ISupportErrorInfo,
	public TStColl_Topics
{
public:
	CCollTopics()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COLLTOPICS)
DECLARE_NOT_AGGREGATABLE(CCollTopics)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCollTopics)
	COM_INTERFACE_ENTRY(ICollTopics)
	COM_INTERFACE_ENTRY(IStCollItem)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)

	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ICollTopics
public:
};

#endif //__COLLTOPICS_H_
