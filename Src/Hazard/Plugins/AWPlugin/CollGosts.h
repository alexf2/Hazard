// CollGosts.h : Declaration of the CCollGosts

#ifndef __COLLGOSTS_H_
#define __COLLGOSTS_H_

#include "resource.h"       // main symbols

#include "StColl.h"

class CCollGosts;


typedef CStColl<
  CCollGosts,
  IGostTable,  
  IDispatchImpl<ICollGosts, &IID_ICollGosts, &LIBID_AWPLUGINLib>,
  &IID_ICollGosts,
  &CLSID_CollTopics, 
  CPersistToStreamImpl<CCollGosts, CGostTable, &IID_ICollGosts, &CLSID_CollTopics > > TStColl_Gosts;


class CIntCreator;

/////////////////////////////////////////////////////////////////////////////
// CCollGosts
class ATL_NO_VTABLE CCollGosts : 
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CCollGosts, &CLSID_CollGosts>,
	public ISupportErrorInfo,
	public TStColl_Gosts
	
{

friend class CIntCreator;

public:
	CCollGosts()
	{
	}

//DECLARE_REGISTRY_RESOURCEID(IDR_COLLGOSTS)
DECLARE_NOT_AGGREGATABLE(CCollGosts)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCollGosts)
	COM_INTERFACE_ENTRY(ICollGosts)
	COM_INTERFACE_ENTRY(IStCollItem)	
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)

	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ICollGosts
public:
};

#endif //__COLLGOSTS_H_
