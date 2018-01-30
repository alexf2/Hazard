// CollFactors.h : Declaration of the CCollFactors

#ifndef __COLLFACTORS_H_
#define __COLLFACTORS_H_

#include "resource.h"       // main symbols
#include "StColl.h"

class CCollFactors;

typedef CStColl<
  CCollFactors,
  ICollFTables,
  IDispatchImpl<ICollFactors, &IID_ICollFactors, &LIBID_AWPLUGINLib>,
  &IID_ICollFactors,
  &CLSID_CollFactors, 
  CPersistToStorageImpl<CCollFactors, CCollFTables, &IID_ICollFactors, &CLSID_CollFactors > > TStColl_Factors;


/////////////////////////////////////////////////////////////////////////////
// CCollFactors
class ATL_NO_VTABLE CCollFactors : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollFactors, &CLSID_CollFactors>,
	public ISupportErrorInfo,
	public TStColl_Factors
	
{
public:
	CCollFactors()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COLLFACTORS)
DECLARE_NOT_AGGREGATABLE(CCollFactors)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCollFactors)
	COM_INTERFACE_ENTRY(ICollFactors)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IStCollItem)

	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)	
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ICollFactors
public:
};

#endif //__COLLFACTORS_H_
