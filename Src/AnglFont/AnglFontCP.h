#ifndef _ANGLFONTCP_H_
#define _ANGLFONTCP_H_

template <class T>
class CProxy_IAngledFontEvents : public IConnectionPointImpl<T, &DIID__IAngledFontEvents, CComDynamicUnkArray>
{
	//Warning this class may be recreated by the wizard.
public:
	HRESULT Fire_Changed(LONG id)
	{
		CComVariant varResult;
		T* pT = static_cast<T*>(this);
		int nConnectionIndex;

		CComVariant* pvars = new CComVariant[1];
		if( pvars == NULL ) return E_OUTOFMEMORY;

		pvars[0] = id;
		DISPPARAMS disp = { pvars, NULL, 1, 0 };
		int nConnections = m_vec.GetSize();
		
		for (nConnectionIndex = 0; nConnectionIndex < nConnections; nConnectionIndex++)
		{
			pT->Lock();
			CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
			pT->Unlock();
			IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
			if (pDispatch != NULL)
			{
				VariantClear(&varResult);				
				pDispatch->Invoke(0x1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
			}
		}
		delete[] pvars;
		return varResult.vt == VT_EMPTY ? S_OK:varResult.scode;	
	
	}
};
#endif