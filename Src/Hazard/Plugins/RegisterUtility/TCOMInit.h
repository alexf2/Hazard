#if !defined(_TCOM_INIT_)
#define _TCOM_INIT_

//#include <windows.h>
#include <objbase.h>

class TCOMInit
 {
public:
   TCOMInit( DWORD co = COINIT_APARTMENTTHREADED )
	{
	  m_hr = CoInitializeEx( NULL, co );
	}
   ~TCOMInit()
	{
	  if( m_hr == S_OK || m_hr == S_FALSE ) CoUninitialize(), m_hr = E_UNEXPECTED;
	}

   bool operator!() const
	{
	  return !(m_hr == S_OK || m_hr == S_FALSE );
	}

   operator HRESULT() const
	{
	  return m_hr;
	}
private:
  HRESULT m_hr;
 };

#endif
