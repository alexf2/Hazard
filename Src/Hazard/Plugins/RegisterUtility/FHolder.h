#if !defined(_F_HOLDER_)
#define _F_HOLDER_

#include <stdio.h>

class FHolder
 {
public:
   FHolder(): m_pf( NULL )
	{
	}
   ~FHolder()
	{
	  Close();
	}

   FHolder& operator=( FILE* pf )
	{
	  Close();
	  m_pf = pf;
	  return *this;
	}

   operator FILE*()
	{
	  return m_pf;
	}

   bool operator!() const
	{
	  return m_pf == NULL ? true:false;
	}

   void Close()
	{
	  if( m_pf )
	    fclose( m_pf ), m_pf = NULL;
	}

private:
   FILE *m_pf;   
 };

#endif
