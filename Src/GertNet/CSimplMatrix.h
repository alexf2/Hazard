// CollFac.h : Declaration of the CCollFac

#ifndef __CSIMPLEMATRIX_H_
#define __CSIMPLEMATRIX_H_

#include "resource.h"       // main symbols
#include "PassErr.h"



template<class T> struct CSMOutParams
 {
   T m_min, m_max;
   double m_mx, m_dx;   
   void __fastcall Clear()
	{
	  m_min = m_max = 0;
	  m_mx = m_dx = 0;	  
	}
 };

template<class T, class TOut, class TMin, class TMax> class CSimpleMatrix
 {
public:
  CSimpleMatrix():m_ppT( NULL )
   {
     m_lRows = m_lClms = 0; m_bDrt = true;
   }
  ~CSimpleMatrix()
   {
     Clear();     
   }

  CSimpleMatrix( const CSimpleMatrix& rM )
   {
     m_ppT = NULL;
	 m_lRows = m_lClms = 0;

     this->operator=( rM );
   }


  long __fastcall SizeElem()
   {
     return sizeof(m_ppT[0][0]);
   }

  CSimpleMatrix& __fastcall operator=( const CSimpleMatrix& rM )
   {
     Clear();
     if( rM.m_ppT )
	  {
	    if( Init(rM.m_lRows, rM.m_lClms) )
		 {
           for( int i = 0; i < m_lRows; ++i )
	         if( rM.m_ppT[ i ] ) 
			   memcpy( m_ppT[ i ], rM.m_ppT[ i ], sizeof(m_ppT[0][0]) * m_lClms );
		 }
	  }
	 return *this;
   }

  bool __fastcall operator!()
   {
     return  m_ppT == NULL;
   }
  long __fastcall Rows()
   {
     return m_lRows;
   }
  long __fastcall Clms()
   {
     return m_lClms;
   }

  void __fastcall Clear()
   {
     if( m_ppT )
	  {
	    for( int i = 0; i < m_lRows; ++i )
	      if( m_ppT[ i ] ) delete[] m_ppT[ i ];

	    delete[] m_ppT;
	    m_ppT = NULL;
	  }
   }

  bool __fastcall Init( long lRows, long lClms )
   {
     m_lRows = lRows; m_lClms = lClms;

     m_ppT = new T*[ lRows ];
     if( !m_ppT ) return false;
	 memset( m_ppT, 0, sizeof(T*) * lRows );

	 for( int i = 0; i < lRows; ++i )
	  {
	    m_ppT[ i ] = new T[ m_lClms ];   
		if( !m_ppT[i] ) { Clear(); return false; }
	  }

	 return true;
   }

  void __fastcall Zero()
   {
     for( int i = 0; i < m_lRows; ++i )
	   memset( m_ppT[ i ], 0, SizeElem() * m_lClms );
   }

  T** __fastcall Direct()
   {
     return m_ppT;
   }

  template<class TO>
  void __fastcall OutParams( CSMOutParams<TO>* pOut, long lDivN )
   {
     for( int i = 0; i < m_lRows; ++i, ++pOut )
       CalcRow( *pOut, m_ppT[i], lDivN );
	  /*{
	    T tMin, tMax;
	    MinMax( i, tMin, tMax );
		pOut->m_min = tMin; pOut->m_max = tMax;
	    pOut->m_mx = MxRow( i );
	    pOut->m_dx = DxRow( i );
	  }*/
   }  

  template<class TO>
  void __fastcall OutParams0( CSMOutParams<TO>* pOut, long lDivN )
   {     
     for( int i = 0; i < m_lRows; ++i, ++pOut )
	  {
		double dVal = double(m_ppT[i][0]) / double(lDivN);
		if( m_ppT[i][0] < pOut->m_min ) pOut->m_min = m_ppT[i][0];
		if( m_ppT[i][0] > pOut->m_max ) pOut->m_max = m_ppT[i][0];
		pOut->m_mx += dVal;
		pOut->m_dx += dVal * dVal;
	  }
   }  

  template<class TO>
  void __fastcall OutParams0_End( CSMOutParams<TO>* pOut, long lDivK )
   {     
     double dK = lDivK;
     for( int i = 0; i < m_lRows; ++i, ++pOut )
	  {			    
		pOut->m_mx /= dK;
		pOut->m_dx /= dK;
		pOut->m_dx = sqrt( (pOut->m_dx - pOut->m_mx*pOut->m_mx) * (dK/(dK - 1.0)) );
	  }
   }  


  template<class TO>
  void __fastcall CalcRow( CSMOutParams<TO>& pOut, T* pDta, long N )
   {
     //m_lClms
     T tMin = TMax()();
	 T tMax = TMin()();
	 double dSumm = 0, dSumm2 = 0, dDivn = N;
	 for( int i = m_lClms; i > 0; --i, ++pDta )
	  {
        if( *pDta < tMin ) tMin = *pDta;
		if( *pDta > tMax ) tMax = *pDta;

		double dVal = double(*pDta) / dDivn;
		dSumm += dVal;
		dSumm2 += dVal * dVal;
	  }
	 
	 pOut.m_min = tMin; pOut.m_max = tMax;
	 pOut.m_mx = dSumm / double(m_lClms);
	 pOut.m_dx = sqrt( (dSumm2 / double(m_lClms) - pOut.m_mx*pOut.m_mx) * (double(m_lClms)/(double(m_lClms) - 1)) );
   }

  template<class TC1, class TC2, class TC3, class TC4>
  void __fastcall OutParamsDiv( CSMOutParams<T>* pOut, CSimpleMatrix<TC1, TC2, TC3, TC4>* pDiv )
   {
     for( int i = 0; i < m_lRows; ++i )
	   for( int j = 0; j < m_lClms; ++j )
		 if( pDiv->m_ppT[ i ][ j ] )
		   m_ppT[ i ][ j ] = double(m_ppT[ i ][ j ]) / double(pDiv->m_ppT[ i ][ j ]);

     for( i = 0; i < m_lRows; ++i, ++pOut )
	   MinMax( i, pOut->m_min, pOut->m_max ),	   
	   pOut->m_mx = MxRow( i ),
	   pOut->m_dx = DxRow( i );
   }

  

  T& __fastcall operator()( long lRow, long lClm )
   {
     return m_ppT[ lRow ][ lClm ];
   }
  T* __fastcall operator[]( long lRow )
   {
     return m_ppT[ lRow ];
   }

  double __fastcall MxRow( long lRow )
   {
     return Mx( m_ppT[ lRow ] );
   }

  double __fastcall DxRow( long lRow )
   {
     return Dx( m_ppT[ lRow ] );
   }

  void __fastcall MinMax( long lRow, T& tMin, T& tMax )
   {
     T* p = m_ppT[ lRow ];
	 tMin = TMax()();
	 tMax = TMin()();
	 for( int i = 0; i < m_lClms; ++i, ++p )
	  {
	    if( *p < tMin ) tMin = *p;
		if( *p > tMax ) tMax = *p;
	  }
   }
  

  void __fastcall SetDrt( bool b )
   {
     m_bDrt = b;
   }
  bool __fastcall GetDrt() { return  m_bDrt; }

//protected:
   T** m_ppT;
   long m_lRows, m_lClms;
   bool m_bDrt;

   double __fastcall Mx( T* p )
	{
	  TOut iRes = 0;
	  for( int i = 0; i < m_lClms; ++i, ++p )
	    iRes += TOut(*p);

	  return double(iRes) / double(m_lClms);	  
	}

   double __fastcall Dx( T* p )
	{
	  T* pTmp = p;
	  TOut iRes = 0;
	  for( int i = 0; i < m_lClms; ++i, ++p )
	    iRes += TOut(*p);
	  
	  double dM = double(iRes) / double(m_lClms);
	  double iRes2 = 0; p = pTmp;
	  for( i = 0; i < m_lClms; ++i, ++p )
	   {
	     double iSub = double(*p) - dM;
		 iSub *= iSub;
		 iRes2 += iSub;
	   }

	  if( m_lClms > 1 )
	    return sqrt( iRes2 / double(m_lClms - 1) );
	  else
	    return sqrt( iRes2 );
	}
 };


#endif
