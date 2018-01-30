// CollLingvo.h : Declaration of the CCollLingvo

#ifndef __HYPERITERATOR_H_
#define __HYPERITERATOR_H_

#include "resource.h"       // main symbols


template<class T, class TSZ, class TSZ_HYPER, class Pfn>
class CArrayHyperIterator
 {
public:
 CArrayHyperIterator( T* pArray, TSZ tszSize, TSZ tszMaxIter, Pfn pfn, TSZ_HYPER tszhTotal = -1 ):
   m_pfn( pfn )
   {
     m_pArray = pArray; m_tszSize = tszSize;
	 m_tszMaxIter = tszMaxIter;
	 
	 if( tszhTotal == -1 )
	  {
	    m_tszhTotal = 0;
		for( int i = 1; i <= tszMaxIter; ++i )
		  m_tszhTotal += Fac64MN( tszSize, i );
	  }
	 else m_tszhTotal = tszhTotal;
	 m_tszhCurrentIter = 0;
   }

  void __fastcall Iterate();

  T* m_pArray;
  TSZ m_tszSize,  m_tszMaxIter;
  TSZ_HYPER m_tszhTotal;
  TSZ_HYPER m_tszhCurrentIter;

  auto_ptr<TSZ> m_apIdx;
  Pfn m_pfn;
 };

template<class T, class TSZ, class TSZ_HYPER, class Pfn>
void __fastcall CArrayHyperIterator<T, TSZ, TSZ_HYPER, Pfn>::Iterate()
 {
   m_tszhCurrentIter = 1;
   m_apIdx.release();

   m_apIdx = auto_ptr<TSZ>( new TSZ[m_tszMaxIter] );

   for( int iIterStop = 1; iIterStop <= m_tszMaxIter; ++iIterStop )
	{
	  m_apIdx.get()[ 0 ] = 0;
	  TSZ tszCurCycle = 0;
	  TSZ tszTo0 = m_tszSize - iIterStop + 1;

	  while( m_apIdx.get()[ 0 ] < tszTo0 )
	   {
		 if( m_apIdx.get()[ tszCurCycle ] < tszTo0 + tszCurCycle )	   
		  {
			if( tszCurCycle == iIterStop - 1 )
			 {
			   if( !m_pfn(m_pArray, m_apIdx.get(), iIterStop, m_tszhCurrentIter) )
				 break;
			   ++m_tszhCurrentIter;
			   m_apIdx.get()[ tszCurCycle ]++;
			 }
			else
			  m_apIdx.get()[ tszCurCycle + 1 ] = m_apIdx.get()[ tszCurCycle++ ] + 1;
		  }
		 else m_apIdx.get()[ --tszCurCycle ]++;	  
	   }
	}
 }

#endif 
