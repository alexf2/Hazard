#ifndef __COLLLFUNCTORS_H_
#define __COLLLFUNCTORS_H_

template<class T>
struct CApplyDirty
 {
   CApplyDirty( bool bFl )
	{
	  m_bFl = bFl;
	}
   void operator()( CComObject<T>* pObj )
	{
	  pObj->m_bRequiresSave = m_bFl;
	}
   bool m_bFl;
 };

template<class TColl, class TCollIntfPtr, class TContent, class PfnColl, class PfnCtx>
inline void ForEach_InColl( TCollIntfPtr* cpPtr, PfnColl& pfnColl, PfnCtx& pfnCtx )
 {
   CComObject<TColl>* pCF = static_cast<CComObject<TColl>*>( cpPtr );
   CComObject<TColl>::TMAP::iterator it1( pCF->m_coll.begin() );
   CComObject<TColl>::TMAP::iterator it2( pCF->m_coll.end() );
   pfnColl( pCF );
   for( ; it1 != it2; ++it1 )
	{
	  CComObject<TContent>* pFac = static_cast<CComObject<TContent>*>( it1->second );
	  pfnCtx( pFac );	  
	}
 }

#endif