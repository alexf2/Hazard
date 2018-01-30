// SafetyPrecaution.cpp : Implementation of CSafetyPrecaution
#include "stdafx.h"
#include "GertNet.h"
#include "CRndFunction.h"
#include "TnodeSFTree.h"

#include "PassErr.h"



#include <tchar.h>



#include <sstream>
#include <string>
#include <iomanip>
#include <list>
using namespace std;

#include "HALFDIV.hpp"
#include "INTEGRAL.hpp"


/*inline basic_string<wchar_t> bb( bool b )
 {
   return b ? L"true": L"false";
 }

void TstMM()
 {
   TOutNumber a1( 0.5, 1, 5 );
   TOutNumber b1( 0.5, 1, 5 );

   TOutNumber a2( 0.5, 1, 5 );
   TOutNumber b2( 0.5, 5, 10 );

   TOutNumber a3( 0.5, 1, 5 );
   TOutNumber b3( 0.5, 7, 10 );

   TOutNumber a4( 0.5, 1, 5 );
   TOutNumber b4( 0.5, 1, 10 );

   TOutNumber a5( 0.5, 5, 7 );
   TOutNumber b5( 0.5, 1, 10 );

   TOutNumber a6( 0.5, 5, 10 );
   TOutNumber b6( 0.5, 1, 10 );

   TOutNumber a7( 0.5, 1, 10 );
   TOutNumber b7( 0.5, 5, 15 );

   TOutNumber a8( 0.5, 1, 5 );
   TOutNumber b8( 0.5, 5, 10 );

   TOutNumber a9( 0.5, 1, 5 );
   TOutNumber b9( 0.5, 7, 10 );

   TOutNumber a10( 0.5, 1, 1 );
   TOutNumber b10( 0.5, 5, 10 );

   TOutNumber a11( 0.5, 1, 1 );
   TOutNumber b11( 0.5, 1, 5 );

   TOutNumber a12( 0.5, 5, 5 );
   TOutNumber b12( 0.5, 1, 10 );

   TOutNumber a13( 0.5, 5, 5 );
   TOutNumber b13( 0.5, 1, 5 );

   TOutNumber a14( 0.5, 1, 5 );
   TOutNumber b14( 0.5, 10, 10 );

   TOutNumber a15( 0.5, 1, 1 );
   TOutNumber b15( 0.5, 5, 5 );

   TOutNumber a16( 0.5, 1, 1 );
   TOutNumber b16( 0.5, 1, 1 );

   basic_stringstream<wchar_t> strm;

   strm << L"1)" << bb(a1 < b1) << L", " << bb(b1 < a1) << endl;
   strm << L"2)" << bb(a2 < b2) << L", " << bb(b2 < a2) << endl;
   strm << L"3)" << bb(a3 < b3) << L", " << bb(b3 < a3) << endl;
   strm << L"4)" << bb(a4 < b4) << L", " << bb(b4 < a4) << endl;
   strm << L"5)" << bb(a5 < b5) << L", " << bb(b5 < a5) << endl;
   strm << L"6)" << bb(a6 < b6) << L", " << bb(b6 < a6) << endl;

   strm << L"7)" << bb(a7 < b7) << L", " << bb(b7 < a7) << endl;
   strm << L"8)" << bb(a8 < b8) << L", " << bb(b8 < a8) << endl;
   strm << L"9)" << bb(a9 < b9) << L", " << bb(b9 < a9) << endl;
   strm << L"10)" << bb(a10 < b10) << L", " << bb(b10 < a10) << endl;
   strm << L"11)" << bb(a11 < b11) << L", " << bb(b11 < a11) << endl;
   strm << L"12)" << bb(a12 < b12) << L", " << bb(b12 < a12) << endl;

   strm << L"13)" << bb(a13 < b13) << L", " << bb(b13 < a13) << endl;
   strm << L"14)" << bb(a14 < b14) << L", " << bb(b14 < a14) << endl;
   strm << L"15)" << bb(a15 < b15) << L", " << bb(b15 < a15) << endl;
   strm << L"16)" << bb(a16 < b16) << L", " << bb(b16 < a16) << endl;
   

   OutputDebugStringW( strm.str().c_str() );
   int ii=1;

   static int i = 0;
   if( i ) return;
   i=1;

   TOutNumber a01( 1, 1, 10 );
   TOutNumber a02( 0.2, 5, 7 );
   SET_TOutNumber ss;
   ss + a01 + a02;
   DumpSet_TOutNumber( ss );   
   i=0;
 }*/

SET_TOutNumber& __fastcall operator+( SET_TOutNumber& rSet, TOutNumber& rON )
 {
   //DumpSet_TOutNumber( rSet );
   //TstMM();

   pair<IT_SET_TOutNumber, bool> pairIns = rSet.insert( rON );
   if( pairIns.second == false ) 
	{
	  //pairIns.first->m_lCnt++;
	  pairIns.first->m_tpP += rON.m_tpP; 

      
      if( pairIns.first->m_shNumber1 != rON.m_shNumber1 )
	   {
	     SET_TOutNumber::value_type vtTmp = *pairIns.first;

		 if( fabs(vtTmp.m_tpP - rON.m_tpP) < 0.000000001 )	  
		   vtTmp.m_shNumber1 = (vtTmp.m_shNumber1 + rON.m_shNumber1) / 2.0;	  
		 else
		  {
			double dPa, dPb, dLeft;
			if( vtTmp.m_shNumber1 < rON.m_shNumber1 )
			  dPa = vtTmp.m_tpP, dPb = rON.m_tpP, dLeft = vtTmp.m_shNumber1;
			else
			  dPb = vtTmp.m_tpP, dPa = rON.m_tpP, dLeft = rON.m_shNumber1;

			vtTmp.m_shNumber1 = dLeft + fabs(vtTmp.m_shNumber1 - rON.m_shNumber1)*( 1.0 - 1.0/(dPb/dPa + 1.0) );
		  }

		 
		 rSet.erase( pairIns.first );
		 rSet.insert( vtTmp );
	   }
	  



	  /*bool b1 = pairIns.first->IsPoint(),
		   b2 = rON.IsPoint();

	  if( b1 && b2 || 
	      rON.m_shNumber1 == pairIns.first->m_shNumber1 && rON.m_shNumber2 == pairIns.first->m_shNumber2
		)
	    pairIns.first->m_tpP += rON.m_tpP; 
	  else
	   {
         TOutNumber n1;
		 TOutNumber n1n, n2n, n3n;
		 TOutNumber n3;
		 		 

		 if( pairIns.first->m_shNumber1 < rON.m_shNumber1 )
		   n1 = *pairIns.first, n3 = rON;
		 else
		   n3 = *pairIns.first, n1 = rON;

		 rSet.erase( pairIns.first );

		 if( n3.IsPoint() )
		  {
            n1n.m_shNumber1 = n1.m_shNumber1;
			n1n.m_shNumber2 = n3.m_shNumber1;
			n1n.m_tpP = n1.CutOfP( n1n );

			n2n.m_shNumber1 = n3.m_shNumber1;
			n2n.m_shNumber2 = n3.m_shNumber1;
			n2n.m_tpP = n3.m_tpP;

			n3n.m_shNumber1 = n3.m_shNumber1;
			n3n.m_shNumber2 = n1.m_shNumber2;
			n3n.m_tpP = n1.CutOfP( n3n );

			rSet + n1n + n2n + n3n;
		  }
		 else if( n1.IsPoint() )
		  {
            int ii1=1;
		  }
         else if( n1.m_shNumber2 >= n3.m_shNumber2 )
		  {
		    if( n1.m_shNumber1 == n3.m_shNumber1 )
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber2;
			   n1n.m_tpP = n1.CutOfP( n1n ) + n3.m_tpP;

			   n2n.m_shNumber1 = n3.m_shNumber2;
			   n2n.m_shNumber2 = n1.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n );

			   rSet + n1n + n2n;
			 }
			else if( n1.m_shNumber2 == n3.m_shNumber2 )
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber1;
			   n1n.m_tpP = n1.CutOfP( n1n );

			   n2n.m_shNumber1 = n3.m_shNumber1;
			   n2n.m_shNumber2 = n1.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n ) + n3.m_tpP;

			   rSet + n1n + n2n;
			 }
			else
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber1;
			   n1n.m_tpP = n1.CutOfP( n1n );

			   n2n.m_shNumber1 = n3.m_shNumber1;
			   n2n.m_shNumber2 = n3.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n ) + n3.m_tpP;

			   n3n.m_shNumber1 = n3.m_shNumber2;
			   n3n.m_shNumber2 = n1.m_shNumber2;
			   n3n.m_tpP = n1.CutOfP( n3n );

			   rSet + n1n + n2n + n3n;
			 }		    
		  }
		 else if( n1.m_shNumber2 < n3.m_shNumber2 )
		  {
		    if( n1.m_shNumber1 == n3.m_shNumber1 )
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n1.m_shNumber2;
			   n1n.m_tpP = n3.CutOfP( n1n ) + n1.m_tpP;

			   n2n.m_shNumber1 = n1.m_shNumber2;
			   n2n.m_shNumber2 = n3.m_shNumber2;
			   n2n.m_tpP = n3.CutOfP( n2n );

			   rSet + n1n + n2n;
			 }
			else
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber1;
			   n1n.m_tpP = n1.CutOfP( n1n );

			   n2n.m_shNumber1 = n3.m_shNumber1;
			   n2n.m_shNumber2 = n1.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n ) + n3.CutOfP( n2n );

			   n3n.m_shNumber1 = n1.m_shNumber2;
			   n3n.m_shNumber2 = n3.m_shNumber2;
			   n3n.m_tpP = n3.CutOfP( n3n );

			   rSet + n1n + n2n + n3n;
			 }		    			
		  }		 
		 else
		 {		   
		 }
	   }*/
	}
   return rSet;
 }


SET_TOutNumber2& __fastcall operator+( SET_TOutNumber2& rSet, TOutNumber2& rON )
 {
   //DumpSet_TOutNumber( rSet );
   //TstMM();

   pair<IT_SET_TOutNumber2, bool> pairIns = rSet.insert( rON );
   if( pairIns.second == false ) 
	{
	  pairIns.first->m_lCnt++;

	  bool b1 = pairIns.first->IsPoint(),
		   b2 = rON.IsPoint();

	  if( b1 && b2 || 
	      rON.m_shNumber1 == pairIns.first->m_shNumber1 && rON.m_shNumber2 == pairIns.first->m_shNumber2
		)
	    pairIns.first->m_tpP += rON.m_tpP; 
	  else
	   {
         TOutNumber2 n1;
		 TOutNumber2 n1n, n2n, n3n;
		 TOutNumber2 n3;
		 		 

		 if( pairIns.first->m_shNumber1 < rON.m_shNumber1 )
		   n1 = *pairIns.first, n3 = rON;
		 else
		   n3 = *pairIns.first, n1 = rON;

		 rSet.erase( pairIns.first );

		 if( n3.IsPoint() )
		  {
            n1n.m_shNumber1 = n1.m_shNumber1;
			n1n.m_shNumber2 = n3.m_shNumber1;
			n1n.m_tpP = n1.CutOfP( n1n );

			n2n.m_shNumber1 = n3.m_shNumber1;
			n2n.m_shNumber2 = n3.m_shNumber1;
			n2n.m_tpP = n3.m_tpP;

			n3n.m_shNumber1 = n3.m_shNumber1;
			n3n.m_shNumber2 = n1.m_shNumber2;
			n3n.m_tpP = n1.CutOfP( n3n );

			rSet + n1n + n2n + n3n;
		  }
		 else if( n1.IsPoint() )
		  {
            int ii1=1;
		  }
         else if( n1.m_shNumber2 >= n3.m_shNumber2 )
		  {
		    if( n1.m_shNumber1 == n3.m_shNumber1 )
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber2;
			   n1n.m_tpP = n1.CutOfP( n1n ) + n3.m_tpP;

			   n2n.m_shNumber1 = n3.m_shNumber2;
			   n2n.m_shNumber2 = n1.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n );

			   rSet + n1n + n2n;
			 }
			else if( n1.m_shNumber2 == n3.m_shNumber2 )
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber1;
			   n1n.m_tpP = n1.CutOfP( n1n );

			   n2n.m_shNumber1 = n3.m_shNumber1;
			   n2n.m_shNumber2 = n1.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n ) + n3.m_tpP;

			   rSet + n1n + n2n;
			 }
			else
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber1;
			   n1n.m_tpP = n1.CutOfP( n1n );

			   n2n.m_shNumber1 = n3.m_shNumber1;
			   n2n.m_shNumber2 = n3.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n ) + n3.m_tpP;

			   n3n.m_shNumber1 = n3.m_shNumber2;
			   n3n.m_shNumber2 = n1.m_shNumber2;
			   n3n.m_tpP = n1.CutOfP( n3n );

			   rSet + n1n + n2n + n3n;
			 }		    
		  }
		 else if( n1.m_shNumber2 < n3.m_shNumber2 )
		  {
		    if( n1.m_shNumber1 == n3.m_shNumber1 )
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n1.m_shNumber2;
			   n1n.m_tpP = n3.CutOfP( n1n ) + n1.m_tpP;

			   n2n.m_shNumber1 = n1.m_shNumber2;
			   n2n.m_shNumber2 = n3.m_shNumber2;
			   n2n.m_tpP = n3.CutOfP( n2n );

			   rSet + n1n + n2n;
			 }
			else
			 {
			   n1n.m_shNumber1 = n1.m_shNumber1;
			   n1n.m_shNumber2 = n3.m_shNumber1;
			   n1n.m_tpP = n1.CutOfP( n1n );

			   n2n.m_shNumber1 = n3.m_shNumber1;
			   n2n.m_shNumber2 = n1.m_shNumber2;
			   n2n.m_tpP = n1.CutOfP( n2n ) + n3.CutOfP( n2n );

			   n3n.m_shNumber1 = n1.m_shNumber2;
			   n3n.m_shNumber2 = n3.m_shNumber2;
			   n3n.m_tpP = n3.CutOfP( n3n );

			   rSet + n1n + n2n + n3n;
			 }		    			
		  }		 
		 else
		 {		   
		 }
	   }
	}
   return rSet;
 }


void __fastcall DumpSet_TOutNumber( SET_TOutNumber& rS )
 {
   USES_CONVERSION; 

   basic_stringstream<wchar_t> strm;
   strm << setprecision( 7 ) << L"----------------" << endl;
   IT_SET_TOutNumber it( rS.begin() );
   int iCnt = 0;
   for( ; it != rS.end() ; ++it, ++iCnt )
	{
	  /*strm << setw(0) << L"(" << setw(7) << it->m_shNumber1 <<
	  setw(0) << L"," << setw(7) << it->m_shNumber2 << L")" << 
	  setw(12) << it->m_tpP << endl;*/

	  strm << setw(0) << L"(" << setw(7) << it->m_shNumber1 <<
	  setw(0) << L")" << 
	  setw(12) << it->m_tpP << endl;

	  if( iCnt > 10 ) 
	    OutputDebugString( OLE2CT(strm.str().c_str()) ), 
		iCnt = 0, strm.str( basic_string<wchar_t>(L""));
	}

   if( iCnt != 0 )
     OutputDebugString( OLE2CT(strm.str().c_str()) );
 }


TNodeSFTree& __fastcall TNodeSFTree::operator+( AP_SET_TOutNumber& rAPset )
 {
   m_lst.push_back( rAPset );
   return *this;
 }

bool __fastcall TNodeSFTree::Grow( SET_TOutNumber& rsetResult )
 {
   rsetResult.clear();

   if( m_pfnCancelHelper &&				
	   (m_pGN->*m_pfnCancelHelper)() 
	 ) return true;

   int iSz = m_lst.size();
   if( iSz < 1 ) return false;

   IT_SET_TOutNumber* parrIters = new IT_SET_TOutNumber[ iSz ];

   m_i64CurrentIter = 0;

   IT_L_AP_SET_TOutNumber itLst( m_lst.begin() );
   parrIters[ 0 ] = m_lst.front()->begin();
   int iIdx = 0;
   
   while( parrIters[ 0 ] != m_lst.front()->end() )
	{
      if( parrIters[iIdx] != (*itLst)->end() )
	   {
         if( iIdx == iSz - 1 )
		  {
		    TOutNumber onTmp( m_cCmp );
			onTmp.m_shNumber1 = (m_en == EN_Plus ? 0:1);
			//onTmp.m_shNumber1 = onTmp.m_shNumber2 = (m_en == EN_Plus ? 0:1);
			onTmp.m_tpP = 1;
			if( m_en == EN_Plus )
			  for( int i = iSz - 1; i >= 0; --i )
			   {
				 TOutNumber& rNumberI = *parrIters[ i ];
				 //onTmp.m_shNumber += rNumberI.m_shNumber;
				 onTmp.m_shNumber1 += rNumberI.m_shNumber1; 
				 //onTmp.m_shNumber2 += rNumberI.m_shNumber2;
				 onTmp.m_tpP *= rNumberI.m_tpP;
			   }
			 else
			  for( int i = iSz - 1; i >= 0; --i )
			   {
				 TOutNumber& rNumberI = *parrIters[ i ];
				 //onTmp.m_shNumber *= rNumberI.m_shNumber;
				 onTmp.m_shNumber1 *= rNumberI.m_shNumber1; 
				 //onTmp.m_shNumber2 *= rNumberI.m_shNumber2;
				 onTmp.m_tpP *= rNumberI.m_tpP;
			   }

			rsetResult + onTmp;

			++m_i64CurrentIter; parrIters[iIdx]++;

			if( m_pfnCancelHelper && 			    
				m_i64CurrentIter % 100 == 0 &&
				(m_pGN->*m_pfnCancelHelper)() 
			  ) 
			 {
			   delete[] parrIters;
			   return true;
			 }
			 
		  }
		 else
		   parrIters[ ++iIdx ] = (*++itLst)->begin();
	   }
	  else parrIters[ --iIdx ]++, --itLst;
	}

   delete[] parrIters;

   return false;
 }

TNodeSFTree& __fastcall TNodeSFTree::operator+( CRndFunction& rFN )
 {   
   Vec_RPoint::iterator it1( rFN.m_vec.begin() );
   Vec_RPoint::iterator it2( rFN.m_vec.end() );

   m_lst.push_back( AP_SET_TOutNumber(new SET_TOutNumber()) );
   SET_TOutNumber& rSet = *m_lst.back();
   
   double x1, x2;
   for( ; it1 != it2; ++it1 )
	 if( it1 == rFN.m_vec.begin() || (x1=(it1 - 1)->dX) == (x2=it1->dX) ) continue;
	 else
	  {	
	    //TOutNumber onTmp( it1->dX - (it1 - 1)->dX, (it1 - 1)->shY, it1->shY );
		//rSet.insert( onTmp );		

	    if( (it1 - 1)->shY == it1->shY )
		 {		   
		   //TOutNumber onTmp( x2 - x1, it1->shY, m_cCmp );
		   TOutNumber onTmp( rFN.Pr(x1, x2), it1->shY, m_cCmp );
		   rSet.insert( onTmp );				
		 }
		else
		 {
		   int lN = sNDiv;
		   //int lN = 7;
		   double dP = rFN.Pr(x1, x2) / double(lN);
		   double y1, y2;		   
		   y1 = (it1 - 1)->shY, y2 = it1->shY;

		   double dStep = (y2 - y1) / double(lN + 1);

		   for( int i = 1; i <= lN; ++i )
			{
			  TOutNumber onTmp( dP, y1 + double(i)*dStep, m_cCmp );
		      rSet.insert( onTmp );		
			}
		 }
	  }

   /*SET_TOutNumber::iterator itS1( rSet.begin() );
   double dP = 0;
   for( ; itS1 != rSet.end(); ++itS1 )
     dP += itS1->m_tpP;
   if( rSet.size() ) rSet.begin()->m_tpP += 1.0 - dP;*/

//DumpSet_TOutNumber( rSet );

   return *this;
 }

void __fastcall DivideByValue( SET_TOutNumber& rS, double dValDiv, double& rdPLower, double& rdPGreater )
 {
   IT_SET_TOutNumber it( rS.begin() );
   for( ; it != rS.end(); ++it )
	 //if( it->m_shNumber > dValDiv ) break;
	 if( it->m_shNumber1 > dValDiv ) break;
	 /*else if( it->m_shNumber2 > dValDiv ) 
	  {
	    //DumpSet_TOutNumber( rS );
	    double dIsz1 = it->m_shNumber2 - it->m_shNumber1;
	    it->m_shNumber1 = dValDiv;
		double dIsz2 = it->m_shNumber2 - dValDiv;
		it->m_tpP *= dIsz2 / dIsz1;
		break;
		//DumpSet_TOutNumber( rS );
	  }*/

   rdPLower = 0; 
   rdPGreater = 0;
   
   IT_SET_TOutNumber it1( rS.begin() );
   for( ; it1 != it; ++it1 )
	 rdPLower += it1->m_tpP;

   it1 = it;
   for( ; it1 != rS.end(); ++it1 )
	 rdPGreater += it1->m_tpP;

   //DumpSet_TOutNumber( rS );
   if( it != rS.begin() ) rS.erase( rS.begin(), it );
   //DumpSet_TOutNumber( rS );
   //int ii=1;
 }

bool __fastcall DivideByP( SET_TOutNumber& rS, double dP, double& dVal )
 {
   /*iVal = 0;
   double dDeltaP = DBL_MAX;
   IT_SET_TOutNumber it( rS.begin() );
   for( ; it != rS.end(); )
	{
      double dDelta = it->m_tpP - dP;
	  if( dDelta >= 0 )
	   {
		 if( dDelta < dDeltaP )
		  {
			dDeltaP = dDelta; iVal = it->m_shNumber;
		  }
		 it++;
	   }
	  else rS.erase( it++ );	  
	}
   return !(dDeltaP == DBL_MAX);*/

   double dSummP = 0;
   dVal = -1;
   long sz = rS.size();
   SET_TOutNumber::reverse_iterator it( rS.rbegin() );
   for( ; it != rS.rend(); ++it )	
     /*if( dSummP >= dP )
	  {
        //dVal = it->m_shNumber;
	    dVal = it->m_shNumber1;
		break;
	  }
	 else dSummP += it->m_tpP;*/
	{
	 //TOutNumber& tt = *it;

	  double d1 = dSummP;
	  dSummP += it->m_tpP;
	  if( dSummP >= dP )
	   {	     
	     /*if( !it->IsPoint() )
		  {
		  
		    //DumpSet_TOutNumber( rS );
		    //TOutNumber oo( *it );

		    it->m_shNumber1 = it->m_shNumber2 - (it->m_shNumber2 - it->m_shNumber1)*((dP - d1) / it->m_tpP);
			it->m_tpP = dP - d1;
		    dVal = it->m_shNumber1;

			SET_TOutNumber::reverse_iterator itTst( it ); 
			++itTst; 
			if( itTst != rS.rend() ) it = itTst;			

			break;
		  }*/
		 if( dSummP == dP || fabs(dSummP - dP) < fabs(d1 - dP) )
		  {
		    SET_TOutNumber::reverse_iterator itTst( it ); 
			++itTst; 
			if( itTst != rS.rend() ) it = itTst;			

		    dVal = it->m_shNumber1;
		    break;
		  }
		 else
		  {		    
		    dVal = it->m_shNumber1;
		    break;
		  }
	     

	     /*if( dSummP == dP )
		  {
		    dVal = it->m_shNumber1;
		    break;
		  }

		 dVal = it->m_shNumber1 + (it->m_shNumber2 - it->m_shNumber1)*((dP - d1) / (dP - dSummP));
		 break;*/
	   }
	}

      
	if( dVal != -1 )
	 {
	 /*SET_TOutNumber::iterator it22( rS.begin() );
	   for( ; it22 != rS.end(); ++it22 )
		 if( !it22->IsPoint() )
		  {
		    int ii=1;
		  }*/
	   //DumpSet_TOutNumber( rS );
	   
       IT_SET_TOutNumber itF = rS.find( *it );
	   if( itF != rS.end() ) rS.erase( rS.begin(), ++itF );
	   
	   //DumpSet_TOutNumber( rS );
	 }

    /*if( dVal == -1 )
	 {
	   DumpSet_TOutNumber( rS );
	 }*/

	return dVal != -1;
 }


TNodeSFTree& __fastcall TNodeSFTree::operator+( TNodeSFTree& rT )
 {
   AP_SET_TOutNumber apMy( new SET_TOutNumber() );
   AP_SET_TOutNumber apRT( new SET_TOutNumber() );
   if( Grow(*apMy) ) return *this;
   if( rT.Grow(*apRT) ) return *this;
   m_lst.clear();
   rT.m_lst.clear();

   m_lst.push_back( apMy );
   m_lst.push_back( apRT );

   m_en = EN_Plus;
   return *this;
 }
TNodeSFTree& __fastcall TNodeSFTree::operator*( TNodeSFTree& rT )
 {
   AP_SET_TOutNumber apMy( new SET_TOutNumber() );
   AP_SET_TOutNumber apRT( new SET_TOutNumber() );
   if( Grow(*apMy) ) return *this;
   if( rT.Grow(*apRT) ) return *this;
   m_lst.clear();
   rT.m_lst.clear();

   m_lst.push_back( apMy );
   m_lst.push_back( apRT );

   m_en = EN_Mul;
   return *this;
 }

/*
struct TANodesStat
 {
   double dAvgP; 
   double dSummP;
   double dMinVal; 
   double dMaxVal;
   long lSz;
   double dAvg;
   double dDx;
   bool bFlInit;
 };*/

void __fastcall GetStatistic( VEC_TOutNumber2* pV, SET_TOutNumber& rS, TANodesStat& rStat, TStatFlag sf, double* pVent )
 {
   if( sf == TSF_No ) return;

   IT_SET_TOutNumber it( rS.begin() );
   rStat.dSummP = 0; rStat.dMinVal = DBL_MAX; rStat.dMaxVal = DBL_MIN;   
   double dSumm = 0, dSumm2 = 0;
   long lCnt = 0;
   
   if( pVent )
	 for( ; it != rS.end(); ++it )	 
	   if( it->m_shNumber1 <= *pVent )
		 rStat.dSummP += it->m_tpP, lCnt++;
	   else break;
   else
	 for( ; it != rS.end(); ++it )	 	   
	   rStat.dSummP += it->m_tpP, lCnt++;
	   

   double dMul = (1.0 - rStat.dSummP) / double(lCnt);
   double dMTmp;
   it = rS.begin();
   __int64 i64Cnt = 0;
   if( !pVent )
	 for( ; it != rS.end(); ++it )	 
	  {
	    i64Cnt++;
		rStat.dMinVal = min( rStat.dMinVal, it->m_shNumber1 ),
		//rStat.dMaxVal = max( rStat.dMaxVal, it->m_shNumber2 );
		rStat.dMaxVal = max( rStat.dMaxVal, it->m_shNumber1 );
		 
		//dMTmp = it->IsPoint() ? it->m_shNumber1:(it->m_shNumber1 + it->m_shNumber2) / 2.0;
		dMTmp = it->m_shNumber1;

		dSumm += dMTmp * it->m_tpP,
		dSumm2 += dMTmp * dMTmp * it->m_tpP;
	  }
   else
	 for( ; it != rS.end(); ++it )
	   if( it->m_shNumber1 <= *pVent )
		{
		  i64Cnt++;
		  rStat.dMinVal = min( rStat.dMinVal, it->m_shNumber1 ),
		  //rStat.dMaxVal = max( rStat.dMaxVal, it->m_shNumber2 );
		  rStat.dMaxVal = max( rStat.dMaxVal, it->m_shNumber1 );

		  //dMTmp = it->IsPoint() ? it->m_shNumber1:(it->m_shNumber1 + it->m_shNumber2) / 2.0;
		  dMTmp = it->m_shNumber1;

		  dSumm += dMTmp * it->m_tpP,
		  dSumm2 += dMTmp * dMTmp * it->m_tpP;
		}
	   else break;

   if( !i64Cnt )
	 rStat.dMinVal = 0, rStat.dMaxVal = 0;

   if( rStat.dSummP != 0 )
     dSumm *= 1.0 / rStat.dSummP,
     dSumm2 *= 1.0 / rStat.dSummP;

   rStat.dAvgP = rStat.dSummP / double(lCnt); 
   /*if( i64Cnt != 0 )
     rStat.dAvg = dSumm / double(i64Cnt);
   else
	 rStat.dAvg = 0;*/
   rStat.dAvg = dSumm;
   rStat.cySz.int64 = lCnt;

   
   /*double dRes2 = 0; 
   it = rS.begin();
   for( ; it != rS.end(); ++it )
	{
	  double dSub = it->m_shNumber1 - rStat.dAvg;
	  dSub *= dSub;
	  dRes2 += dSub;
	}

   if( rS.size() > 1 )
     rStat.dDx = sqrt( dRes2 / double(rS.size() - 1) );
   else rStat.dDx = 0;*/
   rStat.dDx = sqrt(dSumm2 - dSumm*dSumm);

   rStat.bFlInit = true;

   if( pV && sf == TSF_Full ) 
	{
	  pV->assign( rS.size() );
	  copy( rS.begin(), rS.end(), pV->begin() );
	}
 }

void __fastcall TNodeATree::CalcMD( const CRndFunction& rFN, double& dM, double& dD, double& dMaxVal )
 {
   Vec_RPoint::const_iterator it1( rFN.m_vec.begin() );
   Vec_RPoint::const_iterator it2( rFN.m_vec.end() );

   dM = 0, dD = 0, dMaxVal = 0;

   for( ; it1 != it2; ++it1 )
	 if( it1 == rFN.m_vec.begin() || (it1 - 1)->dX == it1->dX ) continue;
	 else
	  {	
	    double dP = fabs( (it1 - 1)->dX - it1->dX );
	    if( (it1 - 1)->shY == it1->shY )		 
		  dM += it1->shY * dP;		 
		else
		  dM += ((it1 - 1)->shY + it1->shY) / 2.0 * dP;
		
		dMaxVal = max( dMaxVal, max((it1 - 1)->shY, it1->shY) );
	  }

	it1 = rFN.m_vec.begin();
	for( ; it1 != it2; ++it1 )
	 if( it1 == rFN.m_vec.begin() || (it1 - 1)->dX == it1->dX ) continue;
	 else
	  {	
	    double dP = fabs( (it1 - 1)->dX - it1->dX );
	    if( (it1 - 1)->shY == it1->shY )		 
		 {
		   double dDiff = it1->shY - dM;
		   dD += dDiff*dDiff * dP;
		 }
		else
		 {
		   /*double dB = (it1 - 1)->shY - it1->shY;
		   dD += dB*dB / 12.0 * dP;*/
		   double i1 = (it1 - 1)->shY, i2 = it1->shY;
		   double i21 = i2 - i1;
		   dD += dP * ( (i1 + dM)*(i1 + dM) + i21*(i1 - dM) + 1.0/3.0*i21*i21 );
		 }				
	  }
 }

struct TGammaExtraData: public TExtraData
 {
   TGammaExtraData( double m, double d )
	{
	  m_dNorma = 1; 
	  //double dd = d * d;
	  m_dA = m*m / d;
	  m_dB = m / d;
	}

   double m_dA, m_dB;
   double m_dNorma;   
 };

struct TGammaExtraDataX: public TGammaExtraData
 {   
   TGammaExtraDataX( const TGammaExtraData& rG, double dM ):
     TGammaExtraData( 1, 1 )
	{
	  m_dM = dM;

	  m_dA = rG.m_dA;
	  m_dB = rG.m_dB;
      m_dNorma = rG.m_dNorma;
	}
   
   double m_dM;
 };

static double __fastcall fGamma( double x, TGammaExtraData* pExtDta )
 {   
   return x == 0.0 ? 0.0: 
     (pExtDta->m_dNorma * pow(pExtDta->m_dB, pExtDta->m_dA) * pow(x, pExtDta->m_dA - 1.0) * exp(-pExtDta->m_dB * x) * (pExtDta->m_dA - 1.0));
 }

static double __fastcall fGammaX( double x, TGammaExtraData* pExtDta )
 {   
   return x == 0.0 ? 0.0: 
     x * (pExtDta->m_dNorma * pow(pExtDta->m_dB, pExtDta->m_dA) * pow(x, pExtDta->m_dA - 1.0) * exp(-pExtDta->m_dB * x) * (pExtDta->m_dA - 1.0));
 }

static double __fastcall fGammaXD( double x, TGammaExtraDataX* pExtDta )
 {   
   const double dDiff = x - pExtDta->m_dM;
   return x == 0.0 ? 0.0: 
     dDiff*dDiff * (pExtDta->m_dNorma * pow(pExtDta->m_dB, pExtDta->m_dA) * pow(x, pExtDta->m_dA - 1.0) * exp(-pExtDta->m_dB * x) * (pExtDta->m_dA - 1.0));
 }

void __fastcall TNodeATree::VentilCut( double dVent, TNodeATree& rPass, TNodeATree& rLeft, TNodeATree& rRight, double dNorma, double dH )
 {
   const double dMaxVal = m_dMaxVal;

   rLeft.m_dMaxVal = dVent; 
   rRight.m_dMaxVal = rPass.m_dMaxVal = m_dMaxVal;
   
   TGammaExtraData ged( m_dM, m_dD );
   double dIr = Simpson( (THalfFunction*)fGamma, &ged, dVent, dMaxVal, dH );
   ged.m_dNorma = dNorma / Simpson( (THalfFunction*)fGamma, &ged, 0, dMaxVal, dH );

   rLeft.m_dM = Simpson( (THalfFunction*)fGamma, &ged, 0, dVent, dH );
   //rRight.m_dM = Simpson( (THalfFunction*)fGamma, &ged, dVent, dMaxVal, dH );
   rRight.m_dM = dNorma - rLeft.m_dM;

   
   ged.m_dNorma = 1.0 / dIr;
   
   rPass.m_dM = Simpson( (THalfFunction*)fGammaX, &ged, dVent, dMaxVal, dH );
   TGammaExtraDataX gedx( ged, rPass.m_dM );
   rPass.m_dD = Simpson( (THalfFunction*)fGammaXD, &gedx, dVent, dMaxVal, dH );
 }

static double __fastcall NormDx( double x, double m )
 {
   return x*x*x/3.0 - 2*m*x*x/2.0 + m*m*x;
 }

void __fastcall TNodeATree::VentilCut2( double dVent, TNodeATree& rPass, TNodeATree& rLeft, TNodeATree& rRight, double dNorma, double dH )
 {
   rLeft.m_dMaxVal = dVent; 
   rRight.m_dMaxVal = rPass.m_dMaxVal = m_dMaxVal;
   
   double dMul = 1.0 / m_dMaxVal;
   double dIr = dMul * (m_dMaxVal - dVent);
   double dNorma2 = dNorma /  (dMul * m_dMaxVal);

   rLeft.m_dM = dNorma2 * dMul * dVent;   
   rRight.m_dM = dNorma - rLeft.m_dM;
   
   dNorma2 = 1.0 / dIr;
   
   dMul = 1.0 / (2.0*m_dMaxVal);
   rPass.m_dM = dNorma2 * dMul * (m_dMaxVal*m_dMaxVal - dVent*dVent);
   dMul = 1.0 / m_dMaxVal;
   rPass.m_dD = dNorma2 * dMul * (NormDx(m_dMaxVal, rPass.m_dM) - NormDx(dVent, rPass.m_dM));	
 }

