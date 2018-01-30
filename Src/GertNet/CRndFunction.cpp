// FChange.cpp : Implementation of CFChange
#include "stdafx.h"
#include "GertNet.h"
#include "Factor.h"
#include "CRndFunction.h"

#include "PassErr.h"

#include <tchar.h>
#include <comdef.h>

#include <map>
#include <sstream>
#include <string>
#include <iomanip>
using namespace std;

#define PI 3.14159265358979L

TSincBnd arrbsFacBnd[ NUMBER_SINKS ] =
 {
  {L"C01",  {0,1,-1}, ST2 },//ST2 = 0,
  {L"C02",  {0,2,-1}, ST4 },//ST4 = 1,
  {L"H01",  {0,1,-1}, ST6 },//ST6 = 2,
  {L"H02",  {0,2,-1}, ST8 },//ST8 = 3,
  {L"M01",  {0,2,-1}, ST10 },//ST10 = 4
  {L"M02",  {0,1,-1}, ST12 },//ST12 = 5
  {L"T01",  {0,1,-1}, ST14 },//ST14 = 6
  {L"T02",  {0,2,-1}, ST16 },//ST16 = 7

  {L"M04",  {0,1,-1}, ST28 },//ST28 = 8,
  {L"C04",  {1,0,-1}, ST26 },//ST26 = 9,
  {L"H03",  {0,1, 2, 3,-1}, ST24 },//ST24 = 10
  {L"H04",  {0,1,-1}, ST22 },//ST22 = 11
  {L"H06",  {0,1,-1}, ST20 },//ST20 = 12
  {L"T03",  {1,0,-1}, ST18 },//ST18 = 13

  {L"M05",  {0,2,-1}, ST44 },//ST44 = 14
  {L"C03",  {2,0,-1}, ST42 },//ST42 = 15
  {L"H13",  {0,1,-1}, ST40 },//ST40 = 16
  {L"H14",  {0,1, 2,-1}, ST38 },//ST38 = 17
  {L"H09",  {0,1, 2,-1}, ST36 },//ST36 = 18
  {L"H05",  {0,1,-1}, ST34 },//ST34 = 19
  {L"H07",  {0,1,-1}, ST32 },//ST32 = 20
  {L"H08",  {0,1,-1}, ST30 },//ST30 = 21

  {L"H12",  {1,2,-1}, ST54 },//ST54 = 22
  {L"H14",  {0,1, 2,-1}, ST52 },//ST52 = 23
  {L"M03",  {2,0,-1}, ST50 },//ST50 = 24
  {L"T04",  {1,0,-1}, ST48 },//ST48 = 25
  {L"T05",  {1,0,-1}, ST46 },//ST46 = 26

  {L"M08",  {0,1,-1}, ST62 },//ST62 = 27
  {L"T06",  {0,1,-1}, ST60 },//ST60 = 28
  {L"M06",  {1,0,-1}, ST58 },//ST58 = 29
  {L"M07",  {1,0,-1}, ST56 }//ST56 = 30
 };

void __fastcall CRndFunction::Div2( float shI1, float shI2, bool bIndistinct ) RFTHROW0
 {
   double dX = m_dFacVal;
   float aharrY[ 2 ] = { shI1, shI2 };
   InternalDiv( 1, &dX, aharrY, bIndistinct );
 }

void __fastcall CRndFunction::Div3( float shI1, float shI2, float shI3, bool bIndistinct ) RFTHROW0
 {
       //'       (IA)          (IB)      (IC)    '}
       //' 0______________AA__________BB______1.0'}
   double darrX[ 2 ];
   if( m_dFacVal >= 0.5 )
     darrX[ 0 ] = (1.0/3.0 + (m_dFacVal - 0.5) * 4.0/3.0),
	 darrX[ 1 ] = (2.0/3.0 + (m_dFacVal - 0.5) * 2.0/3.0);
   else
	 darrX[ 0 ] = (1.0/3.0 - (0.5 - m_dFacVal) * 2.0/3.0),
	 darrX[ 1 ] = (2.0/3.0 - (0.5 - m_dFacVal) * 4.0/3.0);

   if( darrX[ 0 ] != darrX[ 1 ] )
	{
	  float aharrY[ 3 ] = { shI1, shI2, shI3 };
	  InternalDiv( 2, darrX, aharrY, bIndistinct );
	}
   else
	{
	  float aharrY[ 2 ] = { shI1, shI3 };
	  InternalDiv( 1, darrX, aharrY, bIndistinct );
	}
 }
void __fastcall CRndFunction::Div4( float shI1, float shI2, float shI3, float shI4, bool bIndistinct ) RFTHROW0
 {
   //'         (IA)              (IB)         (IC)       (ID)    '}
   //' 0___________________AA_____________BB__________CC______1.0'}
   double darrX[ 3 ];
   if( m_dFacVal >= 0.5 )
     darrX[ 0 ] = (1.0/4.0 + (m_dFacVal - 0.5) * 3.0/2.0),
	 darrX[ 1 ] = (1.0/2.0 + (m_dFacVal - 0.5)),
     darrX[ 2 ] = (3.0/4.0 + (m_dFacVal - 0.5) / 2.0);
   else
	 darrX[ 0 ] = (1.0/4.0 - (0.5 - m_dFacVal) / 2.0),
	 darrX[ 1 ] = (1.0/2.0 - (0.5 - m_dFacVal)),
     darrX[ 2 ] = (3.0/4.0 - (0.5 - m_dFacVal) * 3.0/2.0);

   if( darrX[ 0 ] == darrX[ 1 ] && darrX[ 1 ] == darrX[ 2 ] )
	{
	  float aharrY[ 2 ] = { shI1, shI4 };
	  InternalDiv( 1, darrX, aharrY, bIndistinct );
	}
   else
	{
	  float aharrY[ 4 ] = { shI1, shI2, shI3, shI4 };
	  InternalDiv( 3, darrX, aharrY, bIndistinct );
	}
 }


double __fastcall CRndFunction::operator()( double dRandomVal01 ) RFTHROW0
 {
   //double dRandomVal01 = double(rand()) / double(RAND_MAX); 

   if( dRandomVal01 < 0 ) dRandomVal01 = 0;
   else if( dRandomVal01 > 1 ) dRandomVal01 = 1;

   //Vec_RPoint::size_type sz = m_vec.size();
   Vec_RPoint::iterator it1( m_vec.begin() );
   Vec_RPoint::iterator it2( m_vec.end() );
   for( ; it1 != it2; ++it1 )
	 if( dRandomVal01 > it1->dX ) continue;
	 else
	  {
	    CRPoint& rPt2 = *it1;
        if( rPt2.dX == dRandomVal01 ) return rPt2.shY;

		CRPoint& rPt1 = *(it1 - 1);
		if( rPt1.shY == rPt2.shY ) return rPt2.shY;
		else 
		 return double((double(rPt2.shY)*(dRandomVal01 - rPt1.dX) + double(rPt1.shY)*(rPt2.dX - dRandomVal01))) / double(rPt2.dX - rPt1.dX);
	  }

	return 0;
 }

void __fastcall CRndFunction::InternalDiv( int iSz, double* pdX, float* pshY, bool bIndistinct ) RFTHROW0
 {
   Vec_RPoint::size_type sz = iSz * 2 + 2;
   if( m_vec.size() != sz )
	 m_vec.clear(), m_vec.assign( sz );

   if( iSz == 1 && (*pdX == 0 || *pdX == 1) )
	{
      if( *pdX == 0 )
	    m_vec[0].dX = 0, m_vec[0].shY = pshY[0],
		m_vec[1].dX = StdDelta(), m_vec[1].shY = pshY[1],
		m_vec[2].dX = 1, m_vec[2].shY = pshY[1],
		m_vec[3].dX = 1, m_vec[3].shY = pshY[1];
	  else
	    m_vec[0].dX = 0, m_vec[0].shY = pshY[0],
		m_vec[1].dX = 0, m_vec[1].shY = pshY[0],
		m_vec[2].dX = 1 - StdDelta(), m_vec[2].shY = pshY[0],
		m_vec[3].dX = 1, m_vec[3].shY = pshY[1];
	}
   else
	{
	  m_vec[0].dX = 0, m_vec[0].shY = pshY[0];
	  m_vec[sz - 1].dX = 1, m_vec[sz - 1].shY = pshY[ iSz ];
	  for( int i = 0, j = 1; i < iSz; ++i )
	    if( !bIndistinct )
		 {
		   m_vec[ j ].dX = pdX[ i ];
		   m_vec[ j ].shY = pshY[ i ];
		   m_vec[ ++j ].dX = pdX[ i ];
		   m_vec[ j++ ].shY = pshY[ i + 1 ];
		 }
		else
		 {
		   m_vec[ j ].dX = LeftX( m_vec[ j - 1 ].dX, pdX[ i ], bIndistinct );			
		   m_vec[ j ].shY = pshY[ i ];
		   m_vec[ ++j ].dX = RightX( pdX[ i ], (i == iSz - 1) ? 1:pdX[ i + 1 ], bIndistinct );
		   m_vec[ j++ ].shY = pshY[ i + 1 ];
		 }
	}
 }

void __fastcall CRndFunction::Dump( basic_string<WCHAR>& rS )
 {
   basic_stringstream<WCHAR> strm;
   long lFl = strm.flags() | ios::fixed;
   Vec_RPoint::size_type sz = m_vec.size();
   strm << setiosflags(lFl) << setprecision(3);
   for( int i = 0; i < sz; ++i )
	 strm << L"(" << m_vec[ i ].dX << L":" << m_vec[ i ].shY << L") ";
   strm << endl;

   rS = strm.str();

   
   /*map< bstr_t, long > mm;
   for( int j = 0; j < NUMBER_SINKS; ++j )
	{
	  pair< map< bstr_t, long >::iterator, bool > pp = 
	    mm.insert( pair<bstr_t, long>( arrbsFacBnd[j].m_bs, (long)arrbsFacBnd[j].arrshInd ) );
      if( !pp.second )
	   {
	    int yy=1;
	   }
	}
   int y =1;*/
 }

inline double fChi( double x, double q, double t )
 {
   return -1.0/PI * atan( (q - x)/t ); 
 }

inline double fChiTransform( double x, double Scale, double Placement )
 {
   return tan( PI*(x - 0.5) ) * Scale + Placement;
 }

double __fastcall CRndFunction::Pr( double x1, double x2 ) RFTHROW0
 {
   //const double q = 0.5, t = 0.2;
   switch( m_pdt )
	{
	  case PDT_Uniform:
		return x2 - x1;
	  //case PDT_Cauchy:
	  default:
		return  1.0/m_dNorma * ( fChi(x2, m_fPlacement, m_fScale) - fChi(x1, m_fPlacement, m_fScale) );
	};   
 }

void __fastcall CRndFunction::GenDistribution() RFTHROW0
 {
   //double m_dNorma, m_dCachiX0, m_dCachiX1;
   if( m_pdt == PDT_Cauchy )
	{
	  m_dNorma = 1; 
	  m_dNorma = Pr( 0, 1 );

	  m_dCachiX0 = 1.0L/PI * atan( -(m_fPlacement/m_fScale) ) + 0.5L;
	  m_dCachiX1 = 1.0L/PI * atan( (1.0L - m_fPlacement)/m_fScale ) + 0.5L;
	}
 }


double __fastcall CRndFunction::CachiRnd() RFTHROW0
 {
   double dMul = m_dCachiX1 - m_dCachiX0; 
   double dRes = fChiTransform( m_dCachiX0 + dMul * (double(rand()) / double(RAND_MAX)), m_fScale, m_fPlacement );
   if( dRes < 0 ) dRes = 0;
   else if( dRes > 1 ) dRes = 1;
   return dRes;
 }

double __fastcall CRndFunction::Density( const double x ) RFTHROW0
 {
   switch( m_pdt )
	{
	  case PDT_Uniform:
		return 1;
	  case PDT_Cauchy:	  
	   {
	     const double dTmp = (x - m_fPlacement) / m_fScale;
		 return  1.0L/m_dNorma * 1.0L / (PI * m_fScale * (1.0L + dTmp*dTmp));
	   }
	};   
 }
