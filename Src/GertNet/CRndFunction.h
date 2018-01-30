// CollSF.h : Declaration of the CCollSF

#ifndef __RNDFUNCTION_H_
#define __RNDFUNCTION_H_

#include "resource.h"       // main symbols
#include "PassErr.h"
#include "Factor.h"
#include "idx.h"


#include <algorithm>
#include <vector>
#include <sstream>
#include <iomanip>
using namespace std;


//#define __RNDFUNCTION_H_USE_EXCEPTIONS

#ifdef __RNDFUNCTION_H_USE_EXCEPTIONS
  #define RFTHROW0  
#else  
  #define RFTHROW0 throw()
#endif

struct CRPoint
 {
   double dX; //X: 0..1
   float  shY; 
 };

typedef vector<CRPoint> Vec_RPoint;
typedef Vec_RPoint::iterator IT_Vec_RPoint;

#define DELTA_EPS  0.1
class CRndFunction
 {

public:

   CRndFunction() RFTHROW0
	{
	  m_pFactor = NULL; //reset
	  m_sValue = -1; //reset
	  m_dFacVal = 0;

      m_pdt = PDT_Uniform;
      m_fPlacement = 0.5; m_fScale = 0.19;

	  m_tr = (TrustLevel)1;

	  for( int i = 0; i < sizeof(arrshInd) / sizeof(arrshInd[0]); ++i )
	    arrshInd[ i ] = -1;
	}

   void __fastcall Reset() RFTHROW0
	{
	  m_pFactor = NULL; //reset
	  m_sValue = -1; //reset
	  m_vec.clear();
	  for( int i = 0; i < sizeof(arrshInd) / sizeof(arrshInd[0]); ++i )
	    arrshInd[ i ] = -1;
	}
   bool __fastcall IsUnbound() RFTHROW0
	{
	  return m_pFactor == NULL && m_sValue == -1;
	}

   void __fastcall GenDistribution() RFTHROW0;
   double __fastcall Density( const double x ) RFTHROW0;

   void __fastcall Div2( float shI1, float shI2, bool bIndistinct ) RFTHROW0;
   void __fastcall Div3( float shI1, float shI2, float shI3, bool bIndistinct ) RFTHROW0;
   void __fastcall Div4( float shI1, float shI2, float shI3, float shI4, bool bIndistinct ) RFTHROW0;

   double __fastcall Pr( double x1, double x2 ) RFTHROW0;

   double __fastcall operator()() RFTHROW0
	{
	  double dRandom;
	  switch( m_pdt )
	   {
	     case PDT_Uniform:
		   dRandom = double(rand()) / double(RAND_MAX);
		   break;
		 //case PDT_Cauchy:
		 default:
		   dRandom = CachiRnd();
		   break;
	   };
	  return this->operator()( dRandom );
	}
   double __fastcall operator()( double dRandomVal01 ) RFTHROW0;

   void __fastcall Dump( basic_string<WCHAR>& );

   CComObject<CFactor>* m_pFactor;
   CComBSTR m_bsShortName_Factor;

//chached values
   Vec_RPoint m_vec;//X: 0..1

   short m_sValue;
   double m_dFacVal;
   TrustLevel m_tr;   
   float arrshInd[ NUMBER_IDX ];

   PrpbDistrTyp m_pdt;
   float m_fPlacement, m_fScale;

private:
   double m_dNorma, m_dCachiX0, m_dCachiX1;
   double __fastcall CachiRnd() RFTHROW0;


   void __fastcall InternalDiv( int iSz, double* pdX, float* pshY, bool bIndistinct ) RFTHROW0;
   double __fastcall LeftX( double dPrivX, double dCurrX, bool bIndistinct ) RFTHROW0
	{	  
	  if( !bIndistinct ) return dCurrX;
	  double dDelta = StdDelta();
	  double dLX;
	  if( (dLX=dCurrX - dDelta) <= dPrivX )
	    return dPrivX + fabs(dCurrX - dPrivX) / 2.0;
	  else return dLX;
	}
   double __fastcall RightX( double dCurrX, double dNextX, bool bIndistinct ) RFTHROW0
	{	  
	  if( !bIndistinct ) return dCurrX;
	  double dDelta = StdDelta();
	  double dRX;
	  if( (dRX=dCurrX + dDelta) >= dNextX )
	    return dNextX - fabs(dNextX - dCurrX) / 2.0;
	  else return dRX;
	}
   double __fastcall StdDelta() RFTHROW0
	{
	  static const double dMul = 1.0/12.0/2.0; 
	  switch(m_tr)
	   {
	     case TL_Low:
		   return 3 * dMul;
		 case TL_Normal:
		   return 2 * dMul;
		 case TL_High:
		   return dMul;
	   };
	}
 };

#define NUMBER_SINKS 31

enum ST_Enum {
   ST2 = 0,  
   ST4 = 1,      
   ST6 = 2,      
   ST8 = 3,      
   ST10 = 4,     
   ST12 = 5,     
   ST14 = 6,     
   ST16 = 7,     

   ST28 = 8,     
   ST26 = 9,     
   ST24 = 10,     
   ST22 = 11,     
   ST20 = 12,     
   ST18 = 13,     

   ST44 = 14,      
   ST42 = 15,      
   ST40 = 16,      
   ST38 = 17,      
   ST36 = 18,      
   ST34 = 19,      
   ST32 = 20,      
   ST30 = 21,      

   ST54 = 22,      
   ST52 = 23,      
   ST50 = 24,      
   ST48 = 25,      
   ST46 = 26,      

   ST62 = 27,      
   ST60 = 28,      
   ST58 = 29,      
   ST56 = 30      
 };


#define IDX_TERM -1

struct TSincBnd
 {
   BSTR m_bs;
   short arrshInd0[ NUMBER_IDX ];
   short m_shIdx0;
 };

extern TSincBnd arrbsFacBnd[ NUMBER_SINKS ]; //factor's short names

#endif
