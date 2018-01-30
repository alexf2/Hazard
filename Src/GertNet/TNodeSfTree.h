#ifndef __TNODESFTREE_H_
#define __TNODESFTREE_H_


#include <set>
#include <list>
#include <vector>
using namespace std;


extern  double dUnionThreshold;
extern  short sNDiv;


typedef enum {
      EN_Plus, EN_Mul
	} EN_NodeTypes;


class CRndFunction;

typedef double TPtyp;

#pragma pack(push, 1)
struct TOutNumber
 {
   TOutNumber( TPtyp tpP, double shNumber1, char cCmpMode )
	{
	  m_cCmpMode = cCmpMode;
	  m_tpP = tpP; 
	  
	  m_shNumber1 = shNumber1;	  
	  //m_lCnt = 1;
	}
   TOutNumber( char cCmpMode )
	{ 
	  m_cCmpMode = cCmpMode;
	  m_tpP = -1; 
	  //m_shNumber = -1;
	  m_shNumber1 = -1; //m_shNumber2 = -1;
	  //m_lCnt = 1;
	}
   TOutNumber( const TOutNumber& rD )
	{
	  this->operator=( rD );
	}
   /*TOutNumber& operator=( const TOutNumber& rD )
	{ 
	  m_tpP = rD.m_tpP;
	  //m_shNumber = rD.m_shNumber;
	  m_shNumber1 = rD.m_shNumber1;
	  m_shNumber2 = rD.m_shNumber2;
      return *this;
	}*/

   /*double CutOfP( const TOutNumber& rD )
	{
	  return m_tpP * ((rD.m_shNumber2 - rD.m_shNumber1) / (m_shNumber2 - m_shNumber1));
	}*/

   bool __fastcall operator==( const TOutNumber& rD ) const;
	
   /*bool IsPoint() const
	{
	  return m_shNumber1 == m_shNumber2;	  
	}*/
   
   TPtyp m_tpP;
   double m_shNumber1;// m_shNumber2;
   //long m_lCnt;
   char m_cCmpMode; //0 - direct, 1 - Threshold
 };

#pragma pack(pop)

inline bool operator<( const TOutNumber& rD1, const TOutNumber& rD2 )
 {      
   if( rD1.m_cCmpMode == 1 && fabs(rD2.m_shNumber1 - rD1.m_shNumber1) < dUnionThreshold ) return false;
   return rD1.m_shNumber1 < rD2.m_shNumber1;
	 

   /*bool b1 = rD1.IsPoint(), b2 = rD2.IsPoint();
   
   if( b1 && b2 )
	 return rD1.m_shNumber1 < rD2.m_shNumber1;

   if( b1 )
     return rD1.m_shNumber1 <= rD2.m_shNumber1;

   if( b2 )
     return rD1.m_shNumber2 <= rD2.m_shNumber1;
   

   return rD1.m_shNumber1 < rD2.m_shNumber1 && rD1.m_shNumber2 <= rD2.m_shNumber1;
   */
 }

inline bool TOutNumber::operator==( const TOutNumber& rD ) const
 {
   return !operator<(*this, rD) && !operator<(rD, *this);
 }

typedef vector<TOutNumber> VEC_TOutNumber;
typedef VEC_TOutNumber::iterator IT_VEC_TOutNumber;

typedef set< TOutNumber, less<TOutNumber> > SET_TOutNumber;
typedef SET_TOutNumber::iterator IT_SET_TOutNumber;

typedef auto_ptr<SET_TOutNumber> AP_SET_TOutNumber;

typedef list<AP_SET_TOutNumber> L_AP_SET_TOutNumber;
typedef L_AP_SET_TOutNumber::iterator IT_L_AP_SET_TOutNumber;


SET_TOutNumber& __fastcall operator+( SET_TOutNumber& rSet, TOutNumber& rON );


void __fastcall DumpSet_TOutNumber( SET_TOutNumber& rS );
void __fastcall DivideByValue( SET_TOutNumber& rS, double dValDiv, double& rdPLower, double& rdPGreater );
bool __fastcall DivideByP( SET_TOutNumber& rS, double dP, double& dVal );

struct TANodesStat
 {
   double dAvgP; 
   double dSummP;
   double dMinVal; 
   double dMaxVal;
   CY cySz;
   double dAvg;
   double dDx;
   bool bFlInit;
 };

struct TSampleStat
 {
   double dMinVal; 
   double dMaxVal;
   __int64 l64Sz;
   double dAvg;

   double dM2;
 };
 

class CMGertNet;
typedef bool (__fastcall CMGertNet::*PFN_CancelHelper)();

struct TNodeSFTree
 {
   
   TNodeSFTree( EN_NodeTypes en, char cCmp, PFN_CancelHelper pfn, CMGertNet* pGN, LPCTSTR lpName = "<none>" ): m_en( en )
	{
	  m_cCmp = cCmp;
	  m_sName = lpName;
	  m_i64CurrentIter = 0;

	  m_pfnCancelHelper = pfn, m_pGN = pGN;
	}

   bool __fastcall Grow( SET_TOutNumber& rsetResult );
   TNodeSFTree& __fastcall operator+( AP_SET_TOutNumber& rAPset );
   TNodeSFTree& __fastcall operator+( CRndFunction& rFN );

   TNodeSFTree& __fastcall operator+( TNodeSFTree& );
   TNodeSFTree& __fastcall operator*( TNodeSFTree& );

   TNodeSFTree& __fastcall operator=( TNodeSFTree& rT )
	{
	  m_lst.clear();
	  m_lst = rT.m_lst;
	  m_cCmp = rT.m_cCmp;
	  return *this;
	}

   EN_NodeTypes m_en;
   L_AP_SET_TOutNumber m_lst;
   __int64 m_i64CurrentIter;
   string m_sName; 
   char m_cCmp;

   PFN_CancelHelper m_pfnCancelHelper;
   CMGertNet* m_pGN;
 };

#pragma pack(push, 1)
struct TOutNumber2
 {
   TOutNumber2( TPtyp tpP, double shNumber1, double shNumber2 )
	{	  
	  m_tpP = tpP; 
	  //m_shNumber = shNumber;
	  if( shNumber1 < shNumber2 )
	    m_shNumber1 = shNumber1,
	    m_shNumber2 = shNumber2;
	  else
	    m_shNumber1 = shNumber2,
	    m_shNumber2 = shNumber1;

	  m_lCnt = 1;
	}
   TOutNumber2()
	{ 	  
	  m_tpP = -1; 
	  //m_shNumber = -1;
	  m_shNumber1 = -1; m_shNumber2 = -1;
	  m_lCnt = 1;
	}
   TOutNumber2( const TOutNumber2& rD )
	{
	  this->operator=( rD );
	}
   /*TOutNumber& operator=( const TOutNumber& rD )
	{ 
	  m_tpP = rD.m_tpP;
	  //m_shNumber = rD.m_shNumber;
	  m_shNumber1 = rD.m_shNumber1;
	  m_shNumber2 = rD.m_shNumber2;
      return *this;
	}*/

   double CutOfP( const TOutNumber2& rD )
	{
	  return m_tpP * ((rD.m_shNumber2 - rD.m_shNumber1) / (m_shNumber2 - m_shNumber1));
	}

   bool operator==( const TOutNumber2& rD ) const;

   TOutNumber2& operator=( const TOutNumber2& rD )
	{
	  m_tpP = rD.m_tpP;
	  m_shNumber1 = rD.m_shNumber1;
	  m_shNumber2 = rD.m_shNumber2;
	  m_lCnt = rD.m_lCnt;

	  return *this;
	}

   TOutNumber2& operator=( const TOutNumber& rD )
	{
	  m_tpP = rD.m_tpP;
	  m_shNumber1 = rD.m_shNumber1;
	  m_shNumber2 = rD.m_shNumber1;
	  
	  return *this;
	}
	
   bool IsPoint() const
	{
	  return m_shNumber1 == m_shNumber2;	  
	}
   
   TPtyp m_tpP;
   double m_shNumber1, m_shNumber2;
   long m_lCnt;   
 };
#pragma pack(pop)

inline bool operator<( const TOutNumber2& rD1, const TOutNumber2& rD2 )
 {      
   bool b1 = rD1.IsPoint(), b2 = rD2.IsPoint();
   //return rD1.m_shNumber < rD2.m_shNumber;
   if( b1 && b2 )
	 return rD1.m_shNumber1 < rD2.m_shNumber1;

   if( b1 )
     return rD1.m_shNumber1 <= rD2.m_shNumber1;

   if( b2 )
     return rD1.m_shNumber2 <= rD2.m_shNumber1;
   //if( b2 )
     //return rD1.m_shNumber2 < rD2.m_shNumber1;

   return rD1.m_shNumber1 < rD2.m_shNumber1 && rD1.m_shNumber2 <= rD2.m_shNumber1;
 }

inline bool TOutNumber2::operator==( const TOutNumber2& rD ) const
 {
   return !operator<(*this, rD) && !operator<(rD, *this);
 }

typedef vector<TOutNumber2> VEC_TOutNumber2;
typedef VEC_TOutNumber2::iterator IT_VEC_TOutNumber2;

typedef set< TOutNumber2, less<TOutNumber2> > SET_TOutNumber2;
typedef SET_TOutNumber2::iterator IT_SET_TOutNumber2;

typedef auto_ptr<SET_TOutNumber2> AP_SET_TOutNumber2;

typedef list<AP_SET_TOutNumber2> L_AP_SET_TOutNumber2;
typedef L_AP_SET_TOutNumber2::iterator IT_L_AP_SET_TOutNumber2;


SET_TOutNumber2& __fastcall operator+( SET_TOutNumber2& rSet, TOutNumber2& rON );

void __fastcall GetStatistic( VEC_TOutNumber2* pV, SET_TOutNumber& rS, TANodesStat&, TStatFlag, double* );

struct TNodeATree
 {   
   TNodeATree( LPCTSTR lpName = "<none>" )
	{	  
	  m_sName = lpName;	  
	  m_dM = m_dD = m_dMaxVal = 0;
	}

   TNodeATree( const CRndFunction& rFN, LPCTSTR lpName = "<none>" )
	{	  
	  m_sName = lpName;	  

	  double dM, dD, dMaxVal;
	  CalcMD( rFN, m_dM, m_dD, m_dMaxVal );	  
	}
      
   TNodeATree& __fastcall operator+( const CRndFunction& rFN )
	{
	  double dM, dD, dMaxVal;
	  CalcMD( rFN, dM, dD, dMaxVal );

	  m_dM += dM; m_dD += dD;
	  m_dMaxVal += dMaxVal;
	  return *this;
	}
   TNodeATree& __fastcall operator*( const CRndFunction& rFN )
	{
	  double dM, dD, dMaxVal;
	  CalcMD( rFN, dM, dD, dMaxVal );

	  m_dD = m_dD*dD + m_dM*m_dM*dD + dM*dM*m_dD;
	  m_dM *= dM; 
	  m_dMaxVal *= dMaxVal;
	  return *this;
	}

   TNodeATree& __fastcall operator+( const TNodeATree& rA )
    {
	  m_dM += rA.m_dM; m_dD += rA.m_dD;
	  m_dMaxVal += rA.m_dMaxVal;
	  return *this;
	}
   TNodeATree& __fastcall operator*( const TNodeATree& rA )
	{	  
	  m_dD = m_dD*rA.m_dD + m_dM*m_dM*rA.m_dD + rA.m_dM*rA.m_dM*m_dD;
	  m_dM *= rA.m_dM; 
	  m_dMaxVal *= rA.m_dMaxVal;
	  return *this;
	}

   TNodeATree& __fastcall operator=( const TNodeATree& rT )
	{	  
	  m_sName = rT.m_sName;
	  m_dM = rT.m_dM;
	  m_dD = rT.m_dD;
	  m_dMaxVal = rT.m_dMaxVal;

	  return *this;
	}


   void __fastcall VentilCut( double dVent, TNodeATree& rPass, TNodeATree& rLeft, TNodeATree& rRight, double dNorma, double dH );
   void __fastcall VentilCut2( double dVent, TNodeATree& rPass, TNodeATree& rLeft, TNodeATree& rRight, double dNorma, double dH );
   
   string m_sName;       

   double m_dM, m_dD, m_dMaxVal;
protected:   
   static void __fastcall CalcMD( const CRndFunction& rF, double& mM, double& dD, double& dMaxVal );
	
 };

#endif
