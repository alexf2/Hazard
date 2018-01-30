#if !defined(_TL1_INTEGRAL_)
#define _TL1_INTEGRAL_

#include "halfdiv.hpp"

double Trapecion( THalfFunction* pF, TExtraData* pD,
						double a, double b, double h
					 );

double Simpson( THalfFunction* pF, TExtraData* pD,
					 double a, double b, double h
				  );

double Median( THalfFunction* pF, TExtraData* pD,
					 double a, double b, double h
				  );


double MedianTrap( THalfFunction* pF, TExtraData* pD,
					 double a, double b, double h
				  );


double Ailer( THalfFunction* pF, TExtraData* pD,
			 THalfFunction* pFp, TExtraData* pDp,
			 double a, double b, double h
		   );

double Gauss( THalfFunction* pF, TExtraData* pD,
				  double a, double b, int power
				);

struct TLezVal
 {
    TLezVal( double _x, double _a ): x(_x), a(_a)
	 {
	 };

	double x, a;
 };

typedef list<TLezVal> LST_TLezVal;
typedef LST_TLezVal::iterator IT_LST_TLezVal;

inline LST_TLezVal& operator+( LST_TLezVal& rL, TLezVal& lv )
 {
   rL.push_back( lv );
   return rL;
 }

struct TLstVal
 {
   TLstVal( short _rng, LST_TLezVal _lst ):
     rng( _rng ), lst( _lst )
	{
	}

   short rng;
   LST_TLezVal lst;
 };

typedef list<TLstVal> LST_TLstVal;
typedef LST_TLstVal::iterator IT_LST_TLstVal;

inline LST_TLstVal& operator+( LST_TLstVal& rL, TLstVal& lv )
 {
   rL.push_back( lv );
   return rL;
 }

extern  LST_TLstVal lstVal;

#endif
