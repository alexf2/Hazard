#if !defined(_HALF_DIVISION_)
#define _HALF_DIVISION_

class TNoRoot {};

struct TExtraData
 {
 };

typedef double __fastcall THalfFunction( double x, TExtraData* pExtDta );


double __fastcall HalfDivizion( THalfFunction* pF, TExtraData* pD,
							double a, double b, double eps
						 );

double __fastcall Newton( THalfFunction* pF, THalfFunction* pFp1, TExtraData* pD,
					double d0, double dEps
				 );

#endif

