#include "stdafx.h"

#include "integral.hpp"

LST_TLstVal lstVal;

int InitLst()
 {
   lstVal +

    TLstVal( 1, (LST_TLezVal() + TLezVal(0, 2)) ) +

    TLstVal( 2, (LST_TLezVal() + 
	  TLezVal(-0.577350269190, 1) + 
	  TLezVal( 0.577350269190, 1)) ) +

	TLstVal( 3, (LST_TLezVal() + 
	  TLezVal(-0.774596669241, 0.555555555556) + 
	  TLezVal( 0.0,            0.888888888889) + 
	  TLezVal( 0.774596669241, 0.555555555556)) ) +

	TLstVal( 4, (LST_TLezVal() + 
	  TLezVal(-0.861136311594, 0.347854845137) + 
	  TLezVal(-0.339981043585, 0.652145154863) + 
	  TLezVal( 0.339981043585, 0.652145154862) + 
	  TLezVal( 0.861136311594, 0.347854845137)) ) +


	TLstVal( 5, (LST_TLezVal() + 
	  TLezVal(-0.906179845939, 0.236926885056) + 
	  TLezVal(-0.538469310106, 0.478628670499) + 
	  TLezVal( 0.0,            0.568888888889) + 
	  TLezVal( 0.538469310106, 0.478628670500) + 
	  TLezVal( 0.906179845939, 0.236926885056)) ) +

	TLstVal( 6, (LST_TLezVal() + 
	  TLezVal(-0.932469514203, 0.171324492379) + 
	  TLezVal(-0.661209386466, 0.360761573048) + 
	  TLezVal(-0.238619186083, 0.467913934573) + 
	  TLezVal( 0.238619186083, 0.467913934573) + 
	  TLezVal( 0.661209386466, 0.360761573048) + 
	  TLezVal( 0.932469514203, 0.171324492379)) ) +

	TLstVal( 7, (LST_TLezVal() + 
	  TLezVal(-0.949107912343, 0.129484966169) + 
	  TLezVal(-0.741531185599, 0.279705391489) + 
	  TLezVal(-0.405845151377, 0.381830050505) + 
	  TLezVal( 0.0,            0.417959183673) + 
	  TLezVal( 0.405845151377, 0.381830050505) + 
	  TLezVal( 0.741531185599, 0.279705391489) + 
	  TLezVal( 0.949107912343, 0.129484966169)) ) +

	TLstVal( 8, (LST_TLezVal() + 
	  TLezVal(-0.960289856498, 0.101228536290) + 
	  TLezVal(-0.796666477413, 0.222381034453) + 
	  TLezVal(-0.525532409916, 0.313706645878) + 
	  TLezVal(-0.183434642496, 0.362683783378) + 
	  TLezVal( 0.183434642496, 0.362683783378) + 
	  TLezVal( 0.525532409916, 0.313706645878) + 
	  TLezVal( 0.796666477414, 0.222381034453) + 
	  TLezVal( 0.960289856498, 0.101228536290)) ) +

	TLstVal( 9, (LST_TLezVal() + 
	  TLezVal(-0.968160239508, 0.081274388362) + 
	  TLezVal(-0.836031107327, 0.180648160695) + 
	  TLezVal(-0.613371432701, 0.260610696403) + 
	  TLezVal(-0.324253423404, 0.312347077040) + 
	  TLezVal( 0.0,            0.330239355001) + 
	  TLezVal( 0.324253423404, 0.312347077040) + 
	  TLezVal( 0.613371432701, 0.260610696403) + 
	  TLezVal( 0.836031107327, 0.180648160695) + 
	  TLezVal( 0.968160239508, 0.081274388362)) ) +


	TLstVal( 12, (LST_TLezVal() + 
	  TLezVal(-0.981560634247, 0.047175336387) + 
	  TLezVal(-0.904117256370, 0.106939325995) + 
	  TLezVal(-0.769902674194, 0.160078328543) + 
	  TLezVal(-0.587317954287, 0.203167426723) + 
	  TLezVal(-0.367831498998, 0.233492536538) + 
	  TLezVal(-0.125233408511, 0.249147045813) + 
	  TLezVal( 0.125233408511, 0.249147045813) + 
	  TLezVal( 0.367831498998, 0.233492536538) + 
	  TLezVal( 0.587317954287, 0.203167426723) + 
	  TLezVal( 0.769902674194, 0.160078328543) + 
	  TLezVal( 0.904117256370, 0.106939325995) + 
	  TLezVal( 0.981560634247, 0.047175336387)) ) + 

	TLstVal( 16, (LST_TLezVal() + 
	  TLezVal(-0.989400934992, 0.027152459412) +
	  TLezVal(-0.944575023073, 0.062253523939) +
	  TLezVal(-0.865631202388, 0.095158511683) +
	  TLezVal(-0.755404408355, 0.124628971256) +
	  TLezVal(-0.617876244403, 0.149595988816) +
	  TLezVal(-0.458016777657, 0.169156519395) +
	  TLezVal(-0.281603550779, 0.182603415044) +
	  TLezVal(-0.095012509838, 0.189450610456) +
	  TLezVal( 0.095012509838, 0.189450610454) +
	  TLezVal( 0.281603550779, 0.182603415045) +
	  TLezVal( 0.458016777657, 0.169156519395) +
	  TLezVal( 0.617876244403, 0.149595988817) +
	  TLezVal( 0.755404408355, 0.124628971255) +
	  TLezVal( 0.865631202388, 0.095158511683) +
	  TLezVal( 0.944575023073, 0.062253523939) +
	  TLezVal( 0.989400934992, 0.027152459412)) ) +

	TLstVal( 20, (LST_TLezVal() + 
	  TLezVal(-0.993128599184, 0.017614007141) +
	  TLezVal(-0.963971927277, 0.040601429799) +
	  TLezVal(-0.912234428253, 0.062672048331) +
	  TLezVal(-0.839116971822, 0.083276741583) +
	  TLezVal(-0.746331906460, 0.101930119812) +
	  TLezVal(-0.636053680726, 0.118194531969) +
	  TLezVal(-0.510867001951, 0.131688638440) +
	  TLezVal(-0.373706088715, 0.142096109329) +
	  TLezVal(-0.227785851142, 0.149172986462) +
	  TLezVal(-0.076526521133, 0.152753387141) +
	  TLezVal( 0.076526521133, 0.152753387121) +
	  TLezVal( 0.227785851142, 0.149172986481) +
	  TLezVal( 0.373706088715, 0.142096109312) +
	  TLezVal( 0.510867001951, 0.131688638455) +
	  TLezVal( 0.636053680726, 0.118194531957) +
	  TLezVal( 0.746331906460, 0.101930119822) +
	  TLezVal( 0.839116971823, 0.083276741571) +
	  TLezVal( 0.912234428247, 0.062672048335) +
	  TLezVal( 0.963971927279, 0.040601429805) +
	  TLezVal( 0.993128599187, 0.017614007136)) ) +

	TLstVal( 32, (LST_TLezVal() + 
	  TLezVal(-0.997263874237, 7.018560764070e-3) +
	  TLezVal(-0.985611559613, 0.016274447249) +
	  TLezVal(-0.964762252370, 0.025391995620) +
	  TLezVal(-0.934906153002, 0.034273924652) +
	  TLezVal(-0.896321135089, 0.042835878171) +
	  TLezVal(-0.849367617730, 0.050998147837) +
	  TLezVal(-0.794483793769, 0.058683931546) +
	  TLezVal(-0.732182119336, 0.065822452010) +
	  TLezVal(-0.663044267316, 0.072345498134) +
	  TLezVal(-0.587715757276, 0.078194255601) +
	  TLezVal(-0.506899908927, 0.083311506097) +
	  TLezVal(-0.421351276131, 0.087652562148) +
	  TLezVal(-0.331868602282, 0.091173369341) +
	  TLezVal(-0.239287362204, 0.093844934606) +
	  TLezVal(-0.144471961583, 0.095638174978) +
	  TLezVal(-0.048307665688, 0.096540625604) +
	  TLezVal( 0.048307665688, 0.096539575893) +
	  TLezVal( 0.144471961583, 0.095639194665) +
	  TLezVal( 0.239287362254, 0.093843971932) +
	  TLezVal( 0.331868602274, 0.091174253613) +
	  TLezVal( 0.421351276131, 0.087651770791) +
	  TLezVal( 0.506899908932, 0.083312196782) +
	  TLezVal( 0.587715757233, 0.078193667218) +
	  TLezVal( 0.663044266684, 0.072345985802) +
	  TLezVal( 0.732182119219, 0.065822059623) +
	  TLezVal( 0.794483796762, 0.058684246025) +
	  TLezVal( 0.849367642363, 0.050997953649) +
	  TLezVal( 0.896321157186, 0.042835961844) +
	  TLezVal( 0.934906064079, 0.034273766648) +
	  TLezVal( 0.964762247462, 0.025392198686) +
	  TLezVal( 0.985611586604, 0.016274354335) +
	  TLezVal( 0.997263864227, 7.018578137582e-3)) ); 


   return 0;
 }

int iInit = InitLst();



/*const TLezVal lzVal[8][8] =
{
  { {0,2} },

  { {-0.577350,1}, {0.577350,1} },

  { {-0.774597,0.555556}, {0,0.888889}, {0.774597,0.555556} },

  { {-0.861136,0.347855}, {-0.339981,0.652145},
	 {0.339981,0.652145}, {0.861136,0.347855} },

  { {-0.906180,0.236927}, {-0.538469,0.478629},
	 {0,0.568889}, {0.538469,0.478629},
	 {0.906180,0.236927} },

  { {-0.932470,0.171325}, {-0.661209,0.360762},
	 {-0.238619,0.467914}, {0.238619,0.467914},
	 {0.661209,0.360762}, {0.932470,0.171325} },

  { {-0.949108,0.129485}, {-0.741531,0.279705},
	 {-0.405845,0.381830}, {0,0.417959},
	 {0.405845,0.381830}, {0.741531,0.279705},
	 {0.949108,0.129485} },

  { {-0.960290,0.101229}, {-0.796666,0.222381},
	 {-0.525532,0.313707}, {-0.183435,0.362684},
	 {0.183435,0.362684}, {0.525532,0.313707},
	 {0.796666,0.222381}, {0.960290,0.101229} }
 };*/

double Trapecion( THalfFunction* pF, TExtraData* pD,
						double a, double b, double h
					 )
 {
	if( a > b )
	 {
		double t = a; a = b; b = t;
	 }

	double dItegr = 0.5 * (pF( a, pD ) + pF( b, pD ));

	double dRight = b - h;
	for( a += h ; a <= dRight; a += h )
	  dItegr += pF( a, pD );

	return  h * dItegr;
 }

double Median( THalfFunction* pF, TExtraData* pD,
						double a, double b, double h
					 )
 {
	if( a > b )
	 {
		double t = a; a = b; b = t;
	 }

	double dItegr = 0;
		
	for( a += h / 2.0; a < b; a += h )
	  dItegr += pF( a, pD );

	return  h * dItegr;
 }

double MedianTrap( THalfFunction* pF, TExtraData* pD,
						double a, double b, double h
					 )
 {
	

	return  (2*Median(pF, pD, a, b, h) + Trapecion(pF, pD, a, b, h)) / 3.0;
 }



double Simpson( THalfFunction* pF, TExtraData* pD,
						double a, double b, double h
					 )
 {    

	if( a > b )
	 {
		double t = a; a = b; b = t;
	 }

	double dItegr = pF( a, pD ) + pF( b, pD );

	double dRight = b - h;
	unsigned int iOdd = 0;
	for( a += h ; a <= dRight; a += h, iOdd = ~iOdd )
	  if( !iOdd ) dItegr += 4 * pF( a, pD );
	  else dItegr += 2 * pF( a, pD );

	return  h * dItegr / 3.0;
 }

double Ailer( THalfFunction* pF, TExtraData* pD,
			 THalfFunction* pFp, TExtraData* pDp,
			 double a, double b, double h
		   )
 {
   _ASSERTE( pF != NULL );
   _ASSERTE( pFp != NULL );

   if( a > b )
	 {
		double t = a; a = b; b = t;
	 }

   double a1 = a;
   double dItegr = 0.5 * (pF(a, pD) + pF(b, pD));

   double dRight = b - h;   
   for( a += h; a <= dRight; a += h )
	 dItegr += pF( a, pD );
	
	   
   return  h * dItegr + h*h/12.0 * (pFp(a1, pDp) - pFp(b, pDp));
 }

struct CBfPow: public unary_function<TLstVal&, bool>
 {
   CBfPow( int iPow ): m_iPow( iPow )
	{
	}

   result_type operator()( argument_type arg )
	{
	  return arg.rng == m_iPow;
	}

   int m_iPow;
 };

double Gauss( THalfFunction* pF, TExtraData* pD,
				  double a, double b, int power
				)
 {
   _ASSERTE( pF != NULL );   
	/*double dIntegr = 0;
	double bma = (b - a) / 2.0;
	double bpa = (b + a) / 2.0;

	const TLezVal *pV = lzVal[ power - 1 ];
	for( int i = 0; i < power; i++, pV++ )
	  dIntegr += pV->a * pF( bma*pV->x + bpa, pD );

	return  bma * dIntegr;8*/

    CBfPow bfPow( power ); 
    IT_LST_TLstVal it = find_if( lstVal.begin(), lstVal.end(), bfPow );
    if( it == lstVal.end() ) 
	  _RPT1( _CRT_ERROR, "Число узлов %d не поддерживается", power );

	double dIntegr = 0;
	double bma = (b - a) / 2.0;
	double bpa = (b + a) / 2.0;

	IT_LST_TLezVal it2( it->lst.begin() );
	for( int i = 0; i < power; ++i, ++it2 )
	  dIntegr += it2->a * pF( bma*it2->x + bpa, pD );

	return  bma * dIntegr;
	  
 }


