// TstIntegral.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "integral.hpp"


#include <ole2.h>

using namespace std;
#include <fstream>
#include <sstream>
#include <string>
#include <limits>
#include <iomanip>

double f1( double x, TExtraData* pExtDta )
 {
   double dX = x - 2;
   return dX*dX - 2;
 }
double f1p( double x, TExtraData* pExtDta )
 {   
   return 2.0*x - 4.0;
 }

double f1i( double a, double b )
 {   
#define Mi(x) (pow(x, 3)/3.0 - 2.0*x*x + 2.0*x)
   return Mi(b) - Mi(a);
 }

double f2( double x, TExtraData* pExtDta )
 {
   double dX = 0.5*x*x*x + 2*x*x + 2;
   return dX*dX;
 }

double fG( double x, TExtraData* pExtDta )
 {
   const double a = 2.259532910365, b = 0.153919135583;
   const double dNorma = 1.0 / 1.434827811054;   
   return x == 0 ? 0: 
     (dNorma * pow(b, a) * pow(x, a - 1) * exp(-b * x) * (a - 1));
 }

double fGp( double x, TExtraData* pExtDta )
 { 
   const double a = 2.259532910365, b = 0.153919135583;
   const double dNorma = 1.0 / 1.434827811054;   
   return x == 0 ? 0:
	 (dNorma * pow(b, a) * (a - 1) * exp(-b * x) * pow(x, a) * ((a - 1)*pow(x, -2) - b*pow(x, -1)));
 }

double GaussIter( THalfFunction* pF, TExtraData* pD, double a, double b, int power, float fPercent )
 {
   double i = 0.0;
   double dStep = (b - a) / 100.0 * fPercent;
   double i1 = a, i2 = a + fPercent;

   for( ; i2 <= b; i1 += fPercent, i2 += fPercent )
     i += Gauss( pF, NULL, i1, i2, power );

   if( i2 > b && i1 < b )
	 i += Gauss( pF, NULL, i1, b, power );

   return i;
 }

static const int __declspec(uuid("00000000-0000-0000-c000-000000000046")) *  volatile iVar;  


void TstVar()
 {
   VARIANT v1, v2; VariantInit( &v1 ), VariantInit( &v2 );
   v2.vt = v1.vt = VT_R8;

   v1.dblVal = 10, v2.dblVal = 20;
   HRESULT hr = VarCmp( &v1, &v2, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

   v1.dblVal = 20, v2.dblVal = 20;
   hr = VarCmp( &v1, &v2, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

   v1.dblVal = 21, v2.dblVal = 20;
   hr = VarCmp( &v1, &v2, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

   COleCurrency cc1( 100, 0 );
   COleCurrency cc2( 200, 0 );
   COleVariant vv1( cc1 ), vv2( cc2 );
   hr = VarCmp( (LPVARIANT)vv1, (LPVARIANT)vv2, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );
   hr = VarCmp( (LPVARIANT)vv1, &v1, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );
   hr = VarCmp( &v1, (LPVARIANT)vv1, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 0 );

   COleVariant vvt;
   hr = VarDiv( (LPVARIANT)vv1, &v2, &vvt );
   hr = VarDiv( &v2, (LPVARIANT)vv1, &vvt );
   hr = hr;

   basic_stringstream<wchar_t> strm;
   strm << L"Test 111";
   basic_string<wchar_t> s1 = strm.str().c_str();
   //strm.seekp( 0, ios::beg );
   strm.str( L"" );
   strm << L"Test 2";
   basic_string<wchar_t> s2 = strm.str().c_str();
 }

int main(int argc, char* argv[])
 {    

   TstVar();
   return 0;

   double i, i2, i3, a = 1;
   double i1_ = 0.0, i2_ = 85, step = 0.1;

   /*basic_fstream<TCHAR> fStrm( _T("e:\ii.txt"), ios::in|ios::out|ios::binary|ios::trunc );
   if( !fStrm ) throw exception( _T("Can't create/open output file") );

   fStrm << setprecision( 15 );

   fStrm << "Gauss 16:" << endl;   
   i = Gauss( fG, NULL, i1_, i2_, 16 );
   fStrm << setw( 20 ) << i << "\t" << setw( 20 ) << (i - a) << endl;
   fStrm << "Gauss 20:" << endl;   
   i = Gauss( fG, NULL, i1_, i2_, 20 );
   fStrm << setw( 20 ) << i << "\t" << setw( 20 ) << (i - a) << endl;
   fStrm << "Gauss 32:" << endl;   
   i = Gauss( fG, NULL, i1_, i2_, 32 );
   fStrm << setw( 20 ) << i << "\t" << setw( 20 ) << (i - a) << endl;

   fStrm << "Gauss 16: 0.1 - 10" << endl;      
   for( double dStep = 0.1; dStep <= 10; dStep += 0.1 )
	{
	  DWORD dw1 = GetTickCount();
      i = GaussIter( fG, NULL, i1_, i2_, 16, dStep );
	  fStrm << (GetTickCount() - dw1) << "\t";
      fStrm << dStep << "\t"  << setw(20) << (i - a) << endl;
	}



   fStrm.close();
   return 0;*/

   

    //printf( "Analytic = \t\t%1.12f\n", (a=f1i(0, 5)) ); 

	/*i = Gauss(f1, NULL, 0, 5, 8);
	printf( "Gauss 8 = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );

	i = Gauss(f1, NULL, 0, 5, 12);
	printf( "Gauss 12 = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );

	i = Gauss(f1, NULL, 0, 5, 20);
	printf( "Gauss 20 = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );*/

	i = Gauss(fG, NULL, i1_, i2_, 32);
	printf( "Gauss 32 = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );

	i = GaussIter( fG, NULL, i1_, i2_, 32, 0.7 );
	printf( "GaussIter 1% = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );

	i = GaussIter( fG, NULL, i1_, i2_, 32, 5 );
	printf( "GaussIter 5% = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );

	i = GaussIter( fG, NULL, i1_, i2_, 32, 10 );
	printf( "GaussIter 10% = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );

	DWORD dw1 = GetTickCount();
	i = Simpson(fG, NULL, i1_, i2_, step);
	printf( "%d\n", GetTickCount() - dw1 );
	printf( "Simpson = \t\t%1.12f \t%1.12f\n", i, fabs(a-i) );
	

    i2 = Ailer(fG, NULL, fGp, NULL, i1_, i2_, step);
	printf( "Ailer = \t\t%1.12f \t%1.12f\n", i2, fabs(a-i2) );

	i3 = Trapecion(fG, NULL, i1_, i2_, step);
	printf( "Trapecion = \t\t%1.12f \t%1.12f\n", i3, fabs(a-i3) );

	i3 = Median(fG, NULL, i1_, i2_, step);
	printf( "Median = \t\t%1.12f \t%1.12f\n", i3, fabs(a-i3) );

	i3 = MedianTrap(fG, NULL, i1_, i2_, step);
	printf( "MedianTrap = \t\t%1.12f \t%1.12f\n", i3, fabs(a-i3) );

	//printf( "pp = %1.12f\n", fGp(1, NULL) );
    getch();
	return 0;
 }

