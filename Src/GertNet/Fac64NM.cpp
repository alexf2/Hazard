#include "stdafx.h"
#include "Fac64NM.h"

__int64 __fastcall Fac64From( __int64 i64Arg, __int64 i64From )
 {
   __int64 i = i64Arg;
   i64Arg = 1;
   for( ; i >= i64From; --i )
	 i64Arg *= i;

   return i64Arg;
 }

__int64 __fastcall Fac64MN( __int64 i64M, __int64 i64N )
 {
   __int64 i64MmN = i64M - i64N;

   /*__int64 k1 = (i64M == 1) ? 1:2, //m
	       k2 = (i64N == 1) ? 1:2, //n
		   k3 = (i64MmN == 1) ? 1:2; //m-n*/
   __int64 k1, k2, k3;
   k1 = k2 = k3 = 1;

   for( ; k1 <= i64M && k2 <= i64N; ++k1, ++k2 );

   if( i64MmN >= k2 )
	{
	  k3 = k2;
      for( ; k1 <= i64M && k3 <= i64MmN; ++k1, ++k3 );
	}

   
   return Fac64From( i64M, k1 )  / 
     (Fac64From( i64N, k2 ) * Fac64From( i64MmN, k3 ) * (k3 >= k2 ? Fac64From( k2 - 1, 1 ):1));
 }

