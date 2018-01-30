#include "stdafx.h"
#include <new>
#include <new.h>

#pragma init_seg(lib)

int my_new_handler( size_t )  
 {
   throw std::bad_alloc();
   return 0;
 }

struct my_new_handler_obj
 {
   _PNH _old_new_handler;
   int _old_new_mode;

   my_new_handler_obj() throw()
	{
      _old_new_mode = _set_new_mode( 0 );
      _old_new_handler = _set_new_handler( my_new_handler );
    }

   ~my_new_handler_obj() throw()
	{
      _set_new_handler( _old_new_handler );
      _set_new_mode( _old_new_mode );
    }
 
 } _g_new_handler_obj;

