#ifndef __PASSERR_H_
#define __PASSERR_H_

#include "resource.h"       // main symbols

void PassError( LPCOLESTR lp, HRESULT hr, const CLSID& clsid, const IID& iid );


#endif 
