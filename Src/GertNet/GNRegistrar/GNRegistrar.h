/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Tue Nov 07 06:04:38 2000
 */
/* Compiler settings for G:\WORK\MAG\Alexf\GertNet\GNRegistrar\GNRegistrar.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __GNRegistrar_h__
#define __GNRegistrar_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IRegistrarHelper_FWD_DEFINED__
#define __IRegistrarHelper_FWD_DEFINED__
typedef interface IRegistrarHelper IRegistrarHelper;
#endif 	/* __IRegistrarHelper_FWD_DEFINED__ */


#ifndef __RegistrarHelper_FWD_DEFINED__
#define __RegistrarHelper_FWD_DEFINED__

#ifdef __cplusplus
typedef class RegistrarHelper RegistrarHelper;
#else
typedef struct RegistrarHelper RegistrarHelper;
#endif /* __cplusplus */

#endif 	/* __RegistrarHelper_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IRegistrarHelper_INTERFACE_DEFINED__
#define __IRegistrarHelper_INTERFACE_DEFINED__

/* interface IRegistrarHelper */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IRegistrarHelper;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("56955A41-B40D-11D4-905F-00504E02C39D")
    IRegistrarHelper : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Register( 
            /* [in] */ BSTR ObjName,
            /* [in] */ BSTR Path,
            /* [in] */ short RegFlag,
            /* [in] */ short TypeLibFlag,
            /* [in] */ BSTR PathHelp) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Hresult( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LastError( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IRegistrarHelperVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IRegistrarHelper __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IRegistrarHelper __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Register )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [in] */ BSTR ObjName,
            /* [in] */ BSTR Path,
            /* [in] */ short RegFlag,
            /* [in] */ short TypeLibFlag,
            /* [in] */ BSTR PathHelp);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Hresult )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LastError )( 
            IRegistrarHelper __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        END_INTERFACE
    } IRegistrarHelperVtbl;

    interface IRegistrarHelper
    {
        CONST_VTBL struct IRegistrarHelperVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRegistrarHelper_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IRegistrarHelper_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IRegistrarHelper_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IRegistrarHelper_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IRegistrarHelper_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IRegistrarHelper_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IRegistrarHelper_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IRegistrarHelper_Register(This,ObjName,Path,RegFlag,TypeLibFlag,PathHelp)	\
    (This)->lpVtbl -> Register(This,ObjName,Path,RegFlag,TypeLibFlag,PathHelp)

#define IRegistrarHelper_get_Hresult(This,pVal)	\
    (This)->lpVtbl -> get_Hresult(This,pVal)

#define IRegistrarHelper_get_LastError(This,pVal)	\
    (This)->lpVtbl -> get_LastError(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IRegistrarHelper_Register_Proxy( 
    IRegistrarHelper __RPC_FAR * This,
    /* [in] */ BSTR ObjName,
    /* [in] */ BSTR Path,
    /* [in] */ short RegFlag,
    /* [in] */ short TypeLibFlag,
    /* [in] */ BSTR PathHelp);


void __RPC_STUB IRegistrarHelper_Register_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IRegistrarHelper_get_Hresult_Proxy( 
    IRegistrarHelper __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IRegistrarHelper_get_Hresult_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IRegistrarHelper_get_LastError_Proxy( 
    IRegistrarHelper __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IRegistrarHelper_get_LastError_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IRegistrarHelper_INTERFACE_DEFINED__ */



#ifndef __GNREGISTRARLib_LIBRARY_DEFINED__
#define __GNREGISTRARLib_LIBRARY_DEFINED__

/* library GNREGISTRARLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_GNREGISTRARLib;

EXTERN_C const CLSID CLSID_RegistrarHelper;

#ifdef __cplusplus

class DECLSPEC_UUID("56955A42-B40D-11D4-905F-00504E02C39D")
RegistrarHelper;
#endif
#endif /* __GNREGISTRARLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
