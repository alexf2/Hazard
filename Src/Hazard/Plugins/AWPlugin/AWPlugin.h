/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Oct 29 03:31:30 2000
 */
/* Compiler settings for G:\WORK\MAG\Alexf\Hazard\Plugins\AWPlugin\AWPlugin.idl:
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

#ifndef __AWPlugin_h__
#define __AWPlugin_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IGostTable_FWD_DEFINED__
#define __IGostTable_FWD_DEFINED__
typedef interface IGostTable IGostTable;
#endif 	/* __IGostTable_FWD_DEFINED__ */


#ifndef __IStCollItem_FWD_DEFINED__
#define __IStCollItem_FWD_DEFINED__
typedef interface IStCollItem IStCollItem;
#endif 	/* __IStCollItem_FWD_DEFINED__ */


#ifndef __ICollGosts_FWD_DEFINED__
#define __ICollGosts_FWD_DEFINED__
typedef interface ICollGosts ICollGosts;
#endif 	/* __ICollGosts_FWD_DEFINED__ */


#ifndef __ICollTopics_FWD_DEFINED__
#define __ICollTopics_FWD_DEFINED__
typedef interface ICollTopics ICollTopics;
#endif 	/* __ICollTopics_FWD_DEFINED__ */


#ifndef __IFactorTable_FWD_DEFINED__
#define __IFactorTable_FWD_DEFINED__
typedef interface IFactorTable IFactorTable;
#endif 	/* __IFactorTable_FWD_DEFINED__ */


#ifndef __ICollFTables_FWD_DEFINED__
#define __ICollFTables_FWD_DEFINED__
typedef interface ICollFTables ICollFTables;
#endif 	/* __ICollFTables_FWD_DEFINED__ */


#ifndef __IIntCreator_FWD_DEFINED__
#define __IIntCreator_FWD_DEFINED__
typedef interface IIntCreator IIntCreator;
#endif 	/* __IIntCreator_FWD_DEFINED__ */


#ifndef __ICollFactors_FWD_DEFINED__
#define __ICollFactors_FWD_DEFINED__
typedef interface ICollFactors ICollFactors;
#endif 	/* __ICollFactors_FWD_DEFINED__ */


#ifndef __IStCollItem_FWD_DEFINED__
#define __IStCollItem_FWD_DEFINED__
typedef interface IStCollItem IStCollItem;
#endif 	/* __IStCollItem_FWD_DEFINED__ */


#ifndef __ICollGosts_FWD_DEFINED__
#define __ICollGosts_FWD_DEFINED__
typedef interface ICollGosts ICollGosts;
#endif 	/* __ICollGosts_FWD_DEFINED__ */


#ifndef __CollTopics_FWD_DEFINED__
#define __CollTopics_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollTopics CollTopics;
#else
typedef struct CollTopics CollTopics;
#endif /* __cplusplus */

#endif 	/* __CollTopics_FWD_DEFINED__ */


#ifndef __CollFTables_FWD_DEFINED__
#define __CollFTables_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollFTables CollFTables;
#else
typedef struct CollFTables CollFTables;
#endif /* __cplusplus */

#endif 	/* __CollFTables_FWD_DEFINED__ */


#ifndef __IntCreator_FWD_DEFINED__
#define __IntCreator_FWD_DEFINED__

#ifdef __cplusplus
typedef class IntCreator IntCreator;
#else
typedef struct IntCreator IntCreator;
#endif /* __cplusplus */

#endif 	/* __IntCreator_FWD_DEFINED__ */


#ifndef __CollFactors_FWD_DEFINED__
#define __CollFactors_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollFactors CollFactors;
#else
typedef struct CollFactors CollFactors;
#endif /* __cplusplus */

#endif 	/* __CollFactors_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_AWPlugin_0000 */
/* [local] */ 


enum EN_AWPluginLib_DispIds
    {	dispid_DefaultCU	= 5
    };
typedef /* [uuid][public] */ 
enum ObjStatus_tag
    {	OS_None	= 0,
	OS_New	= 1,
	OS_Updated	= 2,
	OS_Deleted	= 3
    }	ObjStatus;

typedef /* [uuid][public] */ 
enum FTSlotTyp_tag
    {	FT_None	= 0,
	FT_Self	= 1,
	FT_Gost	= 2
    }	FTSlotTyp;

typedef /* [uuid][public] */ 
enum StdSortSeqv_tag
    {	SSS_AsIs	= 0,
	SSS_ByKeyAsc	= 1,
	SSS_ByKeyDesc	= 2,
	SSS_ByNameAsc	= 3,
	SSS_ByNameDesc	= 4
    }	StdSortSeqv;



extern RPC_IF_HANDLE __MIDL_itf_AWPlugin_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_AWPlugin_0000_v0_0_s_ifspec;

#ifndef __IGostTable_INTERFACE_DEFINED__
#define __IGostTable_INTERFACE_DEFINED__

/* interface IGostTable */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IGostTable;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F39A15E1-9890-11D4-900E-00504E02C39D")
    IGostTable : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NumberSlots( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InsertSlots( 
            /* [in] */ long Number,
            /* [defaultvalue][optional][in] */ long IdxStart = -1) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RemoveSlots( 
            /* [defaultvalue][optional][in] */ long Number = -1,
            /* [defaultvalue][optional][in] */ long IdxStart = -1) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NDescr( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NDescr( 
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NLingvoVal( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NLingvoVal( 
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NComment( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NComment( 
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE get_NValue( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][defaultcollelem][id][propput] */ HRESULT STDMETHODCALLTYPE put_NValue( 
            /* [in] */ long ZeroIndex,
            /* [in] */ float newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ClmName( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ClmName( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetUniformGrad( 
            /* [in] */ VARIANT_BOOL Revers) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Check( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LastErrStrZeroIdx( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IGostTableVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IGostTable __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IGostTable __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IGostTable __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NumberSlots )( 
            IGostTable __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *InsertSlots )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long Number,
            /* [defaultvalue][optional][in] */ long IdxStart);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemoveSlots )( 
            IGostTable __RPC_FAR * This,
            /* [defaultvalue][optional][in] */ long Number,
            /* [defaultvalue][optional][in] */ long IdxStart);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NDescr )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NDescr )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NLingvoVal )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NLingvoVal )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NComment )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NComment )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][defaultcollelem][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NValue )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][defaultcollelem][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NValue )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ float newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ClmName )( 
            IGostTable __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ClmName )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetUniformGrad )( 
            IGostTable __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL Revers);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Check )( 
            IGostTable __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LastErrStrZeroIdx )( 
            IGostTable __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        END_INTERFACE
    } IGostTableVtbl;

    interface IGostTable
    {
        CONST_VTBL struct IGostTableVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IGostTable_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IGostTable_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IGostTable_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IGostTable_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IGostTable_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IGostTable_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IGostTable_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IGostTable_get_NumberSlots(This,pVal)	\
    (This)->lpVtbl -> get_NumberSlots(This,pVal)

#define IGostTable_InsertSlots(This,Number,IdxStart)	\
    (This)->lpVtbl -> InsertSlots(This,Number,IdxStart)

#define IGostTable_RemoveSlots(This,Number,IdxStart)	\
    (This)->lpVtbl -> RemoveSlots(This,Number,IdxStart)

#define IGostTable_get_NDescr(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NDescr(This,ZeroIndex,pVal)

#define IGostTable_put_NDescr(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NDescr(This,ZeroIndex,newVal)

#define IGostTable_get_NLingvoVal(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NLingvoVal(This,ZeroIndex,pVal)

#define IGostTable_put_NLingvoVal(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NLingvoVal(This,ZeroIndex,newVal)

#define IGostTable_get_NComment(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NComment(This,ZeroIndex,pVal)

#define IGostTable_put_NComment(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NComment(This,ZeroIndex,newVal)

#define IGostTable_get_NValue(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NValue(This,ZeroIndex,pVal)

#define IGostTable_put_NValue(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NValue(This,ZeroIndex,newVal)

#define IGostTable_get_ClmName(This,pVal)	\
    (This)->lpVtbl -> get_ClmName(This,pVal)

#define IGostTable_put_ClmName(This,newVal)	\
    (This)->lpVtbl -> put_ClmName(This,newVal)

#define IGostTable_SetUniformGrad(This,Revers)	\
    (This)->lpVtbl -> SetUniformGrad(This,Revers)

#define IGostTable_Check(This)	\
    (This)->lpVtbl -> Check(This)

#define IGostTable_get_LastErrStrZeroIdx(This,pVal)	\
    (This)->lpVtbl -> get_LastErrStrZeroIdx(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_NumberSlots_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_NumberSlots_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGostTable_InsertSlots_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long Number,
    /* [defaultvalue][optional][in] */ long IdxStart);


void __RPC_STUB IGostTable_InsertSlots_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGostTable_RemoveSlots_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [defaultvalue][optional][in] */ long Number,
    /* [defaultvalue][optional][in] */ long IdxStart);


void __RPC_STUB IGostTable_RemoveSlots_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_NDescr_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_NDescr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IGostTable_put_NDescr_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ BSTR newVal);


void __RPC_STUB IGostTable_put_NDescr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_NLingvoVal_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_NLingvoVal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IGostTable_put_NLingvoVal_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ BSTR newVal);


void __RPC_STUB IGostTable_put_NLingvoVal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_NComment_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_NComment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IGostTable_put_NComment_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ BSTR newVal);


void __RPC_STUB IGostTable_put_NComment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_NValue_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_NValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][defaultcollelem][id][propput] */ HRESULT STDMETHODCALLTYPE IGostTable_put_NValue_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ float newVal);


void __RPC_STUB IGostTable_put_NValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_ClmName_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_ClmName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IGostTable_put_ClmName_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IGostTable_put_ClmName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGostTable_SetUniformGrad_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL Revers);


void __RPC_STUB IGostTable_SetUniformGrad_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IGostTable_Check_Proxy( 
    IGostTable __RPC_FAR * This);


void __RPC_STUB IGostTable_Check_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IGostTable_get_LastErrStrZeroIdx_Proxy( 
    IGostTable __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IGostTable_get_LastErrStrZeroIdx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IGostTable_INTERFACE_DEFINED__ */


#ifndef __IStCollItem_INTERFACE_DEFINED__
#define __IStCollItem_INTERFACE_DEFINED__

/* interface IStCollItem */
/* [unique][helpstring][uuid][public][object] */ 


EXTERN_C const IID IID_IStCollItem;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("E3674891-9BC8-11d4-9012-00504E02C39D")
    IStCollItem : public IUnknown
    {
    public:
        virtual /* [defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE get_Key( 
            /* [retval][out] */ long __RPC_FAR *plKey) = 0;
        
        virtual /* [defaultcollelem][hidden][restricted][id][propput] */ HRESULT STDMETHODCALLTYPE put_Key( 
            /* [in] */ long lKeyNew) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Status( 
            /* [retval][out] */ ObjStatus __RPC_FAR *pStatus) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pbsName) = 0;
        
        virtual /* [hidden][restricted][id][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR bsNameNew) = 0;
        
        virtual /* [hidden][restricted][id][propput] */ HRESULT STDMETHODCALLTYPE put_Sign( 
            /* [in] */ long lSignNew) = 0;
        
        virtual /* [hidden][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get_Sign( 
            /* [retval][out] */ long __RPC_FAR *plSign) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Storage( 
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *Stg) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_DefaultCU( 
            /* [in] */ VARIANT_BOOL bMode) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_DefaultCU( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbMode) = 0;
        
        virtual /* [local] */ BSTR STDMETHODCALLTYPE NameByRef( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IStCollItemVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IStCollItem __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IStCollItem __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IStCollItem __RPC_FAR * This);
        
        /* [defaultcollelem][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Key )( 
            IStCollItem __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *plKey);
        
        /* [defaultcollelem][hidden][restricted][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Key )( 
            IStCollItem __RPC_FAR * This,
            /* [in] */ long lKeyNew);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Status )( 
            IStCollItem __RPC_FAR * This,
            /* [retval][out] */ ObjStatus __RPC_FAR *pStatus);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            IStCollItem __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName);
        
        /* [hidden][restricted][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            IStCollItem __RPC_FAR * This,
            /* [in] */ BSTR bsNameNew);
        
        /* [hidden][restricted][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Sign )( 
            IStCollItem __RPC_FAR * This,
            /* [in] */ long lSignNew);
        
        /* [hidden][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Sign )( 
            IStCollItem __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *plSign);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Storage )( 
            IStCollItem __RPC_FAR * This,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *Stg);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DefaultCU )( 
            IStCollItem __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bMode);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DefaultCU )( 
            IStCollItem __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbMode);
        
        /* [local] */ BSTR ( STDMETHODCALLTYPE __RPC_FAR *NameByRef )( 
            IStCollItem __RPC_FAR * This);
        
        END_INTERFACE
    } IStCollItemVtbl;

    interface IStCollItem
    {
        CONST_VTBL struct IStCollItemVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IStCollItem_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IStCollItem_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IStCollItem_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IStCollItem_get_Key(This,plKey)	\
    (This)->lpVtbl -> get_Key(This,plKey)

#define IStCollItem_put_Key(This,lKeyNew)	\
    (This)->lpVtbl -> put_Key(This,lKeyNew)

#define IStCollItem_get_Status(This,pStatus)	\
    (This)->lpVtbl -> get_Status(This,pStatus)

#define IStCollItem_get_Name(This,pbsName)	\
    (This)->lpVtbl -> get_Name(This,pbsName)

#define IStCollItem_put_Name(This,bsNameNew)	\
    (This)->lpVtbl -> put_Name(This,bsNameNew)

#define IStCollItem_put_Sign(This,lSignNew)	\
    (This)->lpVtbl -> put_Sign(This,lSignNew)

#define IStCollItem_get_Sign(This,plSign)	\
    (This)->lpVtbl -> get_Sign(This,plSign)

#define IStCollItem_get_Storage(This,Stg)	\
    (This)->lpVtbl -> get_Storage(This,Stg)

#define IStCollItem_put_DefaultCU(This,bMode)	\
    (This)->lpVtbl -> put_DefaultCU(This,bMode)

#define IStCollItem_get_DefaultCU(This,pbMode)	\
    (This)->lpVtbl -> get_DefaultCU(This,pbMode)

#define IStCollItem_NameByRef(This)	\
    (This)->lpVtbl -> NameByRef(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE IStCollItem_get_Key_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *plKey);


void __RPC_STUB IStCollItem_get_Key_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [defaultcollelem][hidden][restricted][id][propput] */ HRESULT STDMETHODCALLTYPE IStCollItem_put_Key_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [in] */ long lKeyNew);


void __RPC_STUB IStCollItem_put_Key_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IStCollItem_get_Status_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [retval][out] */ ObjStatus __RPC_FAR *pStatus);


void __RPC_STUB IStCollItem_get_Status_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IStCollItem_get_Name_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pbsName);


void __RPC_STUB IStCollItem_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][restricted][id][propput] */ HRESULT STDMETHODCALLTYPE IStCollItem_put_Name_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [in] */ BSTR bsNameNew);


void __RPC_STUB IStCollItem_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][restricted][id][propput] */ HRESULT STDMETHODCALLTYPE IStCollItem_put_Sign_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [in] */ long lSignNew);


void __RPC_STUB IStCollItem_put_Sign_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE IStCollItem_get_Sign_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *plSign);


void __RPC_STUB IStCollItem_get_Sign_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IStCollItem_get_Storage_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *Stg);


void __RPC_STUB IStCollItem_get_Storage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IStCollItem_put_DefaultCU_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bMode);


void __RPC_STUB IStCollItem_put_DefaultCU_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IStCollItem_get_DefaultCU_Proxy( 
    IStCollItem __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbMode);


void __RPC_STUB IStCollItem_get_DefaultCU_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [local] */ BSTR STDMETHODCALLTYPE IStCollItem_NameByRef_Proxy( 
    IStCollItem __RPC_FAR * This);


void __RPC_STUB IStCollItem_NameByRef_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IStCollItem_INTERFACE_DEFINED__ */


#ifndef __ICollGosts_INTERFACE_DEFINED__
#define __ICollGosts_INTERFACE_DEFINED__

/* interface ICollGosts */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollGosts;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F39A15E3-9890-11D4-900E-00504E02C39D")
    ICollGosts : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR strName,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LongKey( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IGostTable __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IGostTable __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Update( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CachedUpdates( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetUpdateMode( 
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ItemNth( 
            /* [in] */ long Number,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_EnumSorting( 
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_EnumSorting( 
            /* [in] */ StdSortSeqv SortSeq) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE KeyNth( 
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE NameNth( 
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollGostsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollGosts __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollGosts __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollGosts __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollGosts __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollGosts __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ BSTR strName,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LongKey )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NameKey )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IGostTable __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NameKey )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IGostTable __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollGosts __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Update )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CachedUpdates )( 
            ICollGosts __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetUpdateMode )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ItemNth )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnumSorting )( 
            ICollGosts __RPC_FAR * This,
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnumSorting )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ StdSortSeqv SortSeq);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *KeyNth )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NameNth )( 
            ICollGosts __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName);
        
        END_INTERFACE
    } ICollGostsVtbl;

    interface ICollGosts
    {
        CONST_VTBL struct ICollGostsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollGosts_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollGosts_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollGosts_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollGosts_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollGosts_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollGosts_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollGosts_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollGosts_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollGosts_Item(This,Key,ppGT)	\
    (This)->lpVtbl -> Item(This,Key,ppGT)

#define ICollGosts_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollGosts_Add(This,strName,ppGT)	\
    (This)->lpVtbl -> Add(This,strName,ppGT)

#define ICollGosts_Remove(This,KeyOrObj)	\
    (This)->lpVtbl -> Remove(This,KeyOrObj)

#define ICollGosts_get_LongKey(This,Name,plKey)	\
    (This)->lpVtbl -> get_LongKey(This,Name,plKey)

#define ICollGosts_get_NameKey(This,Key,Obj,pbsName)	\
    (This)->lpVtbl -> get_NameKey(This,Key,Obj,pbsName)

#define ICollGosts_put_NameKey(This,Key,Obj,bsNewName)	\
    (This)->lpVtbl -> put_NameKey(This,Key,Obj,bsNewName)

#define ICollGosts_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define ICollGosts_Update(This,KeyOrObj)	\
    (This)->lpVtbl -> Update(This,KeyOrObj)

#define ICollGosts_get_CachedUpdates(This,pbCached)	\
    (This)->lpVtbl -> get_CachedUpdates(This,pbCached)

#define ICollGosts_SetUpdateMode(This,Stm,bMode)	\
    (This)->lpVtbl -> SetUpdateMode(This,Stm,bMode)

#define ICollGosts_ItemNth(This,Number,ppGT)	\
    (This)->lpVtbl -> ItemNth(This,Number,ppGT)

#define ICollGosts_get_EnumSorting(This,pSortSeq)	\
    (This)->lpVtbl -> get_EnumSorting(This,pSortSeq)

#define ICollGosts_put_EnumSorting(This,SortSeq)	\
    (This)->lpVtbl -> put_EnumSorting(This,SortSeq)

#define ICollGosts_KeyNth(This,Number,lKey)	\
    (This)->lpVtbl -> KeyNth(This,Number,lKey)

#define ICollGosts_NameNth(This,Number,bsName)	\
    (This)->lpVtbl -> NameNth(This,Number,bsName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollGosts_get__NewEnum_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollGosts_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollGosts_Item_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *Key,
    /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollGosts_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollGosts_get_Count_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollGosts_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollGosts_Add_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ BSTR strName,
    /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollGosts_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollGosts_Remove_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollGosts_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollGosts_get_LongKey_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ long __RPC_FAR *plKey);


void __RPC_STUB ICollGosts_get_LongKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollGosts_get_NameKey_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ IGostTable __RPC_FAR *Obj,
    /* [retval][out] */ BSTR __RPC_FAR *pbsName);


void __RPC_STUB ICollGosts_get_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollGosts_put_NameKey_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ IGostTable __RPC_FAR *Obj,
    /* [in] */ BSTR bsNewName);


void __RPC_STUB ICollGosts_put_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollGosts_Clear_Proxy( 
    ICollGosts __RPC_FAR * This);


void __RPC_STUB ICollGosts_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollGosts_Update_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollGosts_Update_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollGosts_get_CachedUpdates_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);


void __RPC_STUB ICollGosts_get_CachedUpdates_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollGosts_SetUpdateMode_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stm,
    /* [in] */ VARIANT_BOOL bMode);


void __RPC_STUB ICollGosts_SetUpdateMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollGosts_ItemNth_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollGosts_ItemNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollGosts_get_EnumSorting_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);


void __RPC_STUB ICollGosts_get_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICollGosts_put_EnumSorting_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ StdSortSeqv SortSeq);


void __RPC_STUB ICollGosts_put_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollGosts_KeyNth_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ long __RPC_FAR *lKey);


void __RPC_STUB ICollGosts_KeyNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollGosts_NameNth_Proxy( 
    ICollGosts __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ BSTR __RPC_FAR *bsName);


void __RPC_STUB ICollGosts_NameNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollGosts_INTERFACE_DEFINED__ */


#ifndef __ICollTopics_INTERFACE_DEFINED__
#define __ICollTopics_INTERFACE_DEFINED__

/* interface ICollTopics */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollTopics;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F39A15E5-9890-11D4-900E-00504E02C39D")
    ICollTopics : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR strName,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LongKey( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollGosts __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollGosts __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Update( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CachedUpdates( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetUpdateMode( 
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ItemNth( 
            /* [in] */ long Number,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_EnumSorting( 
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_EnumSorting( 
            /* [in] */ StdSortSeqv SortSeq) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE KeyNth( 
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE NameNth( 
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollTopicsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollTopics __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollTopics __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollTopics __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollTopics __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollTopics __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ BSTR strName,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LongKey )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NameKey )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollGosts __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NameKey )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollGosts __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollTopics __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Update )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CachedUpdates )( 
            ICollTopics __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetUpdateMode )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ItemNth )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnumSorting )( 
            ICollTopics __RPC_FAR * This,
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnumSorting )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ StdSortSeqv SortSeq);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *KeyNth )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NameNth )( 
            ICollTopics __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName);
        
        END_INTERFACE
    } ICollTopicsVtbl;

    interface ICollTopics
    {
        CONST_VTBL struct ICollTopicsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollTopics_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollTopics_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollTopics_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollTopics_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollTopics_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollTopics_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollTopics_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollTopics_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollTopics_Item(This,Key,ppGT)	\
    (This)->lpVtbl -> Item(This,Key,ppGT)

#define ICollTopics_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollTopics_Add(This,strName,ppGT)	\
    (This)->lpVtbl -> Add(This,strName,ppGT)

#define ICollTopics_Remove(This,KeyOrObj)	\
    (This)->lpVtbl -> Remove(This,KeyOrObj)

#define ICollTopics_get_LongKey(This,Name,plKey)	\
    (This)->lpVtbl -> get_LongKey(This,Name,plKey)

#define ICollTopics_get_NameKey(This,Key,Obj,pbsName)	\
    (This)->lpVtbl -> get_NameKey(This,Key,Obj,pbsName)

#define ICollTopics_put_NameKey(This,Key,Obj,bsNewName)	\
    (This)->lpVtbl -> put_NameKey(This,Key,Obj,bsNewName)

#define ICollTopics_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define ICollTopics_Update(This,KeyOrObj)	\
    (This)->lpVtbl -> Update(This,KeyOrObj)

#define ICollTopics_get_CachedUpdates(This,pbCached)	\
    (This)->lpVtbl -> get_CachedUpdates(This,pbCached)

#define ICollTopics_SetUpdateMode(This,Stm,bMode)	\
    (This)->lpVtbl -> SetUpdateMode(This,Stm,bMode)

#define ICollTopics_ItemNth(This,Number,ppGT)	\
    (This)->lpVtbl -> ItemNth(This,Number,ppGT)

#define ICollTopics_get_EnumSorting(This,pSortSeq)	\
    (This)->lpVtbl -> get_EnumSorting(This,pSortSeq)

#define ICollTopics_put_EnumSorting(This,SortSeq)	\
    (This)->lpVtbl -> put_EnumSorting(This,SortSeq)

#define ICollTopics_KeyNth(This,Number,lKey)	\
    (This)->lpVtbl -> KeyNth(This,Number,lKey)

#define ICollTopics_NameNth(This,Number,bsName)	\
    (This)->lpVtbl -> NameNth(This,Number,bsName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollTopics_get__NewEnum_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollTopics_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollTopics_Item_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *Key,
    /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollTopics_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollTopics_get_Count_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollTopics_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollTopics_Add_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ BSTR strName,
    /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollTopics_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollTopics_Remove_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollTopics_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollTopics_get_LongKey_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ long __RPC_FAR *plKey);


void __RPC_STUB ICollTopics_get_LongKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollTopics_get_NameKey_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ ICollGosts __RPC_FAR *Obj,
    /* [retval][out] */ BSTR __RPC_FAR *pbsName);


void __RPC_STUB ICollTopics_get_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollTopics_put_NameKey_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ ICollGosts __RPC_FAR *Obj,
    /* [in] */ BSTR bsNewName);


void __RPC_STUB ICollTopics_put_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollTopics_Clear_Proxy( 
    ICollTopics __RPC_FAR * This);


void __RPC_STUB ICollTopics_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollTopics_Update_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollTopics_Update_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollTopics_get_CachedUpdates_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);


void __RPC_STUB ICollTopics_get_CachedUpdates_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollTopics_SetUpdateMode_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stm,
    /* [in] */ VARIANT_BOOL bMode);


void __RPC_STUB ICollTopics_SetUpdateMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollTopics_ItemNth_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollTopics_ItemNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollTopics_get_EnumSorting_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);


void __RPC_STUB ICollTopics_get_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICollTopics_put_EnumSorting_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ StdSortSeqv SortSeq);


void __RPC_STUB ICollTopics_put_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollTopics_KeyNth_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ long __RPC_FAR *lKey);


void __RPC_STUB ICollTopics_KeyNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollTopics_NameNth_Proxy( 
    ICollTopics __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ BSTR __RPC_FAR *bsName);


void __RPC_STUB ICollTopics_NameNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollTopics_INTERFACE_DEFINED__ */


#ifndef __IFactorTable_INTERFACE_DEFINED__
#define __IFactorTable_INTERFACE_DEFINED__

/* interface IFactorTable */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IFactorTable;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F39A15EA-9890-11D4-900E-00504E02C39D")
    IFactorTable : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InsertSlots( 
            /* [in] */ long Number,
            /* [defaultvalue][optional][in] */ long IdxStart = -1) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RemoveSlots( 
            /* [defaultvalue][optional][in] */ long Number = -1,
            /* [defaultvalue][optional][in] */ long IdxStart = -1) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NumberSlots( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NComponentName( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NComponentName( 
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NComment( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NComment( 
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE get_NWeight( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][defaultcollelem][id][propput] */ HRESULT STDMETHODCALLTYPE put_NWeight( 
            /* [in] */ long ZeroIndex,
            /* [in] */ float newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NValuation( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NValuation( 
            /* [in] */ long ZeroIndex,
            /* [in] */ VARIANT __RPC_FAR *pfVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NAveWeighted( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_WeightSumm( 
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ValuationSumm( 
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AveWeightedSumm( 
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_AllComponentsIsAssigned( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bRes) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetWeights( 
            /* [optional][in] */ VARIANT StartIdx) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsWeightsCorrect( 
            /* [optional][out][in] */ VARIANT __RPC_FAR *Summ,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CalcValues( 
            /* [optional][out][in] */ VARIANT __RPC_FAR *fValuationSumm,
            /* [optional][out][in] */ VARIANT __RPC_FAR *fAveWeightedSumm) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NSlotType( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ FTSlotTyp __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NSlotValue( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NSlotValue( 
            /* [in] */ long ZeroIndex,
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NAssSlotType( 
            /* [in] */ long ZeroIndex,
            /* [in] */ FTSlotTyp NewTyp,
            /* [optional][in] */ long K1,
            /* [optional][in] */ long K2) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NBindSlot( 
            /* [in] */ long ZeroIndex,
            /* [in] */ IDispatch __RPC_FAR *Object) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NSlotSelf( 
            /* [in] */ long ZeroIndex,
            /* [optional][out] */ long __RPC_FAR *Key,
            /* [optional][out] */ IFactorTable __RPC_FAR *__RPC_FAR *IFT) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NSlotGost( 
            /* [in] */ long ZeroIndex,
            /* [optional][out] */ long __RPC_FAR *KeyTopic,
            /* [optional][out] */ long __RPC_FAR *KeyGost,
            /* [optional][out] */ IGostTable __RPC_FAR *__RPC_FAR *IGT) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE HandsOffObjects( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetRefToFTable( 
            /* [in] */ long Key,
            /* [retval][out] */ VARIANT __RPC_FAR *RefSlot) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetRefToGost( 
            /* [in] */ long KeyTopic,
            /* [in] */ long KeyGost,
            /* [retval][out] */ VARIANT __RPC_FAR *RefSlot) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NormalizeWeights( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NLocked( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NLocked( 
            /* [in] */ long ZeroIndex,
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NLingvoVal( 
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NLingvoVal( 
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR Val) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IFactorTableVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IFactorTable __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IFactorTable __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IFactorTable __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *InsertSlots )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long Number,
            /* [defaultvalue][optional][in] */ long IdxStart);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemoveSlots )( 
            IFactorTable __RPC_FAR * This,
            /* [defaultvalue][optional][in] */ long Number,
            /* [defaultvalue][optional][in] */ long IdxStart);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NumberSlots )( 
            IFactorTable __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NComponentName )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NComponentName )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NComment )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NComment )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][defaultcollelem][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NWeight )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][defaultcollelem][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NWeight )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ float newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NValuation )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NValuation )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ VARIANT __RPC_FAR *pfVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NAveWeighted )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_WeightSumm )( 
            IFactorTable __RPC_FAR * This,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ValuationSumm )( 
            IFactorTable __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AveWeightedSumm )( 
            IFactorTable __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pfVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AllComponentsIsAssigned )( 
            IFactorTable __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bRes);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetWeights )( 
            IFactorTable __RPC_FAR * This,
            /* [optional][in] */ VARIANT StartIdx);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsWeightsCorrect )( 
            IFactorTable __RPC_FAR * This,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Summ,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CalcValues )( 
            IFactorTable __RPC_FAR * This,
            /* [optional][out][in] */ VARIANT __RPC_FAR *fValuationSumm,
            /* [optional][out][in] */ VARIANT __RPC_FAR *fAveWeightedSumm);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NSlotType )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ FTSlotTyp __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NSlotValue )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NSlotValue )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ long newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NAssSlotType )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ FTSlotTyp NewTyp,
            /* [optional][in] */ long K1,
            /* [optional][in] */ long K2);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NBindSlot )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ IDispatch __RPC_FAR *Object);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NSlotSelf )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [optional][out] */ long __RPC_FAR *Key,
            /* [optional][out] */ IFactorTable __RPC_FAR *__RPC_FAR *IFT);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NSlotGost )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [optional][out] */ long __RPC_FAR *KeyTopic,
            /* [optional][out] */ long __RPC_FAR *KeyGost,
            /* [optional][out] */ IGostTable __RPC_FAR *__RPC_FAR *IGT);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *HandsOffObjects )( 
            IFactorTable __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetRefToFTable )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long Key,
            /* [retval][out] */ VARIANT __RPC_FAR *RefSlot);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetRefToGost )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long KeyTopic,
            /* [in] */ long KeyGost,
            /* [retval][out] */ VARIANT __RPC_FAR *RefSlot);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NormalizeWeights )( 
            IFactorTable __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NLocked )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NLocked )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NLingvoVal )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NLingvoVal )( 
            IFactorTable __RPC_FAR * This,
            /* [in] */ long ZeroIndex,
            /* [in] */ BSTR Val);
        
        END_INTERFACE
    } IFactorTableVtbl;

    interface IFactorTable
    {
        CONST_VTBL struct IFactorTableVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IFactorTable_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IFactorTable_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IFactorTable_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IFactorTable_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IFactorTable_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IFactorTable_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IFactorTable_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IFactorTable_InsertSlots(This,Number,IdxStart)	\
    (This)->lpVtbl -> InsertSlots(This,Number,IdxStart)

#define IFactorTable_RemoveSlots(This,Number,IdxStart)	\
    (This)->lpVtbl -> RemoveSlots(This,Number,IdxStart)

#define IFactorTable_get_NumberSlots(This,pVal)	\
    (This)->lpVtbl -> get_NumberSlots(This,pVal)

#define IFactorTable_get_NComponentName(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NComponentName(This,ZeroIndex,pVal)

#define IFactorTable_put_NComponentName(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NComponentName(This,ZeroIndex,newVal)

#define IFactorTable_get_NComment(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NComment(This,ZeroIndex,pVal)

#define IFactorTable_put_NComment(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NComment(This,ZeroIndex,newVal)

#define IFactorTable_get_NWeight(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NWeight(This,ZeroIndex,pVal)

#define IFactorTable_put_NWeight(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NWeight(This,ZeroIndex,newVal)

#define IFactorTable_get_NValuation(This,ZeroIndex,pfVal)	\
    (This)->lpVtbl -> get_NValuation(This,ZeroIndex,pfVal)

#define IFactorTable_put_NValuation(This,ZeroIndex,pfVal)	\
    (This)->lpVtbl -> put_NValuation(This,ZeroIndex,pfVal)

#define IFactorTable_get_NAveWeighted(This,ZeroIndex,pfVal)	\
    (This)->lpVtbl -> get_NAveWeighted(This,ZeroIndex,pfVal)

#define IFactorTable_get_WeightSumm(This,pVal)	\
    (This)->lpVtbl -> get_WeightSumm(This,pVal)

#define IFactorTable_get_ValuationSumm(This,pfVal)	\
    (This)->lpVtbl -> get_ValuationSumm(This,pfVal)

#define IFactorTable_get_AveWeightedSumm(This,pfVal)	\
    (This)->lpVtbl -> get_AveWeightedSumm(This,pfVal)

#define IFactorTable_get_AllComponentsIsAssigned(This,bRes)	\
    (This)->lpVtbl -> get_AllComponentsIsAssigned(This,bRes)

#define IFactorTable_SetWeights(This,StartIdx)	\
    (This)->lpVtbl -> SetWeights(This,StartIdx)

#define IFactorTable_get_IsWeightsCorrect(This,Summ,pVal)	\
    (This)->lpVtbl -> get_IsWeightsCorrect(This,Summ,pVal)

#define IFactorTable_CalcValues(This,fValuationSumm,fAveWeightedSumm)	\
    (This)->lpVtbl -> CalcValues(This,fValuationSumm,fAveWeightedSumm)

#define IFactorTable_get_NSlotType(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NSlotType(This,ZeroIndex,pVal)

#define IFactorTable_get_NSlotValue(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NSlotValue(This,ZeroIndex,pVal)

#define IFactorTable_put_NSlotValue(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NSlotValue(This,ZeroIndex,newVal)

#define IFactorTable_NAssSlotType(This,ZeroIndex,NewTyp,K1,K2)	\
    (This)->lpVtbl -> NAssSlotType(This,ZeroIndex,NewTyp,K1,K2)

#define IFactorTable_NBindSlot(This,ZeroIndex,Object)	\
    (This)->lpVtbl -> NBindSlot(This,ZeroIndex,Object)

#define IFactorTable_NSlotSelf(This,ZeroIndex,Key,IFT)	\
    (This)->lpVtbl -> NSlotSelf(This,ZeroIndex,Key,IFT)

#define IFactorTable_NSlotGost(This,ZeroIndex,KeyTopic,KeyGost,IGT)	\
    (This)->lpVtbl -> NSlotGost(This,ZeroIndex,KeyTopic,KeyGost,IGT)

#define IFactorTable_HandsOffObjects(This)	\
    (This)->lpVtbl -> HandsOffObjects(This)

#define IFactorTable_GetRefToFTable(This,Key,RefSlot)	\
    (This)->lpVtbl -> GetRefToFTable(This,Key,RefSlot)

#define IFactorTable_GetRefToGost(This,KeyTopic,KeyGost,RefSlot)	\
    (This)->lpVtbl -> GetRefToGost(This,KeyTopic,KeyGost,RefSlot)

#define IFactorTable_NormalizeWeights(This)	\
    (This)->lpVtbl -> NormalizeWeights(This)

#define IFactorTable_get_NLocked(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NLocked(This,ZeroIndex,pVal)

#define IFactorTable_put_NLocked(This,ZeroIndex,newVal)	\
    (This)->lpVtbl -> put_NLocked(This,ZeroIndex,newVal)

#define IFactorTable_get_NLingvoVal(This,ZeroIndex,pVal)	\
    (This)->lpVtbl -> get_NLingvoVal(This,ZeroIndex,pVal)

#define IFactorTable_put_NLingvoVal(This,ZeroIndex,Val)	\
    (This)->lpVtbl -> put_NLingvoVal(This,ZeroIndex,Val)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_InsertSlots_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long Number,
    /* [defaultvalue][optional][in] */ long IdxStart);


void __RPC_STUB IFactorTable_InsertSlots_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_RemoveSlots_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [defaultvalue][optional][in] */ long Number,
    /* [defaultvalue][optional][in] */ long IdxStart);


void __RPC_STUB IFactorTable_RemoveSlots_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NumberSlots_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NumberSlots_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NComponentName_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NComponentName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NComponentName_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ BSTR newVal);


void __RPC_STUB IFactorTable_put_NComponentName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NComment_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NComment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NComment_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ BSTR newVal);


void __RPC_STUB IFactorTable_put_NComment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NWeight_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NWeight_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][defaultcollelem][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NWeight_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ float newVal);


void __RPC_STUB IFactorTable_put_NWeight_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NValuation_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ VARIANT __RPC_FAR *pfVal);


void __RPC_STUB IFactorTable_get_NValuation_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NValuation_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ VARIANT __RPC_FAR *pfVal);


void __RPC_STUB IFactorTable_put_NValuation_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NAveWeighted_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ VARIANT __RPC_FAR *pfVal);


void __RPC_STUB IFactorTable_get_NAveWeighted_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_WeightSumm_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_WeightSumm_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_ValuationSumm_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pfVal);


void __RPC_STUB IFactorTable_get_ValuationSumm_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_AveWeightedSumm_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pfVal);


void __RPC_STUB IFactorTable_get_AveWeightedSumm_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_AllComponentsIsAssigned_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bRes);


void __RPC_STUB IFactorTable_get_AllComponentsIsAssigned_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_SetWeights_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [optional][in] */ VARIANT StartIdx);


void __RPC_STUB IFactorTable_SetWeights_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_IsWeightsCorrect_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [optional][out][in] */ VARIANT __RPC_FAR *Summ,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_IsWeightsCorrect_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_CalcValues_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [optional][out][in] */ VARIANT __RPC_FAR *fValuationSumm,
    /* [optional][out][in] */ VARIANT __RPC_FAR *fAveWeightedSumm);


void __RPC_STUB IFactorTable_CalcValues_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NSlotType_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ FTSlotTyp __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NSlotType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NSlotValue_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NSlotValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NSlotValue_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ long newVal);


void __RPC_STUB IFactorTable_put_NSlotValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_NAssSlotType_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ FTSlotTyp NewTyp,
    /* [optional][in] */ long K1,
    /* [optional][in] */ long K2);


void __RPC_STUB IFactorTable_NAssSlotType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_NBindSlot_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ IDispatch __RPC_FAR *Object);


void __RPC_STUB IFactorTable_NBindSlot_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_NSlotSelf_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [optional][out] */ long __RPC_FAR *Key,
    /* [optional][out] */ IFactorTable __RPC_FAR *__RPC_FAR *IFT);


void __RPC_STUB IFactorTable_NSlotSelf_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_NSlotGost_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [optional][out] */ long __RPC_FAR *KeyTopic,
    /* [optional][out] */ long __RPC_FAR *KeyGost,
    /* [optional][out] */ IGostTable __RPC_FAR *__RPC_FAR *IGT);


void __RPC_STUB IFactorTable_NSlotGost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_HandsOffObjects_Proxy( 
    IFactorTable __RPC_FAR * This);


void __RPC_STUB IFactorTable_HandsOffObjects_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_GetRefToFTable_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long Key,
    /* [retval][out] */ VARIANT __RPC_FAR *RefSlot);


void __RPC_STUB IFactorTable_GetRefToFTable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_GetRefToGost_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long KeyTopic,
    /* [in] */ long KeyGost,
    /* [retval][out] */ VARIANT __RPC_FAR *RefSlot);


void __RPC_STUB IFactorTable_GetRefToGost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactorTable_NormalizeWeights_Proxy( 
    IFactorTable __RPC_FAR * This);


void __RPC_STUB IFactorTable_NormalizeWeights_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NLocked_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NLocked_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NLocked_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IFactorTable_put_NLocked_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactorTable_get_NLingvoVal_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFactorTable_get_NLingvoVal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactorTable_put_NLingvoVal_Proxy( 
    IFactorTable __RPC_FAR * This,
    /* [in] */ long ZeroIndex,
    /* [in] */ BSTR Val);


void __RPC_STUB IFactorTable_put_NLingvoVal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IFactorTable_INTERFACE_DEFINED__ */


#ifndef __ICollFTables_INTERFACE_DEFINED__
#define __ICollFTables_INTERFACE_DEFINED__

/* interface ICollFTables */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollFTables;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F39A15EC-9890-11D4-900E-00504E02C39D")
    ICollFTables : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR strName,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LongKey( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IFactorTable __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IFactorTable __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Update( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CachedUpdates( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetUpdateMode( 
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ItemNth( 
            /* [in] */ long Number,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_EnumSorting( 
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_EnumSorting( 
            /* [in] */ StdSortSeqv SortSeq) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE KeyNth( 
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE NameNth( 
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE DetectRoot( 
            /* [retval][out] */ long __RPC_FAR *pKey) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollFTablesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollFTables __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollFTables __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollFTables __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollFTables __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollFTables __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ BSTR strName,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LongKey )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NameKey )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IFactorTable __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NameKey )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ IFactorTable __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollFTables __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Update )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CachedUpdates )( 
            ICollFTables __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetUpdateMode )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ItemNth )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnumSorting )( 
            ICollFTables __RPC_FAR * This,
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnumSorting )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ StdSortSeqv SortSeq);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *KeyNth )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NameNth )( 
            ICollFTables __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DetectRoot )( 
            ICollFTables __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pKey);
        
        END_INTERFACE
    } ICollFTablesVtbl;

    interface ICollFTables
    {
        CONST_VTBL struct ICollFTablesVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollFTables_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollFTables_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollFTables_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollFTables_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollFTables_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollFTables_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollFTables_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollFTables_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollFTables_Item(This,Key,ppGT)	\
    (This)->lpVtbl -> Item(This,Key,ppGT)

#define ICollFTables_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollFTables_Add(This,strName,ppGT)	\
    (This)->lpVtbl -> Add(This,strName,ppGT)

#define ICollFTables_Remove(This,KeyOrObj)	\
    (This)->lpVtbl -> Remove(This,KeyOrObj)

#define ICollFTables_get_LongKey(This,Name,plKey)	\
    (This)->lpVtbl -> get_LongKey(This,Name,plKey)

#define ICollFTables_get_NameKey(This,Key,Obj,pbsName)	\
    (This)->lpVtbl -> get_NameKey(This,Key,Obj,pbsName)

#define ICollFTables_put_NameKey(This,Key,Obj,bsNewName)	\
    (This)->lpVtbl -> put_NameKey(This,Key,Obj,bsNewName)

#define ICollFTables_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define ICollFTables_Update(This,KeyOrObj)	\
    (This)->lpVtbl -> Update(This,KeyOrObj)

#define ICollFTables_get_CachedUpdates(This,pbCached)	\
    (This)->lpVtbl -> get_CachedUpdates(This,pbCached)

#define ICollFTables_SetUpdateMode(This,Stm,bMode)	\
    (This)->lpVtbl -> SetUpdateMode(This,Stm,bMode)

#define ICollFTables_ItemNth(This,Number,ppGT)	\
    (This)->lpVtbl -> ItemNth(This,Number,ppGT)

#define ICollFTables_get_EnumSorting(This,pSortSeq)	\
    (This)->lpVtbl -> get_EnumSorting(This,pSortSeq)

#define ICollFTables_put_EnumSorting(This,SortSeq)	\
    (This)->lpVtbl -> put_EnumSorting(This,SortSeq)

#define ICollFTables_KeyNth(This,Number,lKey)	\
    (This)->lpVtbl -> KeyNth(This,Number,lKey)

#define ICollFTables_NameNth(This,Number,bsName)	\
    (This)->lpVtbl -> NameNth(This,Number,bsName)

#define ICollFTables_DetectRoot(This,pKey)	\
    (This)->lpVtbl -> DetectRoot(This,pKey)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFTables_get__NewEnum_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollFTables_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFTables_Item_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *Key,
    /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollFTables_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFTables_get_Count_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollFTables_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFTables_Add_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ BSTR strName,
    /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollFTables_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFTables_Remove_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollFTables_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFTables_get_LongKey_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ long __RPC_FAR *plKey);


void __RPC_STUB ICollFTables_get_LongKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFTables_get_NameKey_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ IFactorTable __RPC_FAR *Obj,
    /* [retval][out] */ BSTR __RPC_FAR *pbsName);


void __RPC_STUB ICollFTables_get_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollFTables_put_NameKey_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ IFactorTable __RPC_FAR *Obj,
    /* [in] */ BSTR bsNewName);


void __RPC_STUB ICollFTables_put_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFTables_Clear_Proxy( 
    ICollFTables __RPC_FAR * This);


void __RPC_STUB ICollFTables_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFTables_Update_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollFTables_Update_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollFTables_get_CachedUpdates_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);


void __RPC_STUB ICollFTables_get_CachedUpdates_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFTables_SetUpdateMode_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stm,
    /* [in] */ VARIANT_BOOL bMode);


void __RPC_STUB ICollFTables_SetUpdateMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFTables_ItemNth_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollFTables_ItemNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollFTables_get_EnumSorting_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);


void __RPC_STUB ICollFTables_get_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICollFTables_put_EnumSorting_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ StdSortSeqv SortSeq);


void __RPC_STUB ICollFTables_put_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFTables_KeyNth_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ long __RPC_FAR *lKey);


void __RPC_STUB ICollFTables_KeyNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFTables_NameNth_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ BSTR __RPC_FAR *bsName);


void __RPC_STUB ICollFTables_NameNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFTables_DetectRoot_Proxy( 
    ICollFTables __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pKey);


void __RPC_STUB ICollFTables_DetectRoot_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollFTables_INTERFACE_DEFINED__ */


#ifndef __IIntCreator_INTERFACE_DEFINED__
#define __IIntCreator_INTERFACE_DEFINED__

/* interface IIntCreator */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IIntCreator;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3D09A443-9A18-11D4-9010-00504E02C39D")
    IIntCreator : public IDispatch
    {
    public:
        virtual /* [hidden][restricted][id] */ HRESULT STDMETHODCALLTYPE NewIFactorTable( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *FactorTable) = 0;
        
        virtual /* [hidden][restricted][id] */ HRESULT STDMETHODCALLTYPE NewIGostTable( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *GostTable) = 0;
        
        virtual /* [hidden][restricted][id] */ HRESULT STDMETHODCALLTYPE NewICollGosts( 
            /* [in] */ BSTR Name,
            /* [in] */ IStorage __RPC_FAR *Stg,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *pCollGosts) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE FetchGostTable( 
            /* [in] */ IStorage __RPC_FAR *StmRoot,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *Path,
            /* [in] */ long OFThrough,
            /* [optional][in] */ VARIANT __RPC_FAR *OFEnd,
            /* [optional][in] */ VARIANT __RPC_FAR *CreCached,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *pGT) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE FetchFactorTable( 
            /* [in] */ IStorage __RPC_FAR *StmRoot,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *Path,
            /* [in] */ long OFThrough,
            /* [optional][in] */ VARIANT __RPC_FAR *OFEnd,
            /* [optional][in] */ VARIANT __RPC_FAR *CreCached,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *pFT) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IIntCreatorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IIntCreator __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IIntCreator __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IIntCreator __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [hidden][restricted][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NewIFactorTable )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *FactorTable);
        
        /* [hidden][restricted][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NewIGostTable )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *GostTable);
        
        /* [hidden][restricted][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NewICollGosts )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [in] */ IStorage __RPC_FAR *Stg,
            /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *pCollGosts);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FetchGostTable )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *StmRoot,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *Path,
            /* [in] */ long OFThrough,
            /* [optional][in] */ VARIANT __RPC_FAR *OFEnd,
            /* [optional][in] */ VARIANT __RPC_FAR *CreCached,
            /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *pGT);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FetchFactorTable )( 
            IIntCreator __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *StmRoot,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *Path,
            /* [in] */ long OFThrough,
            /* [optional][in] */ VARIANT __RPC_FAR *OFEnd,
            /* [optional][in] */ VARIANT __RPC_FAR *CreCached,
            /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *pFT);
        
        END_INTERFACE
    } IIntCreatorVtbl;

    interface IIntCreator
    {
        CONST_VTBL struct IIntCreatorVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IIntCreator_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IIntCreator_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IIntCreator_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IIntCreator_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IIntCreator_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IIntCreator_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IIntCreator_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IIntCreator_NewIFactorTable(This,Name,FactorTable)	\
    (This)->lpVtbl -> NewIFactorTable(This,Name,FactorTable)

#define IIntCreator_NewIGostTable(This,Name,GostTable)	\
    (This)->lpVtbl -> NewIGostTable(This,Name,GostTable)

#define IIntCreator_NewICollGosts(This,Name,Stg,pCollGosts)	\
    (This)->lpVtbl -> NewICollGosts(This,Name,Stg,pCollGosts)

#define IIntCreator_FetchGostTable(This,StmRoot,Path,OFThrough,OFEnd,CreCached,pGT)	\
    (This)->lpVtbl -> FetchGostTable(This,StmRoot,Path,OFThrough,OFEnd,CreCached,pGT)

#define IIntCreator_FetchFactorTable(This,StmRoot,Path,OFThrough,OFEnd,CreCached,pFT)	\
    (This)->lpVtbl -> FetchFactorTable(This,StmRoot,Path,OFThrough,OFEnd,CreCached,pFT)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [hidden][restricted][id] */ HRESULT STDMETHODCALLTYPE IIntCreator_NewIFactorTable_Proxy( 
    IIntCreator __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *FactorTable);


void __RPC_STUB IIntCreator_NewIFactorTable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][restricted][id] */ HRESULT STDMETHODCALLTYPE IIntCreator_NewIGostTable_Proxy( 
    IIntCreator __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *GostTable);


void __RPC_STUB IIntCreator_NewIGostTable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][restricted][id] */ HRESULT STDMETHODCALLTYPE IIntCreator_NewICollGosts_Proxy( 
    IIntCreator __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [in] */ IStorage __RPC_FAR *Stg,
    /* [retval][out] */ ICollGosts __RPC_FAR *__RPC_FAR *pCollGosts);


void __RPC_STUB IIntCreator_NewICollGosts_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IIntCreator_FetchGostTable_Proxy( 
    IIntCreator __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *StmRoot,
    /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *Path,
    /* [in] */ long OFThrough,
    /* [optional][in] */ VARIANT __RPC_FAR *OFEnd,
    /* [optional][in] */ VARIANT __RPC_FAR *CreCached,
    /* [retval][out] */ IGostTable __RPC_FAR *__RPC_FAR *pGT);


void __RPC_STUB IIntCreator_FetchGostTable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IIntCreator_FetchFactorTable_Proxy( 
    IIntCreator __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *StmRoot,
    /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *Path,
    /* [in] */ long OFThrough,
    /* [optional][in] */ VARIANT __RPC_FAR *OFEnd,
    /* [optional][in] */ VARIANT __RPC_FAR *CreCached,
    /* [retval][out] */ IFactorTable __RPC_FAR *__RPC_FAR *pFT);


void __RPC_STUB IIntCreator_FetchFactorTable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IIntCreator_INTERFACE_DEFINED__ */


#ifndef __ICollFactors_INTERFACE_DEFINED__
#define __ICollFactors_INTERFACE_DEFINED__

/* interface ICollFactors */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollFactors;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("2E0F6E03-A76B-11D4-9029-00504E02C39D")
    ICollFactors : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR strName,
            /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LongKey( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollFTables __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NameKey( 
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollFTables __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Update( 
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CachedUpdates( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetUpdateMode( 
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ItemNth( 
            /* [in] */ long Number,
            /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_EnumSorting( 
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_EnumSorting( 
            /* [in] */ StdSortSeqv SortSeq) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE KeyNth( 
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE NameNth( 
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollFactorsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollFactors __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollFactors __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollFactors __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollFactors __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *Key,
            /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollFactors __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ BSTR strName,
            /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LongKey )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ long __RPC_FAR *plKey);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NameKey )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollFTables __RPC_FAR *Obj,
            /* [retval][out] */ BSTR __RPC_FAR *pbsName);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NameKey )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ long Key,
            /* [defaultvalue][optional][in] */ ICollFTables __RPC_FAR *Obj,
            /* [in] */ BSTR bsNewName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollFactors __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Update )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ VARIANT __RPC_FAR *KeyOrObj);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CachedUpdates )( 
            ICollFactors __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetUpdateMode )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stm,
            /* [in] */ VARIANT_BOOL bMode);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ItemNth )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnumSorting )( 
            ICollFactors __RPC_FAR * This,
            /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnumSorting )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ StdSortSeqv SortSeq);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *KeyNth )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ long __RPC_FAR *lKey);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *NameNth )( 
            ICollFactors __RPC_FAR * This,
            /* [in] */ long Number,
            /* [retval][out] */ BSTR __RPC_FAR *bsName);
        
        END_INTERFACE
    } ICollFactorsVtbl;

    interface ICollFactors
    {
        CONST_VTBL struct ICollFactorsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollFactors_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollFactors_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollFactors_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollFactors_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollFactors_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollFactors_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollFactors_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollFactors_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollFactors_Item(This,Key,ppGT)	\
    (This)->lpVtbl -> Item(This,Key,ppGT)

#define ICollFactors_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollFactors_Add(This,strName,ppGT)	\
    (This)->lpVtbl -> Add(This,strName,ppGT)

#define ICollFactors_Remove(This,KeyOrObj)	\
    (This)->lpVtbl -> Remove(This,KeyOrObj)

#define ICollFactors_get_LongKey(This,Name,plKey)	\
    (This)->lpVtbl -> get_LongKey(This,Name,plKey)

#define ICollFactors_get_NameKey(This,Key,Obj,pbsName)	\
    (This)->lpVtbl -> get_NameKey(This,Key,Obj,pbsName)

#define ICollFactors_put_NameKey(This,Key,Obj,bsNewName)	\
    (This)->lpVtbl -> put_NameKey(This,Key,Obj,bsNewName)

#define ICollFactors_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define ICollFactors_Update(This,KeyOrObj)	\
    (This)->lpVtbl -> Update(This,KeyOrObj)

#define ICollFactors_get_CachedUpdates(This,pbCached)	\
    (This)->lpVtbl -> get_CachedUpdates(This,pbCached)

#define ICollFactors_SetUpdateMode(This,Stm,bMode)	\
    (This)->lpVtbl -> SetUpdateMode(This,Stm,bMode)

#define ICollFactors_ItemNth(This,Number,ppGT)	\
    (This)->lpVtbl -> ItemNth(This,Number,ppGT)

#define ICollFactors_get_EnumSorting(This,pSortSeq)	\
    (This)->lpVtbl -> get_EnumSorting(This,pSortSeq)

#define ICollFactors_put_EnumSorting(This,SortSeq)	\
    (This)->lpVtbl -> put_EnumSorting(This,SortSeq)

#define ICollFactors_KeyNth(This,Number,lKey)	\
    (This)->lpVtbl -> KeyNth(This,Number,lKey)

#define ICollFactors_NameNth(This,Number,bsName)	\
    (This)->lpVtbl -> NameNth(This,Number,bsName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFactors_get__NewEnum_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollFactors_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFactors_Item_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *Key,
    /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollFactors_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFactors_get_Count_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollFactors_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFactors_Add_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ BSTR strName,
    /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollFactors_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFactors_Remove_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollFactors_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFactors_get_LongKey_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ long __RPC_FAR *plKey);


void __RPC_STUB ICollFactors_get_LongKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFactors_get_NameKey_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ ICollFTables __RPC_FAR *Obj,
    /* [retval][out] */ BSTR __RPC_FAR *pbsName);


void __RPC_STUB ICollFactors_get_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollFactors_put_NameKey_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ long Key,
    /* [defaultvalue][optional][in] */ ICollFTables __RPC_FAR *Obj,
    /* [in] */ BSTR bsNewName);


void __RPC_STUB ICollFactors_put_NameKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFactors_Clear_Proxy( 
    ICollFactors __RPC_FAR * This);


void __RPC_STUB ICollFactors_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFactors_Update_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ VARIANT __RPC_FAR *KeyOrObj);


void __RPC_STUB ICollFactors_Update_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollFactors_get_CachedUpdates_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbCached);


void __RPC_STUB ICollFactors_get_CachedUpdates_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFactors_SetUpdateMode_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stm,
    /* [in] */ VARIANT_BOOL bMode);


void __RPC_STUB ICollFactors_SetUpdateMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFactors_ItemNth_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ ICollFTables __RPC_FAR *__RPC_FAR *ppGT);


void __RPC_STUB ICollFactors_ItemNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICollFactors_get_EnumSorting_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [retval][out] */ StdSortSeqv __RPC_FAR *pSortSeq);


void __RPC_STUB ICollFactors_get_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICollFactors_put_EnumSorting_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ StdSortSeqv SortSeq);


void __RPC_STUB ICollFactors_put_EnumSorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFactors_KeyNth_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ long __RPC_FAR *lKey);


void __RPC_STUB ICollFactors_KeyNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFactors_NameNth_Proxy( 
    ICollFactors __RPC_FAR * This,
    /* [in] */ long Number,
    /* [retval][out] */ BSTR __RPC_FAR *bsName);


void __RPC_STUB ICollFactors_NameNth_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollFactors_INTERFACE_DEFINED__ */



#ifndef __AWPLUGINLib_LIBRARY_DEFINED__
#define __AWPLUGINLib_LIBRARY_DEFINED__

/* library AWPLUGINLib */
/* [helpstring][version][uuid] */ 




EXTERN_C const IID LIBID_AWPLUGINLib;

EXTERN_C const CLSID CLSID_CollTopics;

#ifdef __cplusplus

class DECLSPEC_UUID("F39A15E6-9890-11D4-900E-00504E02C39D")
CollTopics;
#endif

EXTERN_C const CLSID CLSID_CollFTables;

#ifdef __cplusplus

class DECLSPEC_UUID("F39A15ED-9890-11D4-900E-00504E02C39D")
CollFTables;
#endif

EXTERN_C const CLSID CLSID_IntCreator;

#ifdef __cplusplus

class DECLSPEC_UUID("3D09A444-9A18-11D4-9010-00504E02C39D")
IntCreator;
#endif

EXTERN_C const CLSID CLSID_CollFactors;

#ifdef __cplusplus

class DECLSPEC_UUID("2E0F6E04-A76B-11D4-9029-00504E02C39D")
CollFactors;
#endif
#endif /* __AWPLUGINLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize(     unsigned long __RPC_FAR *, unsigned long            , LPSAFEARRAY __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  LPSAFEARRAY_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, LPSAFEARRAY __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  LPSAFEARRAY_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, LPSAFEARRAY __RPC_FAR * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree(     unsigned long __RPC_FAR *, LPSAFEARRAY __RPC_FAR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long __RPC_FAR *, unsigned long            , VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long __RPC_FAR *, VARIANT __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
