/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Tue Apr 17 07:06:31 2001
 */
/* Compiler settings for G:\WORK\MAG\Alexf\GertNet\GertNet.idl:
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

#ifndef __GertNet_h__
#define __GertNet_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __ILingvoEnum_FWD_DEFINED__
#define __ILingvoEnum_FWD_DEFINED__
typedef interface ILingvoEnum ILingvoEnum;
#endif 	/* __ILingvoEnum_FWD_DEFINED__ */


#ifndef __IFChange_FWD_DEFINED__
#define __IFChange_FWD_DEFINED__
typedef interface IFChange IFChange;
#endif 	/* __IFChange_FWD_DEFINED__ */


#ifndef __IICollFChange_FWD_DEFINED__
#define __IICollFChange_FWD_DEFINED__
typedef interface IICollFChange IICollFChange;
#endif 	/* __IICollFChange_FWD_DEFINED__ */


#ifndef __ISafetyPrecaution_FWD_DEFINED__
#define __ISafetyPrecaution_FWD_DEFINED__
typedef interface ISafetyPrecaution ISafetyPrecaution;
#endif 	/* __ISafetyPrecaution_FWD_DEFINED__ */


#ifndef __ICollSF_FWD_DEFINED__
#define __ICollSF_FWD_DEFINED__
typedef interface ICollSF ICollSF;
#endif 	/* __ICollSF_FWD_DEFINED__ */


#ifndef __ICollLingvo_FWD_DEFINED__
#define __ICollLingvo_FWD_DEFINED__
typedef interface ICollLingvo ICollLingvo;
#endif 	/* __ICollLingvo_FWD_DEFINED__ */


#ifndef __ICompaundFiles_FWD_DEFINED__
#define __ICompaundFiles_FWD_DEFINED__
typedef interface ICompaundFiles ICompaundFiles;
#endif 	/* __ICompaundFiles_FWD_DEFINED__ */


#ifndef __IFactor_FWD_DEFINED__
#define __IFactor_FWD_DEFINED__
typedef interface IFactor IFactor;
#endif 	/* __IFactor_FWD_DEFINED__ */


#ifndef __ILongKey_FWD_DEFINED__
#define __ILongKey_FWD_DEFINED__
typedef interface ILongKey ILongKey;
#endif 	/* __ILongKey_FWD_DEFINED__ */


#ifndef __IClonable_FWD_DEFINED__
#define __IClonable_FWD_DEFINED__
typedef interface IClonable IClonable;
#endif 	/* __IClonable_FWD_DEFINED__ */


#ifndef __IBSTRKey_FWD_DEFINED__
#define __IBSTRKey_FWD_DEFINED__
typedef interface IBSTRKey IBSTRKey;
#endif 	/* __IBSTRKey_FWD_DEFINED__ */


#ifndef __IMGertNet_Debug_FWD_DEFINED__
#define __IMGertNet_Debug_FWD_DEFINED__
typedef interface IMGertNet_Debug IMGertNet_Debug;
#endif 	/* __IMGertNet_Debug_FWD_DEFINED__ */


#ifndef __ICollFac_FWD_DEFINED__
#define __ICollFac_FWD_DEFINED__
typedef interface ICollFac ICollFac;
#endif 	/* __ICollFac_FWD_DEFINED__ */


#ifndef __IMGertNet_FWD_DEFINED__
#define __IMGertNet_FWD_DEFINED__
typedef interface IMGertNet IMGertNet;
#endif 	/* __IMGertNet_FWD_DEFINED__ */


#ifndef __IFactorAssign_FWD_DEFINED__
#define __IFactorAssign_FWD_DEFINED__
typedef interface IFactorAssign IFactorAssign;
#endif 	/* __IFactorAssign_FWD_DEFINED__ */


#ifndef __IFactorAssign_FWD_DEFINED__
#define __IFactorAssign_FWD_DEFINED__
typedef interface IFactorAssign IFactorAssign;
#endif 	/* __IFactorAssign_FWD_DEFINED__ */


#ifndef ___IGertNetEvents_FWD_DEFINED__
#define ___IGertNetEvents_FWD_DEFINED__
typedef interface _IGertNetEvents _IGertNetEvents;
#endif 	/* ___IGertNetEvents_FWD_DEFINED__ */


#ifndef ___IFacEvents_FWD_DEFINED__
#define ___IFacEvents_FWD_DEFINED__
typedef interface _IFacEvents _IFacEvents;
#endif 	/* ___IFacEvents_FWD_DEFINED__ */


#ifndef __ICatRegistrar_FWD_DEFINED__
#define __ICatRegistrar_FWD_DEFINED__
typedef interface ICatRegistrar ICatRegistrar;
#endif 	/* __ICatRegistrar_FWD_DEFINED__ */


#ifndef __LingvoEnum_FWD_DEFINED__
#define __LingvoEnum_FWD_DEFINED__

#ifdef __cplusplus
typedef class LingvoEnum LingvoEnum;
#else
typedef struct LingvoEnum LingvoEnum;
#endif /* __cplusplus */

#endif 	/* __LingvoEnum_FWD_DEFINED__ */


#ifndef __CollLingvo_FWD_DEFINED__
#define __CollLingvo_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollLingvo CollLingvo;
#else
typedef struct CollLingvo CollLingvo;
#endif /* __cplusplus */

#endif 	/* __CollLingvo_FWD_DEFINED__ */


#ifndef __CompaundFiles_FWD_DEFINED__
#define __CompaundFiles_FWD_DEFINED__

#ifdef __cplusplus
typedef class CompaundFiles CompaundFiles;
#else
typedef struct CompaundFiles CompaundFiles;
#endif /* __cplusplus */

#endif 	/* __CompaundFiles_FWD_DEFINED__ */


#ifndef __Factor_FWD_DEFINED__
#define __Factor_FWD_DEFINED__

#ifdef __cplusplus
typedef class Factor Factor;
#else
typedef struct Factor Factor;
#endif /* __cplusplus */

#endif 	/* __Factor_FWD_DEFINED__ */


#ifndef __CollFac_FWD_DEFINED__
#define __CollFac_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollFac CollFac;
#else
typedef struct CollFac CollFac;
#endif /* __cplusplus */

#endif 	/* __CollFac_FWD_DEFINED__ */


#ifndef __MGertNet_FWD_DEFINED__
#define __MGertNet_FWD_DEFINED__

#ifdef __cplusplus
typedef class MGertNet MGertNet;
#else
typedef struct MGertNet MGertNet;
#endif /* __cplusplus */

#endif 	/* __MGertNet_FWD_DEFINED__ */


#ifndef __SafetyPrecaution_FWD_DEFINED__
#define __SafetyPrecaution_FWD_DEFINED__

#ifdef __cplusplus
typedef class SafetyPrecaution SafetyPrecaution;
#else
typedef struct SafetyPrecaution SafetyPrecaution;
#endif /* __cplusplus */

#endif 	/* __SafetyPrecaution_FWD_DEFINED__ */


#ifndef __FChange_FWD_DEFINED__
#define __FChange_FWD_DEFINED__

#ifdef __cplusplus
typedef class FChange FChange;
#else
typedef struct FChange FChange;
#endif /* __cplusplus */

#endif 	/* __FChange_FWD_DEFINED__ */


#ifndef __ICollFChange_FWD_DEFINED__
#define __ICollFChange_FWD_DEFINED__

#ifdef __cplusplus
typedef class ICollFChange ICollFChange;
#else
typedef struct ICollFChange ICollFChange;
#endif /* __cplusplus */

#endif 	/* __ICollFChange_FWD_DEFINED__ */


#ifndef __CollSF_FWD_DEFINED__
#define __CollSF_FWD_DEFINED__

#ifdef __cplusplus
typedef class CollSF CollSF;
#else
typedef struct CollSF CollSF;
#endif /* __cplusplus */

#endif 	/* __CollSF_FWD_DEFINED__ */


#ifndef __CatRegistrar_FWD_DEFINED__
#define __CatRegistrar_FWD_DEFINED__

#ifdef __cplusplus
typedef class CatRegistrar CatRegistrar;
#else
typedef struct CatRegistrar CatRegistrar;
#endif /* __cplusplus */

#endif 	/* __CatRegistrar_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_GertNet_0000 */
/* [local] */ 

typedef /* [public][uuid][v1_enum] */ 
enum EnumCompConsts_tag
    {	STGM_DIRECT1	= 0L,
	STGM_TRANSACTED1	= 0x10000L,
	STGM_SIMPLE1	= 0x8000000L,
	STGM_READ1	= 0L,
	STGM_WRITE1	= 0x1L,
	STGM_READWRITE1	= 0x2L,
	STGM_SHARE_DENY_NONE1	= 0x40L,
	STGM_SHARE_DENY_READ1	= 0x30L,
	STGM_SHARE_DENY_WRITE1	= 0x20L,
	STGM_SHARE_EXCLUSIVE1	= 0x10L,
	STGM_PRIORITY1	= 0x40000L,
	STGM_DELETEONRELEASE1	= 0x4000000L,
	STGM_NOSCRATCH1	= 0x100000L,
	STGM_CREATE1	= 0x1000L,
	STGM_CONVERT1	= 0x20000L,
	STGM_FAILIFTHERE1	= 0L,
	STGM_NOSNAPSHOT1	= 0x200000L
    }	EnumCompConsts;

typedef /* [public][uuid][v1_enum] */ 
enum STGC1_tag
    {	STGC_DEFAULT1	= STGC_DEFAULT,
	STGC_OVERWRITE1	= STGC_OVERWRITE,
	STGC_ONLYIFCURRENT1	= STGC_ONLYIFCURRENT,
	STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE1	= STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE,
	STGC_CONSOLIDATE1	= STGC_CONSOLIDATE
    }	STGC1;

typedef /* [uuid][public] */ 
enum TrustLevel_tag
    {	TL_Low	= 0,
	TL_Normal	= 1,
	TL_High	= 2
    }	TrustLevel;

typedef /* [uuid][public] */ 
enum ModelType_tag
    {	MT_Imitate	= 0,
	MT_ImitateIndistinct	= MT_Imitate + 1,
	MT_Analytical	= MT_ImitateIndistinct + 1,
	MT_AnalyticalIndistinct	= MT_Analytical + 1,
	MT_Analytical2	= MT_AnalyticalIndistinct + 1
    }	ModelType;

typedef /* [uuid][public] */ 
enum OptType_tag
    {	OT_Quick	= 0,
	OT_Quick2	= OT_Quick + 1,
	OT_Full	= OT_Quick2 + 1,
	OT_Integer	= OT_Full + 1,
	OT_IntegerAdv	= OT_Integer + 1
    }	OptType;

typedef /* [uuid][public] */ 
enum TDRFlags_tag
    {	TDR_Stg	= 1,
	TDR_Strm	= 2
    }	TDRFlags;

typedef /* [uuid][public] */ 
enum TStatFlag_tag
    {	TSF_No	= 1,
	TSF_Yes	= 2,
	TSF_Full	= 3
    }	TStatFlag;

typedef /* [uuid][v1_enum][public] */ 
enum NotifyMsg_tag
    {	NM_OnEndOfWork	= 0x400 + 10,
	NM_OnStepOfWork	= 0x400 + 11,
	NM_OnCancel	= 0x400 + 12,
	NM_OnEndOfWork2	= 0x400 + 13,
	NM_OnStepOfWork2	= 0x400 + 14,
	NM_OnEndOfWork3	= 0x400 + 15,
	NM_OnStepOfWork3	= 0x400 + 16,
	NM_OnCancel3	= 0x400 + 17,
	NM_OnErrorCalc	= 0x400 + 18
    }	NotifyMsg;

typedef /* [uuid][public] */ 
enum OptTask_tag
    {	OT_FixMoney_MinQ	= 0,
	OT_FixQ_MinMoney	= OT_FixMoney_MinQ + 1
    }	OptTask;

typedef /* [uuid][public] */ 
enum CollSFSorting_tag
    {	CSFS_None	= 0,
	CSFS_Price	= CSFS_None + 1,
	CSFS_DeltaQ	= CSFS_Price + 1,
	CSFS_PriceDescending	= CSFS_DeltaQ + 1,
	CSFS_DeltaQDescending	= CSFS_PriceDescending + 1,
	CSFS_KDescending	= CSFS_DeltaQDescending + 1
    }	CollSFSorting;

typedef /* [uuid][v1_enum][public] */ 
enum ThrdPriority_tag
    {	TP_ABOVE_NORMAL	= 0,
	TP_BELOW_NORMAL	= TP_ABOVE_NORMAL + 1,
	TP_HIGHEST	= TP_BELOW_NORMAL + 1,
	TP_LOWEST	= TP_HIGHEST + 1,
	TP_NORMAL	= TP_LOWEST + 1
    }	ThrdPriority;

typedef /* [uuid][public] */ 
enum TrustChange_tag
    {	TR_NoChange	= 0,
	TR_MinusOne	= 1,
	TR_MinusTwo	= 2,
	TR_PlusOne	= 3,
	TR_PlusTwo	= 4,
	TR_SetLow	= 5,
	TR_SetNormal	= 6,
	TR_SetHigh	= 7
    }	TrustChange;

typedef /* [uuid][public] */ 
enum PrpbDistrTyp_tag
    {	PDT_Uniform	= 0,
	PDT_Cauchy	= 1
    }	PrpbDistrTyp;

typedef /* [uuid][public] */ 
enum NumberFormat_tag
    {	NF_Normal	= 0,
	NF_Scientific	= 1
    }	NumberFormat;

typedef /* [uuid][public] */ 
enum NFormatTyp_tag
    {	NFT_Pr	= 0,
	NFT_Other	= 1
    }	NFormatTyp;

struct  Tag_TypANodesStat
    {
    double dAvgP;
    double dSummP;
    double dMinVal;
    double dMaxVal;
    CY cySz;
    double dAvg;
    double dDx;
    };
typedef struct Tag_TypANodesStat TypANodesStat;



extern RPC_IF_HANDLE __MIDL_itf_GertNet_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_GertNet_0000_v0_0_s_ifspec;

#ifndef __ILingvoEnum_INTERFACE_DEFINED__
#define __ILingvoEnum_INTERFACE_DEFINED__

/* interface ILingvoEnum */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILingvoEnum;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("86822110-0876-11D4-8EE4-00504E02C39D")
    ILingvoEnum : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Quality( 
            /* [in] */ long lOrder,
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Quality( 
            /* [in] */ long lOrder,
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Mark( 
            /* [in] */ long lOrder,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Mark( 
            /* [in] */ long lOrder,
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE MarkForValue( 
            /* [in] */ double dVal,
            /* [retval][out] */ BSTR __RPC_FAR *pbsMark) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE UpdateFrom( 
            /* [in] */ ILingvoEnum __RPC_FAR *pLEnum) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RoundS( 
            /* [in] */ double Val,
            /* [retval][out] */ double __RPC_FAR *dV) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ValueIdx( 
            /* [in] */ double Val,
            /* [retval][out] */ short __RPC_FAR *shIdx) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILingvoEnumVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ILingvoEnum __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ILingvoEnum __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ILingvoEnum __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Quality )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ long lOrder,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Quality )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ long lOrder,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Mark )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ long lOrder,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Mark )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ long lOrder,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ILingvoEnum __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *MarkForValue )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ double dVal,
            /* [retval][out] */ BSTR __RPC_FAR *pbsMark);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ID )( 
            ILingvoEnum __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsDirty )( 
            ILingvoEnum __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UpdateFrom )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ ILingvoEnum __RPC_FAR *pLEnum);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RoundS )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ double Val,
            /* [retval][out] */ double __RPC_FAR *dV);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ValueIdx )( 
            ILingvoEnum __RPC_FAR * This,
            /* [in] */ double Val,
            /* [retval][out] */ short __RPC_FAR *shIdx);
        
        END_INTERFACE
    } ILingvoEnumVtbl;

    interface ILingvoEnum
    {
        CONST_VTBL struct ILingvoEnumVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILingvoEnum_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILingvoEnum_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILingvoEnum_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILingvoEnum_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILingvoEnum_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILingvoEnum_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILingvoEnum_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILingvoEnum_get_Quality(This,lOrder,pVal)	\
    (This)->lpVtbl -> get_Quality(This,lOrder,pVal)

#define ILingvoEnum_put_Quality(This,lOrder,newVal)	\
    (This)->lpVtbl -> put_Quality(This,lOrder,newVal)

#define ILingvoEnum_get_Mark(This,lOrder,pVal)	\
    (This)->lpVtbl -> get_Mark(This,lOrder,pVal)

#define ILingvoEnum_put_Mark(This,lOrder,newVal)	\
    (This)->lpVtbl -> put_Mark(This,lOrder,newVal)

#define ILingvoEnum_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ILingvoEnum_MarkForValue(This,dVal,pbsMark)	\
    (This)->lpVtbl -> MarkForValue(This,dVal,pbsMark)

#define ILingvoEnum_get_ID(This,pVal)	\
    (This)->lpVtbl -> get_ID(This,pVal)

#define ILingvoEnum_get_IsDirty(This,pVal)	\
    (This)->lpVtbl -> get_IsDirty(This,pVal)

#define ILingvoEnum_UpdateFrom(This,pLEnum)	\
    (This)->lpVtbl -> UpdateFrom(This,pLEnum)

#define ILingvoEnum_RoundS(This,Val,dV)	\
    (This)->lpVtbl -> RoundS(This,Val,dV)

#define ILingvoEnum_ValueIdx(This,Val,shIdx)	\
    (This)->lpVtbl -> ValueIdx(This,Val,shIdx)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_get_Quality_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ long lOrder,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB ILingvoEnum_get_Quality_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_put_Quality_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ long lOrder,
    /* [in] */ double newVal);


void __RPC_STUB ILingvoEnum_put_Quality_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_get_Mark_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ long lOrder,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB ILingvoEnum_get_Mark_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_put_Mark_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ long lOrder,
    /* [in] */ BSTR newVal);


void __RPC_STUB ILingvoEnum_put_Mark_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_get_Count_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ILingvoEnum_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_MarkForValue_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ double dVal,
    /* [retval][out] */ BSTR __RPC_FAR *pbsMark);


void __RPC_STUB ILingvoEnum_MarkForValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_get_ID_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB ILingvoEnum_get_ID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_get_IsDirty_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ILingvoEnum_get_IsDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_UpdateFrom_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ ILingvoEnum __RPC_FAR *pLEnum);


void __RPC_STUB ILingvoEnum_UpdateFrom_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_RoundS_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ double Val,
    /* [retval][out] */ double __RPC_FAR *dV);


void __RPC_STUB ILingvoEnum_RoundS_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILingvoEnum_ValueIdx_Proxy( 
    ILingvoEnum __RPC_FAR * This,
    /* [in] */ double Val,
    /* [retval][out] */ short __RPC_FAR *shIdx);


void __RPC_STUB ILingvoEnum_ValueIdx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILingvoEnum_INTERFACE_DEFINED__ */


#ifndef __IFChange_INTERFACE_DEFINED__
#define __IFChange_INTERFACE_DEFINED__

/* interface IFChange */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IFChange;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0482EBD4-15F5-11D4-8EFD-00504E02C39D")
    IFChange : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NameFactor( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NameFactor( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TCh( 
            /* [retval][out] */ TrustChange __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TCh( 
            /* [in] */ TrustChange newVal) = 0;
        
        virtual /* [defaultcollelem][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Delta( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [defaultcollelem][helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Delta( 
            /* [in] */ short newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IFChangeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IFChange __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IFChange __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IFChange __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IFChange __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IFChange __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IFChange __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IFChange __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NameFactor )( 
            IFChange __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NameFactor )( 
            IFChange __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ID )( 
            IFChange __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TCh )( 
            IFChange __RPC_FAR * This,
            /* [retval][out] */ TrustChange __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TCh )( 
            IFChange __RPC_FAR * This,
            /* [in] */ TrustChange newVal);
        
        /* [defaultcollelem][helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Delta )( 
            IFChange __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [defaultcollelem][helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Delta )( 
            IFChange __RPC_FAR * This,
            /* [in] */ short newVal);
        
        END_INTERFACE
    } IFChangeVtbl;

    interface IFChange
    {
        CONST_VTBL struct IFChangeVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IFChange_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IFChange_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IFChange_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IFChange_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IFChange_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IFChange_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IFChange_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IFChange_get_NameFactor(This,pVal)	\
    (This)->lpVtbl -> get_NameFactor(This,pVal)

#define IFChange_put_NameFactor(This,newVal)	\
    (This)->lpVtbl -> put_NameFactor(This,newVal)

#define IFChange_get_ID(This,pVal)	\
    (This)->lpVtbl -> get_ID(This,pVal)

#define IFChange_get_TCh(This,pVal)	\
    (This)->lpVtbl -> get_TCh(This,pVal)

#define IFChange_put_TCh(This,newVal)	\
    (This)->lpVtbl -> put_TCh(This,newVal)

#define IFChange_get_Delta(This,pVal)	\
    (This)->lpVtbl -> get_Delta(This,pVal)

#define IFChange_put_Delta(This,newVal)	\
    (This)->lpVtbl -> put_Delta(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFChange_get_NameFactor_Proxy( 
    IFChange __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFChange_get_NameFactor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFChange_put_NameFactor_Proxy( 
    IFChange __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IFChange_put_NameFactor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFChange_get_ID_Proxy( 
    IFChange __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB IFChange_get_ID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFChange_get_TCh_Proxy( 
    IFChange __RPC_FAR * This,
    /* [retval][out] */ TrustChange __RPC_FAR *pVal);


void __RPC_STUB IFChange_get_TCh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFChange_put_TCh_Proxy( 
    IFChange __RPC_FAR * This,
    /* [in] */ TrustChange newVal);


void __RPC_STUB IFChange_put_TCh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [defaultcollelem][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFChange_get_Delta_Proxy( 
    IFChange __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IFChange_get_Delta_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [defaultcollelem][helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFChange_put_Delta_Proxy( 
    IFChange __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IFChange_put_Delta_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IFChange_INTERFACE_DEFINED__ */


#ifndef __IICollFChange_INTERFACE_DEFINED__
#define __IICollFChange_INTERFACE_DEFINED__

/* interface IICollFChange */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IICollFChange;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0482EBD6-15F5-11D4-8EFD-00504E02C39D")
    IICollFChange : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ long index,
            /* [retval][out] */ IFChange __RPC_FAR *__RPC_FAR *pvar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ IFChange __RPC_FAR *pI,
            /* [defaultvalue][optional][in] */ long index,
            /* [optional][retval][out] */ long __RPC_FAR *indexAss) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ long index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LookNextIndex( 
            /* [retval][out] */ long __RPC_FAR *plIndex) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IICollFChangeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IICollFChange __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IICollFChange __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IICollFChange __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            IICollFChange __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ long index,
            /* [retval][out] */ IFChange __RPC_FAR *__RPC_FAR *pvar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            IICollFChange __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ IFChange __RPC_FAR *pI,
            /* [defaultvalue][optional][in] */ long index,
            /* [optional][retval][out] */ long __RPC_FAR *indexAss);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            IICollFChange __RPC_FAR * This,
            /* [in] */ long index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LookNextIndex )( 
            IICollFChange __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *plIndex);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            IICollFChange __RPC_FAR * This);
        
        END_INTERFACE
    } IICollFChangeVtbl;

    interface IICollFChange
    {
        CONST_VTBL struct IICollFChangeVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IICollFChange_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IICollFChange_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IICollFChange_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IICollFChange_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IICollFChange_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IICollFChange_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IICollFChange_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IICollFChange_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define IICollFChange_Item(This,index,pvar)	\
    (This)->lpVtbl -> Item(This,index,pvar)

#define IICollFChange_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define IICollFChange_Add(This,pI,index,indexAss)	\
    (This)->lpVtbl -> Add(This,pI,index,indexAss)

#define IICollFChange_Remove(This,index)	\
    (This)->lpVtbl -> Remove(This,index)

#define IICollFChange_LookNextIndex(This,plIndex)	\
    (This)->lpVtbl -> LookNextIndex(This,plIndex)

#define IICollFChange_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE IICollFChange_get__NewEnum_Proxy( 
    IICollFChange __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB IICollFChange_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IICollFChange_Item_Proxy( 
    IICollFChange __RPC_FAR * This,
    /* [in] */ long index,
    /* [retval][out] */ IFChange __RPC_FAR *__RPC_FAR *pvar);


void __RPC_STUB IICollFChange_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IICollFChange_get_Count_Proxy( 
    IICollFChange __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IICollFChange_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IICollFChange_Add_Proxy( 
    IICollFChange __RPC_FAR * This,
    /* [in] */ IFChange __RPC_FAR *pI,
    /* [defaultvalue][optional][in] */ long index,
    /* [optional][retval][out] */ long __RPC_FAR *indexAss);


void __RPC_STUB IICollFChange_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IICollFChange_Remove_Proxy( 
    IICollFChange __RPC_FAR * This,
    /* [in] */ long index);


void __RPC_STUB IICollFChange_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IICollFChange_LookNextIndex_Proxy( 
    IICollFChange __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *plIndex);


void __RPC_STUB IICollFChange_LookNextIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IICollFChange_Clear_Proxy( 
    IICollFChange __RPC_FAR * This);


void __RPC_STUB IICollFChange_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IICollFChange_INTERFACE_DEFINED__ */


#ifndef __ISafetyPrecaution_INTERFACE_DEFINED__
#define __ISafetyPrecaution_INTERFACE_DEFINED__

/* interface ISafetyPrecaution */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ISafetyPrecaution;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F4E48353-15DF-11D4-8EFC-00504E02C39D")
    ISafetyPrecaution : public IDispatch
    {
    public:
        virtual /* [defaultcollelem][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [defaultcollelem][helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Cost( 
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Cost( 
            /* [in] */ CURRENCY newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DeltaQ( 
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DeltaQ( 
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_FChanges( 
            /* [retval][out] */ IICollFChange __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NonCompatibles( 
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NonCompatibles( 
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *parrVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Profit( 
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Profit( 
            /* [in] */ CURRENCY newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Ke( 
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Ke( 
            /* [in] */ double newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ISafetyPrecautionVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ISafetyPrecaution __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ISafetyPrecaution __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [defaultcollelem][helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [defaultcollelem][helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Cost )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Cost )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ CURRENCY newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DeltaQ )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DeltaQ )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ID )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_FChanges )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ IICollFChange __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NonCompatibles )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NonCompatibles )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *parrVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsDirty )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Profit )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Profit )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ CURRENCY newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Ke )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Ke )( 
            ISafetyPrecaution __RPC_FAR * This,
            /* [in] */ double newVal);
        
        END_INTERFACE
    } ISafetyPrecautionVtbl;

    interface ISafetyPrecaution
    {
        CONST_VTBL struct ISafetyPrecautionVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISafetyPrecaution_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ISafetyPrecaution_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ISafetyPrecaution_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ISafetyPrecaution_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ISafetyPrecaution_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ISafetyPrecaution_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ISafetyPrecaution_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ISafetyPrecaution_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define ISafetyPrecaution_put_Name(This,newVal)	\
    (This)->lpVtbl -> put_Name(This,newVal)

#define ISafetyPrecaution_get_Cost(This,pVal)	\
    (This)->lpVtbl -> get_Cost(This,pVal)

#define ISafetyPrecaution_put_Cost(This,newVal)	\
    (This)->lpVtbl -> put_Cost(This,newVal)

#define ISafetyPrecaution_get_DeltaQ(This,pVal)	\
    (This)->lpVtbl -> get_DeltaQ(This,pVal)

#define ISafetyPrecaution_put_DeltaQ(This,newVal)	\
    (This)->lpVtbl -> put_DeltaQ(This,newVal)

#define ISafetyPrecaution_get_ID(This,pVal)	\
    (This)->lpVtbl -> get_ID(This,pVal)

#define ISafetyPrecaution_get_FChanges(This,pVal)	\
    (This)->lpVtbl -> get_FChanges(This,pVal)

#define ISafetyPrecaution_get_NonCompatibles(This,pparrVal)	\
    (This)->lpVtbl -> get_NonCompatibles(This,pparrVal)

#define ISafetyPrecaution_put_NonCompatibles(This,parrVal)	\
    (This)->lpVtbl -> put_NonCompatibles(This,parrVal)

#define ISafetyPrecaution_get_IsDirty(This,pVal)	\
    (This)->lpVtbl -> get_IsDirty(This,pVal)

#define ISafetyPrecaution_get_Profit(This,pVal)	\
    (This)->lpVtbl -> get_Profit(This,pVal)

#define ISafetyPrecaution_put_Profit(This,newVal)	\
    (This)->lpVtbl -> put_Profit(This,newVal)

#define ISafetyPrecaution_get_Ke(This,pVal)	\
    (This)->lpVtbl -> get_Ke(This,pVal)

#define ISafetyPrecaution_put_Ke(This,newVal)	\
    (This)->lpVtbl -> put_Ke(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [defaultcollelem][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_Name_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [defaultcollelem][helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_put_Name_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ISafetyPrecaution_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_Cost_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ CURRENCY __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_Cost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_put_Cost_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [in] */ CURRENCY newVal);


void __RPC_STUB ISafetyPrecaution_put_Cost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_DeltaQ_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_DeltaQ_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_put_DeltaQ_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [in] */ double newVal);


void __RPC_STUB ISafetyPrecaution_put_DeltaQ_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_ID_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_ID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_FChanges_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ IICollFChange __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_FChanges_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_NonCompatibles_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrVal);


void __RPC_STUB ISafetyPrecaution_get_NonCompatibles_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_put_NonCompatibles_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *parrVal);


void __RPC_STUB ISafetyPrecaution_put_NonCompatibles_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_IsDirty_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_IsDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_Profit_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ CURRENCY __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_Profit_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_put_Profit_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [in] */ CURRENCY newVal);


void __RPC_STUB ISafetyPrecaution_put_Profit_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_get_Ke_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB ISafetyPrecaution_get_Ke_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ISafetyPrecaution_put_Ke_Proxy( 
    ISafetyPrecaution __RPC_FAR * This,
    /* [in] */ double newVal);


void __RPC_STUB ISafetyPrecaution_put_Ke_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ISafetyPrecaution_INTERFACE_DEFINED__ */


#ifndef __ICollSF_INTERFACE_DEFINED__
#define __ICollSF_INTERFACE_DEFINED__

/* interface ICollSF */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollSF;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("88B55013-176A-11D4-8EFE-00504E02C39D")
    ICollSF : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Sorting( 
            /* [retval][out] */ CollSFSorting __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Sorting( 
            /* [in] */ CollSFSorting newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Profit( 
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Profit( 
            /* [in] */ CURRENCY newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Ke( 
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Ke( 
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Cost( 
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Cost( 
            /* [in] */ CURRENCY newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DeltaQ( 
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DeltaQ( 
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CheckCompatible( 
            /* [defaultvalue][optional][out] */ ICollSF __RPC_FAR *__RPC_FAR *pSFNone,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bResult) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ long index,
            /* [retval][out] */ ISafetyPrecaution __RPC_FAR *__RPC_FAR *pvar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ ISafetyPrecaution __RPC_FAR *pI,
            /* [defaultvalue][optional][in] */ long index,
            /* [optional][retval][out] */ long __RPC_FAR *indexAss) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ long index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LookNextIndex( 
            /* [retval][out] */ long __RPC_FAR *plIndex) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsValidDelta( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InvalidateDelta( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollSFVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollSF __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollSF __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollSF __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Sorting )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ CollSFSorting __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Sorting )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ CollSFSorting newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsDirty )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Profit )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Profit )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ CURRENCY newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Ke )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Ke )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Cost )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Cost )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ CURRENCY newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DeltaQ )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DeltaQ )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ double newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CheckCompatible )( 
            ICollSF __RPC_FAR * This,
            /* [defaultvalue][optional][out] */ ICollSF __RPC_FAR *__RPC_FAR *pSFNone,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bResult);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ long index,
            /* [retval][out] */ ISafetyPrecaution __RPC_FAR *__RPC_FAR *pvar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ ISafetyPrecaution __RPC_FAR *pI,
            /* [defaultvalue][optional][in] */ long index,
            /* [optional][retval][out] */ long __RPC_FAR *indexAss);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollSF __RPC_FAR * This,
            /* [in] */ long index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LookNextIndex )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *plIndex);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollSF __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsValidDelta )( 
            ICollSF __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *InvalidateDelta )( 
            ICollSF __RPC_FAR * This);
        
        END_INTERFACE
    } ICollSFVtbl;

    interface ICollSF
    {
        CONST_VTBL struct ICollSFVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollSF_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollSF_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollSF_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollSF_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollSF_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollSF_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollSF_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollSF_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollSF_get_Sorting(This,pVal)	\
    (This)->lpVtbl -> get_Sorting(This,pVal)

#define ICollSF_put_Sorting(This,newVal)	\
    (This)->lpVtbl -> put_Sorting(This,newVal)

#define ICollSF_get_IsDirty(This,pVal)	\
    (This)->lpVtbl -> get_IsDirty(This,pVal)

#define ICollSF_get_Profit(This,pVal)	\
    (This)->lpVtbl -> get_Profit(This,pVal)

#define ICollSF_put_Profit(This,newVal)	\
    (This)->lpVtbl -> put_Profit(This,newVal)

#define ICollSF_get_Ke(This,pVal)	\
    (This)->lpVtbl -> get_Ke(This,pVal)

#define ICollSF_put_Ke(This,newVal)	\
    (This)->lpVtbl -> put_Ke(This,newVal)

#define ICollSF_get_Cost(This,pVal)	\
    (This)->lpVtbl -> get_Cost(This,pVal)

#define ICollSF_put_Cost(This,newVal)	\
    (This)->lpVtbl -> put_Cost(This,newVal)

#define ICollSF_get_DeltaQ(This,pVal)	\
    (This)->lpVtbl -> get_DeltaQ(This,pVal)

#define ICollSF_put_DeltaQ(This,newVal)	\
    (This)->lpVtbl -> put_DeltaQ(This,newVal)

#define ICollSF_CheckCompatible(This,pSFNone,bResult)	\
    (This)->lpVtbl -> CheckCompatible(This,pSFNone,bResult)

#define ICollSF_Item(This,index,pvar)	\
    (This)->lpVtbl -> Item(This,index,pvar)

#define ICollSF_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollSF_Add(This,pI,index,indexAss)	\
    (This)->lpVtbl -> Add(This,pI,index,indexAss)

#define ICollSF_Remove(This,index)	\
    (This)->lpVtbl -> Remove(This,index)

#define ICollSF_LookNextIndex(This,plIndex)	\
    (This)->lpVtbl -> LookNextIndex(This,plIndex)

#define ICollSF_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define ICollSF_get_IsValidDelta(This,pVal)	\
    (This)->lpVtbl -> get_IsValidDelta(This,pVal)

#define ICollSF_InvalidateDelta(This)	\
    (This)->lpVtbl -> InvalidateDelta(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get__NewEnum_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_Sorting_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ CollSFSorting __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_Sorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollSF_put_Sorting_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ CollSFSorting newVal);


void __RPC_STUB ICollSF_put_Sorting_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_IsDirty_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_IsDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_Profit_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ CURRENCY __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_Profit_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollSF_put_Profit_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ CURRENCY newVal);


void __RPC_STUB ICollSF_put_Profit_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_Ke_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_Ke_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollSF_put_Ke_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ double newVal);


void __RPC_STUB ICollSF_put_Ke_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_Cost_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ CURRENCY __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_Cost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollSF_put_Cost_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ CURRENCY newVal);


void __RPC_STUB ICollSF_put_Cost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_DeltaQ_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_DeltaQ_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICollSF_put_DeltaQ_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ double newVal);


void __RPC_STUB ICollSF_put_DeltaQ_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollSF_CheckCompatible_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [defaultvalue][optional][out] */ ICollSF __RPC_FAR *__RPC_FAR *pSFNone,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bResult);


void __RPC_STUB ICollSF_CheckCompatible_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollSF_Item_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ long index,
    /* [retval][out] */ ISafetyPrecaution __RPC_FAR *__RPC_FAR *pvar);


void __RPC_STUB ICollSF_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_Count_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollSF_Add_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ ISafetyPrecaution __RPC_FAR *pI,
    /* [defaultvalue][optional][in] */ long index,
    /* [optional][retval][out] */ long __RPC_FAR *indexAss);


void __RPC_STUB ICollSF_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollSF_Remove_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [in] */ long index);


void __RPC_STUB ICollSF_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollSF_LookNextIndex_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *plIndex);


void __RPC_STUB ICollSF_LookNextIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollSF_Clear_Proxy( 
    ICollSF __RPC_FAR * This);


void __RPC_STUB ICollSF_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollSF_get_IsValidDelta_Proxy( 
    ICollSF __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICollSF_get_IsValidDelta_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollSF_InvalidateDelta_Proxy( 
    ICollSF __RPC_FAR * This);


void __RPC_STUB ICollSF_InvalidateDelta_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollSF_INTERFACE_DEFINED__ */


#ifndef __ICollLingvo_INTERFACE_DEFINED__
#define __ICollLingvo_INTERFACE_DEFINED__

/* interface ICollLingvo */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollLingvo;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("86822115-0876-11D4-8EE4-00504E02C39D")
    ICollLingvo : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ long index,
            /* [retval][out] */ ILingvoEnum __RPC_FAR *__RPC_FAR *pvar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ ILingvoEnum __RPC_FAR *pI,
            /* [defaultvalue][optional][in] */ long index,
            /* [optional][retval][out] */ long __RPC_FAR *indexAss) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ long index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LookNextIndex( 
            /* [retval][out] */ long __RPC_FAR *plIndex) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE UpdateFrom( 
            /* [in] */ ICollLingvo __RPC_FAR *pIColl) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollLingvoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollLingvo __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollLingvo __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollLingvo __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollLingvo __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ long index,
            /* [retval][out] */ ILingvoEnum __RPC_FAR *__RPC_FAR *pvar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollLingvo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ ILingvoEnum __RPC_FAR *pI,
            /* [defaultvalue][optional][in] */ long index,
            /* [optional][retval][out] */ long __RPC_FAR *indexAss);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ long index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LookNextIndex )( 
            ICollLingvo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *plIndex);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollLingvo __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsDirty )( 
            ICollLingvo __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UpdateFrom )( 
            ICollLingvo __RPC_FAR * This,
            /* [in] */ ICollLingvo __RPC_FAR *pIColl);
        
        END_INTERFACE
    } ICollLingvoVtbl;

    interface ICollLingvo
    {
        CONST_VTBL struct ICollLingvoVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollLingvo_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollLingvo_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollLingvo_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollLingvo_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollLingvo_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollLingvo_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollLingvo_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollLingvo_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollLingvo_Item(This,index,pvar)	\
    (This)->lpVtbl -> Item(This,index,pvar)

#define ICollLingvo_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollLingvo_Add(This,pI,index,indexAss)	\
    (This)->lpVtbl -> Add(This,pI,index,indexAss)

#define ICollLingvo_Remove(This,index)	\
    (This)->lpVtbl -> Remove(This,index)

#define ICollLingvo_LookNextIndex(This,plIndex)	\
    (This)->lpVtbl -> LookNextIndex(This,plIndex)

#define ICollLingvo_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define ICollLingvo_get_IsDirty(This,pVal)	\
    (This)->lpVtbl -> get_IsDirty(This,pVal)

#define ICollLingvo_UpdateFrom(This,pIColl)	\
    (This)->lpVtbl -> UpdateFrom(This,pIColl)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollLingvo_get__NewEnum_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollLingvo_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollLingvo_Item_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [in] */ long index,
    /* [retval][out] */ ILingvoEnum __RPC_FAR *__RPC_FAR *pvar);


void __RPC_STUB ICollLingvo_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollLingvo_get_Count_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollLingvo_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollLingvo_Add_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [in] */ ILingvoEnum __RPC_FAR *pI,
    /* [defaultvalue][optional][in] */ long index,
    /* [optional][retval][out] */ long __RPC_FAR *indexAss);


void __RPC_STUB ICollLingvo_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollLingvo_Remove_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [in] */ long index);


void __RPC_STUB ICollLingvo_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollLingvo_LookNextIndex_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *plIndex);


void __RPC_STUB ICollLingvo_LookNextIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollLingvo_Clear_Proxy( 
    ICollLingvo __RPC_FAR * This);


void __RPC_STUB ICollLingvo_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollLingvo_get_IsDirty_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICollLingvo_get_IsDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollLingvo_UpdateFrom_Proxy( 
    ICollLingvo __RPC_FAR * This,
    /* [in] */ ICollLingvo __RPC_FAR *pIColl);


void __RPC_STUB ICollLingvo_UpdateFrom_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollLingvo_INTERFACE_DEFINED__ */


#ifndef __ICompaundFiles_INTERFACE_DEFINED__
#define __ICompaundFiles_INTERFACE_DEFINED__

/* interface ICompaundFiles */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICompaundFiles;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("48118E77-0F7C-11D4-8EF1-00504E02C39D")
    ICompaundFiles : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CreateCF( 
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OpenCF( 
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CreateStorageInt( 
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OpenStorageInt( 
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage) = 0;
        
        virtual /* [vararg][helpstring][id] */ HRESULT STDMETHODCALLTYPE Sprintf( 
            /* [in] */ BSTR bsFormatString,
            /* [defaultvalue][in] */ long lBuffSize,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrArgs,
            /* [retval][out] */ BSTR __RPC_FAR *bsResult) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE HiLong( 
            /* [in] */ long lArg,
            /* [retval][out] */ long __RPC_FAR *pRes) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LoLong( 
            /* [in] */ long lArg,
            /* [retval][out] */ long __RPC_FAR *pRes) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LShiftLong( 
            /* [in] */ long lArg,
            /* [in] */ long lCnt,
            /* [retval][out] */ long __RPC_FAR *pRes) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RShiftLong( 
            /* [in] */ long lArg,
            /* [in] */ long lCnt,
            /* [retval][out] */ long __RPC_FAR *pRes) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DirOfStorage( 
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ TDRFlags shFlags,
            /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppArrNamesStg,
            /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppArrNamesStrm) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Win32ErrInfo( 
            /* [in] */ long ErrCode,
            /* [retval][out] */ BSTR __RPC_FAR *ResultStr) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CutPath( 
            /* [in] */ BSTR FullPath,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Dir,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Name,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Ext) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Commit( 
            /* [in] */ STGC1 Flags,
            /* [in] */ IStorage __RPC_FAR *Stg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Revert( 
            /* [in] */ IStorage __RPC_FAR *Stg) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OpenStreamInt( 
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR NameStream,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStream __RPC_FAR *__RPC_FAR *ppIStrm) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE PutString( 
            /* [in] */ IStream __RPC_FAR *Strm,
            /* [in] */ BSTR Str) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetString( 
            /* [in] */ IStream __RPC_FAR *Strm,
            /* [retval][out] */ BSTR __RPC_FAR *pStr) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateStreamInt( 
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR NameStream,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStream __RPC_FAR *__RPC_FAR *ppIStrm) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DestroyCF( 
            /* [in] */ IStorage __RPC_FAR *Stg,
            /* [in] */ BSTR Name) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IsEmptyStg( 
            /* [in] */ IStorage __RPC_FAR *Stg,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bEmpty) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE MakeLong( 
            /* [in] */ long Lo,
            /* [in] */ long Hi,
            /* [retval][out] */ long __RPC_FAR *lRes) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICompaundFilesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICompaundFiles __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICompaundFiles __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICompaundFiles __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateCF )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OpenCF )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateStorageInt )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OpenStorageInt )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR bsName,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);
        
        /* [vararg][helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Sprintf )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ BSTR bsFormatString,
            /* [defaultvalue][in] */ long lBuffSize,
            /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrArgs,
            /* [retval][out] */ BSTR __RPC_FAR *bsResult);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *HiLong )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ long lArg,
            /* [retval][out] */ long __RPC_FAR *pRes);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LoLong )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ long lArg,
            /* [retval][out] */ long __RPC_FAR *pRes);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LShiftLong )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ long lArg,
            /* [in] */ long lCnt,
            /* [retval][out] */ long __RPC_FAR *pRes);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RShiftLong )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ long lArg,
            /* [in] */ long lCnt,
            /* [retval][out] */ long __RPC_FAR *pRes);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DirOfStorage )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ TDRFlags shFlags,
            /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppArrNamesStg,
            /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppArrNamesStrm);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Win32ErrInfo )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ long ErrCode,
            /* [retval][out] */ BSTR __RPC_FAR *ResultStr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CutPath )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ BSTR FullPath,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Dir,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Name,
            /* [optional][out][in] */ VARIANT __RPC_FAR *Ext);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Commit )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ STGC1 Flags,
            /* [in] */ IStorage __RPC_FAR *Stg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Revert )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stg);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OpenStreamInt )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR NameStream,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStream __RPC_FAR *__RPC_FAR *ppIStrm);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PutString )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStream __RPC_FAR *Strm,
            /* [in] */ BSTR Str);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetString )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStream __RPC_FAR *Strm,
            /* [retval][out] */ BSTR __RPC_FAR *pStr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CreateStreamInt )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *pStg,
            /* [in] */ BSTR NameStream,
            /* [in] */ EnumCompConsts lAttrOpen,
            /* [retval][out] */ IStream __RPC_FAR *__RPC_FAR *ppIStrm);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DestroyCF )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stg,
            /* [in] */ BSTR Name);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *IsEmptyStg )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ IStorage __RPC_FAR *Stg,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bEmpty);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *MakeLong )( 
            ICompaundFiles __RPC_FAR * This,
            /* [in] */ long Lo,
            /* [in] */ long Hi,
            /* [retval][out] */ long __RPC_FAR *lRes);
        
        END_INTERFACE
    } ICompaundFilesVtbl;

    interface ICompaundFiles
    {
        CONST_VTBL struct ICompaundFilesVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICompaundFiles_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICompaundFiles_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICompaundFiles_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICompaundFiles_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICompaundFiles_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICompaundFiles_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICompaundFiles_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICompaundFiles_CreateCF(This,bsName,lAttrOpen,ppIStorage)	\
    (This)->lpVtbl -> CreateCF(This,bsName,lAttrOpen,ppIStorage)

#define ICompaundFiles_OpenCF(This,bsName,lAttrOpen,ppIStorage)	\
    (This)->lpVtbl -> OpenCF(This,bsName,lAttrOpen,ppIStorage)

#define ICompaundFiles_CreateStorageInt(This,pStg,bsName,lAttrOpen,ppIStorage)	\
    (This)->lpVtbl -> CreateStorageInt(This,pStg,bsName,lAttrOpen,ppIStorage)

#define ICompaundFiles_OpenStorageInt(This,pStg,bsName,lAttrOpen,ppIStorage)	\
    (This)->lpVtbl -> OpenStorageInt(This,pStg,bsName,lAttrOpen,ppIStorage)

#define ICompaundFiles_Sprintf(This,bsFormatString,lBuffSize,pparrArgs,bsResult)	\
    (This)->lpVtbl -> Sprintf(This,bsFormatString,lBuffSize,pparrArgs,bsResult)

#define ICompaundFiles_HiLong(This,lArg,pRes)	\
    (This)->lpVtbl -> HiLong(This,lArg,pRes)

#define ICompaundFiles_LoLong(This,lArg,pRes)	\
    (This)->lpVtbl -> LoLong(This,lArg,pRes)

#define ICompaundFiles_LShiftLong(This,lArg,lCnt,pRes)	\
    (This)->lpVtbl -> LShiftLong(This,lArg,lCnt,pRes)

#define ICompaundFiles_RShiftLong(This,lArg,lCnt,pRes)	\
    (This)->lpVtbl -> RShiftLong(This,lArg,lCnt,pRes)

#define ICompaundFiles_DirOfStorage(This,pStg,shFlags,ppArrNamesStg,ppArrNamesStrm)	\
    (This)->lpVtbl -> DirOfStorage(This,pStg,shFlags,ppArrNamesStg,ppArrNamesStrm)

#define ICompaundFiles_Win32ErrInfo(This,ErrCode,ResultStr)	\
    (This)->lpVtbl -> Win32ErrInfo(This,ErrCode,ResultStr)

#define ICompaundFiles_CutPath(This,FullPath,Dir,Name,Ext)	\
    (This)->lpVtbl -> CutPath(This,FullPath,Dir,Name,Ext)

#define ICompaundFiles_Commit(This,Flags,Stg)	\
    (This)->lpVtbl -> Commit(This,Flags,Stg)

#define ICompaundFiles_Revert(This,Stg)	\
    (This)->lpVtbl -> Revert(This,Stg)

#define ICompaundFiles_OpenStreamInt(This,pStg,NameStream,lAttrOpen,ppIStrm)	\
    (This)->lpVtbl -> OpenStreamInt(This,pStg,NameStream,lAttrOpen,ppIStrm)

#define ICompaundFiles_PutString(This,Strm,Str)	\
    (This)->lpVtbl -> PutString(This,Strm,Str)

#define ICompaundFiles_GetString(This,Strm,pStr)	\
    (This)->lpVtbl -> GetString(This,Strm,pStr)

#define ICompaundFiles_CreateStreamInt(This,pStg,NameStream,lAttrOpen,ppIStrm)	\
    (This)->lpVtbl -> CreateStreamInt(This,pStg,NameStream,lAttrOpen,ppIStrm)

#define ICompaundFiles_DestroyCF(This,Stg,Name)	\
    (This)->lpVtbl -> DestroyCF(This,Stg,Name)

#define ICompaundFiles_IsEmptyStg(This,Stg,bEmpty)	\
    (This)->lpVtbl -> IsEmptyStg(This,Stg,bEmpty)

#define ICompaundFiles_MakeLong(This,Lo,Hi,lRes)	\
    (This)->lpVtbl -> MakeLong(This,Lo,Hi,lRes)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_CreateCF_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ BSTR bsName,
    /* [in] */ EnumCompConsts lAttrOpen,
    /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);


void __RPC_STUB ICompaundFiles_CreateCF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_OpenCF_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ BSTR bsName,
    /* [in] */ EnumCompConsts lAttrOpen,
    /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);


void __RPC_STUB ICompaundFiles_OpenCF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_CreateStorageInt_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *pStg,
    /* [in] */ BSTR bsName,
    /* [in] */ EnumCompConsts lAttrOpen,
    /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);


void __RPC_STUB ICompaundFiles_CreateStorageInt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_OpenStorageInt_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *pStg,
    /* [in] */ BSTR bsName,
    /* [in] */ EnumCompConsts lAttrOpen,
    /* [retval][out] */ IStorage __RPC_FAR *__RPC_FAR *ppIStorage);


void __RPC_STUB ICompaundFiles_OpenStorageInt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [vararg][helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_Sprintf_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ BSTR bsFormatString,
    /* [defaultvalue][in] */ long lBuffSize,
    /* [in] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrArgs,
    /* [retval][out] */ BSTR __RPC_FAR *bsResult);


void __RPC_STUB ICompaundFiles_Sprintf_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_HiLong_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ long lArg,
    /* [retval][out] */ long __RPC_FAR *pRes);


void __RPC_STUB ICompaundFiles_HiLong_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_LoLong_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ long lArg,
    /* [retval][out] */ long __RPC_FAR *pRes);


void __RPC_STUB ICompaundFiles_LoLong_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_LShiftLong_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ long lArg,
    /* [in] */ long lCnt,
    /* [retval][out] */ long __RPC_FAR *pRes);


void __RPC_STUB ICompaundFiles_LShiftLong_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_RShiftLong_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ long lArg,
    /* [in] */ long lCnt,
    /* [retval][out] */ long __RPC_FAR *pRes);


void __RPC_STUB ICompaundFiles_RShiftLong_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_DirOfStorage_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *pStg,
    /* [in] */ TDRFlags shFlags,
    /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppArrNamesStg,
    /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppArrNamesStrm);


void __RPC_STUB ICompaundFiles_DirOfStorage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_Win32ErrInfo_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ long ErrCode,
    /* [retval][out] */ BSTR __RPC_FAR *ResultStr);


void __RPC_STUB ICompaundFiles_Win32ErrInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_CutPath_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ BSTR FullPath,
    /* [optional][out][in] */ VARIANT __RPC_FAR *Dir,
    /* [optional][out][in] */ VARIANT __RPC_FAR *Name,
    /* [optional][out][in] */ VARIANT __RPC_FAR *Ext);


void __RPC_STUB ICompaundFiles_CutPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_Commit_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ STGC1 Flags,
    /* [in] */ IStorage __RPC_FAR *Stg);


void __RPC_STUB ICompaundFiles_Commit_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_Revert_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stg);


void __RPC_STUB ICompaundFiles_Revert_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_OpenStreamInt_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *pStg,
    /* [in] */ BSTR NameStream,
    /* [in] */ EnumCompConsts lAttrOpen,
    /* [retval][out] */ IStream __RPC_FAR *__RPC_FAR *ppIStrm);


void __RPC_STUB ICompaundFiles_OpenStreamInt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_PutString_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStream __RPC_FAR *Strm,
    /* [in] */ BSTR Str);


void __RPC_STUB ICompaundFiles_PutString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_GetString_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStream __RPC_FAR *Strm,
    /* [retval][out] */ BSTR __RPC_FAR *pStr);


void __RPC_STUB ICompaundFiles_GetString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_CreateStreamInt_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *pStg,
    /* [in] */ BSTR NameStream,
    /* [in] */ EnumCompConsts lAttrOpen,
    /* [retval][out] */ IStream __RPC_FAR *__RPC_FAR *ppIStrm);


void __RPC_STUB ICompaundFiles_CreateStreamInt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_DestroyCF_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stg,
    /* [in] */ BSTR Name);


void __RPC_STUB ICompaundFiles_DestroyCF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_IsEmptyStg_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ IStorage __RPC_FAR *Stg,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *bEmpty);


void __RPC_STUB ICompaundFiles_IsEmptyStg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICompaundFiles_MakeLong_Proxy( 
    ICompaundFiles __RPC_FAR * This,
    /* [in] */ long Lo,
    /* [in] */ long Hi,
    /* [retval][out] */ long __RPC_FAR *lRes);


void __RPC_STUB ICompaundFiles_MakeLong_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICompaundFiles_INTERFACE_DEFINED__ */


#ifndef __IFactor_INTERFACE_DEFINED__
#define __IFactor_INTERFACE_DEFINED__

/* interface IFactor */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IFactor;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A754AAC3-1117-11D4-8EF2-00504E02C39D")
    IFactor : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IDEnum( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_IDEnum( 
            /* [in] */ long lVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Value( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TrustLvl( 
            /* [retval][out] */ TrustLevel __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TrustLvl( 
            /* [in] */ TrustLevel newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ResetOverWrap( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddOwerWrap( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Overwrap( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Idx( 
            /* [in] */ short sIdx,
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Idx( 
            /* [in] */ short sIdx,
            /* [in] */ float newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NIdx( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DistrType( 
            /* [retval][out] */ PrpbDistrTyp __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DistrType( 
            /* [in] */ PrpbDistrTyp newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CochiPlacement( 
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CochiPlacement( 
            /* [in] */ float newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CochiScale( 
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CochiScale( 
            /* [in] */ float newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IFactorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IFactor __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IFactor __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IFactor __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IFactor __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IFactor __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IFactor __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IFactor __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            IFactor __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IDEnum )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_IDEnum )( 
            IFactor __RPC_FAR * This,
            /* [in] */ long lVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Value )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Value )( 
            IFactor __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TrustLvl )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ TrustLevel __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TrustLvl )( 
            IFactor __RPC_FAR * This,
            /* [in] */ TrustLevel newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ID )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsDirty )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ResetOverWrap )( 
            IFactor __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddOwerWrap )( 
            IFactor __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Overwrap )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Idx )( 
            IFactor __RPC_FAR * This,
            /* [in] */ short sIdx,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Idx )( 
            IFactor __RPC_FAR * This,
            /* [in] */ short sIdx,
            /* [in] */ float newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NIdx )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DistrType )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ PrpbDistrTyp __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DistrType )( 
            IFactor __RPC_FAR * This,
            /* [in] */ PrpbDistrTyp newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CochiPlacement )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CochiPlacement )( 
            IFactor __RPC_FAR * This,
            /* [in] */ float newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CochiScale )( 
            IFactor __RPC_FAR * This,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CochiScale )( 
            IFactor __RPC_FAR * This,
            /* [in] */ float newVal);
        
        END_INTERFACE
    } IFactorVtbl;

    interface IFactor
    {
        CONST_VTBL struct IFactorVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IFactor_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IFactor_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IFactor_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IFactor_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IFactor_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IFactor_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IFactor_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IFactor_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define IFactor_put_Name(This,newVal)	\
    (This)->lpVtbl -> put_Name(This,newVal)

#define IFactor_get_IDEnum(This,pVal)	\
    (This)->lpVtbl -> get_IDEnum(This,pVal)

#define IFactor_put_IDEnum(This,lVal)	\
    (This)->lpVtbl -> put_IDEnum(This,lVal)

#define IFactor_get_Value(This,pVal)	\
    (This)->lpVtbl -> get_Value(This,pVal)

#define IFactor_put_Value(This,newVal)	\
    (This)->lpVtbl -> put_Value(This,newVal)

#define IFactor_get_TrustLvl(This,pVal)	\
    (This)->lpVtbl -> get_TrustLvl(This,pVal)

#define IFactor_put_TrustLvl(This,newVal)	\
    (This)->lpVtbl -> put_TrustLvl(This,newVal)

#define IFactor_get_ID(This,pVal)	\
    (This)->lpVtbl -> get_ID(This,pVal)

#define IFactor_get_IsDirty(This,pVal)	\
    (This)->lpVtbl -> get_IsDirty(This,pVal)

#define IFactor_ResetOverWrap(This)	\
    (This)->lpVtbl -> ResetOverWrap(This)

#define IFactor_AddOwerWrap(This)	\
    (This)->lpVtbl -> AddOwerWrap(This)

#define IFactor_get_Overwrap(This,pVal)	\
    (This)->lpVtbl -> get_Overwrap(This,pVal)

#define IFactor_get_Idx(This,sIdx,pVal)	\
    (This)->lpVtbl -> get_Idx(This,sIdx,pVal)

#define IFactor_put_Idx(This,sIdx,newVal)	\
    (This)->lpVtbl -> put_Idx(This,sIdx,newVal)

#define IFactor_get_NIdx(This,pVal)	\
    (This)->lpVtbl -> get_NIdx(This,pVal)

#define IFactor_get_DistrType(This,pVal)	\
    (This)->lpVtbl -> get_DistrType(This,pVal)

#define IFactor_put_DistrType(This,newVal)	\
    (This)->lpVtbl -> put_DistrType(This,newVal)

#define IFactor_get_CochiPlacement(This,pVal)	\
    (This)->lpVtbl -> get_CochiPlacement(This,pVal)

#define IFactor_put_CochiPlacement(This,newVal)	\
    (This)->lpVtbl -> put_CochiPlacement(This,newVal)

#define IFactor_get_CochiScale(This,pVal)	\
    (This)->lpVtbl -> get_CochiScale(This,pVal)

#define IFactor_put_CochiScale(This,newVal)	\
    (This)->lpVtbl -> put_CochiScale(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_Name_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_Name_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IFactor_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_IDEnum_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_IDEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_IDEnum_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ long lVal);


void __RPC_STUB IFactor_put_IDEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_Value_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_Value_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_Value_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IFactor_put_Value_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_TrustLvl_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ TrustLevel __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_TrustLvl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_TrustLvl_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ TrustLevel newVal);


void __RPC_STUB IFactor_put_TrustLvl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_ID_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_ID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_IsDirty_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_IsDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactor_ResetOverWrap_Proxy( 
    IFactor __RPC_FAR * This);


void __RPC_STUB IFactor_ResetOverWrap_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IFactor_AddOwerWrap_Proxy( 
    IFactor __RPC_FAR * This);


void __RPC_STUB IFactor_AddOwerWrap_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_Overwrap_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_Overwrap_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_Idx_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ short sIdx,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_Idx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_Idx_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ short sIdx,
    /* [in] */ float newVal);


void __RPC_STUB IFactor_put_Idx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_NIdx_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_NIdx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_DistrType_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ PrpbDistrTyp __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_DistrType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_DistrType_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ PrpbDistrTyp newVal);


void __RPC_STUB IFactor_put_DistrType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_CochiPlacement_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_CochiPlacement_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_CochiPlacement_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ float newVal);


void __RPC_STUB IFactor_put_CochiPlacement_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFactor_get_CochiScale_Proxy( 
    IFactor __RPC_FAR * This,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IFactor_get_CochiScale_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFactor_put_CochiScale_Proxy( 
    IFactor __RPC_FAR * This,
    /* [in] */ float newVal);


void __RPC_STUB IFactor_put_CochiScale_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IFactor_INTERFACE_DEFINED__ */


#ifndef __ILongKey_INTERFACE_DEFINED__
#define __ILongKey_INTERFACE_DEFINED__

/* interface ILongKey */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_ILongKey;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("41118E75-0F7C-11D4-8EF1-00504E02C39A")
    ILongKey : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Set( 
            /* [in] */ long lKey) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Get( 
            /* [retval][out] */ long __RPC_FAR *plKey) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILongKeyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ILongKey __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ILongKey __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ILongKey __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Set )( 
            ILongKey __RPC_FAR * This,
            /* [in] */ long lKey);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Get )( 
            ILongKey __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *plKey);
        
        END_INTERFACE
    } ILongKeyVtbl;

    interface ILongKey
    {
        CONST_VTBL struct ILongKeyVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILongKey_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILongKey_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILongKey_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILongKey_Set(This,lKey)	\
    (This)->lpVtbl -> Set(This,lKey)

#define ILongKey_Get(This,plKey)	\
    (This)->lpVtbl -> Get(This,plKey)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ILongKey_Set_Proxy( 
    ILongKey __RPC_FAR * This,
    /* [in] */ long lKey);


void __RPC_STUB ILongKey_Set_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ILongKey_Get_Proxy( 
    ILongKey __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *plKey);


void __RPC_STUB ILongKey_Get_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILongKey_INTERFACE_DEFINED__ */


#ifndef __IClonable_INTERFACE_DEFINED__
#define __IClonable_INTERFACE_DEFINED__

/* interface IClonable */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_IClonable;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("41118E75-9F71-11D5-8AF1-00504E02C39A")
    IClonable : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Clone( 
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IClonableVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IClonable __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IClonable __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IClonable __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clone )( 
            IClonable __RPC_FAR * This,
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);
        
        END_INTERFACE
    } IClonableVtbl;

    interface IClonable
    {
        CONST_VTBL struct IClonableVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IClonable_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IClonable_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IClonable_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IClonable_Clone(This,ppUnk)	\
    (This)->lpVtbl -> Clone(This,ppUnk)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IClonable_Clone_Proxy( 
    IClonable __RPC_FAR * This,
    /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);


void __RPC_STUB IClonable_Clone_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IClonable_INTERFACE_DEFINED__ */


#ifndef __IBSTRKey_INTERFACE_DEFINED__
#define __IBSTRKey_INTERFACE_DEFINED__

/* interface IBSTRKey */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_IBSTRKey;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("41118E75-0F7C-11D4-8EF2-00504E02C39A")
    IBSTRKey : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Set( 
            /* [in] */ BSTR lKey) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Get( 
            /* [retval][out] */ BSTR __RPC_FAR *plKey) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IBSTRKeyVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IBSTRKey __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IBSTRKey __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IBSTRKey __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Set )( 
            IBSTRKey __RPC_FAR * This,
            /* [in] */ BSTR lKey);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Get )( 
            IBSTRKey __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *plKey);
        
        END_INTERFACE
    } IBSTRKeyVtbl;

    interface IBSTRKey
    {
        CONST_VTBL struct IBSTRKeyVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IBSTRKey_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IBSTRKey_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IBSTRKey_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IBSTRKey_Set(This,lKey)	\
    (This)->lpVtbl -> Set(This,lKey)

#define IBSTRKey_Get(This,plKey)	\
    (This)->lpVtbl -> Get(This,plKey)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IBSTRKey_Set_Proxy( 
    IBSTRKey __RPC_FAR * This,
    /* [in] */ BSTR lKey);


void __RPC_STUB IBSTRKey_Set_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IBSTRKey_Get_Proxy( 
    IBSTRKey __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *plKey);


void __RPC_STUB IBSTRKey_Get_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IBSTRKey_INTERFACE_DEFINED__ */


#ifndef __IMGertNet_Debug_INTERFACE_DEFINED__
#define __IMGertNet_Debug_INTERFACE_DEFINED__

/* interface IMGertNet_Debug */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_IMGertNet_Debug;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("59619591-19F2-11d4-8F02-00504E02C39D")
    IMGertNet_Debug : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE GetPtsForFactor( 
            /* [in] */ BSTR bsFShortName,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrPoints) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetDensityForFactor( 
            /* [in] */ BSTR bsFShortName,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrPoints) = 0;
        
        virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE TestFunc( 
            /* [in] */ BSTR bsFac,
            /* [in] */ double dX,
            /* [retval][out] */ double __RPC_FAR *dY) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IMGertNet_DebugVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IMGertNet_Debug __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IMGertNet_Debug __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IMGertNet_Debug __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPtsForFactor )( 
            IMGertNet_Debug __RPC_FAR * This,
            /* [in] */ BSTR bsFShortName,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrPoints);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetDensityForFactor )( 
            IMGertNet_Debug __RPC_FAR * This,
            /* [in] */ BSTR bsFShortName,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrPoints);
        
        /* [helpstring] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *TestFunc )( 
            IMGertNet_Debug __RPC_FAR * This,
            /* [in] */ BSTR bsFac,
            /* [in] */ double dX,
            /* [retval][out] */ double __RPC_FAR *dY);
        
        END_INTERFACE
    } IMGertNet_DebugVtbl;

    interface IMGertNet_Debug
    {
        CONST_VTBL struct IMGertNet_DebugVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMGertNet_Debug_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IMGertNet_Debug_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IMGertNet_Debug_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IMGertNet_Debug_GetPtsForFactor(This,bsFShortName,pparrPoints)	\
    (This)->lpVtbl -> GetPtsForFactor(This,bsFShortName,pparrPoints)

#define IMGertNet_Debug_GetDensityForFactor(This,bsFShortName,pparrPoints)	\
    (This)->lpVtbl -> GetDensityForFactor(This,bsFShortName,pparrPoints)

#define IMGertNet_Debug_TestFunc(This,bsFac,dX,dY)	\
    (This)->lpVtbl -> TestFunc(This,bsFac,dX,dY)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IMGertNet_Debug_GetPtsForFactor_Proxy( 
    IMGertNet_Debug __RPC_FAR * This,
    /* [in] */ BSTR bsFShortName,
    /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrPoints);


void __RPC_STUB IMGertNet_Debug_GetPtsForFactor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IMGertNet_Debug_GetDensityForFactor_Proxy( 
    IMGertNet_Debug __RPC_FAR * This,
    /* [in] */ BSTR bsFShortName,
    /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrPoints);


void __RPC_STUB IMGertNet_Debug_GetDensityForFactor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT STDMETHODCALLTYPE IMGertNet_Debug_TestFunc_Proxy( 
    IMGertNet_Debug __RPC_FAR * This,
    /* [in] */ BSTR bsFac,
    /* [in] */ double dX,
    /* [retval][out] */ double __RPC_FAR *dY);


void __RPC_STUB IMGertNet_Debug_TestFunc_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IMGertNet_Debug_INTERFACE_DEFINED__ */


#ifndef __ICollFac_INTERFACE_DEFINED__
#define __ICollFac_INTERFACE_DEFINED__

/* interface ICollFac */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICollFac;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A754AAC5-1117-11D4-8EF2-00504E02C39D")
    ICollFac : public IDispatch
    {
    public:
        virtual /* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ BSTR index,
            /* [retval][out] */ IFactor __RPC_FAR *__RPC_FAR *pvar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ IFactor __RPC_FAR *pI,
            /* [in] */ BSTR key) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ BSTR key) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetIndexOfItem( 
            /* [in] */ IFactor __RPC_FAR *pEN,
            /* [retval][out] */ BSTR __RPC_FAR *pKey) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollFacVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollFac __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollFac __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICollFac __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][restricted][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICollFac __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ BSTR index,
            /* [retval][out] */ IFactor __RPC_FAR *__RPC_FAR *pvar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICollFac __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ IFactor __RPC_FAR *pI,
            /* [in] */ BSTR key);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ BSTR key);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIndexOfItem )( 
            ICollFac __RPC_FAR * This,
            /* [in] */ IFactor __RPC_FAR *pEN,
            /* [retval][out] */ BSTR __RPC_FAR *pKey);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            ICollFac __RPC_FAR * This);
        
        END_INTERFACE
    } ICollFacVtbl;

    interface ICollFac
    {
        CONST_VTBL struct ICollFacVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollFac_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollFac_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollFac_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollFac_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICollFac_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICollFac_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICollFac_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICollFac_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICollFac_Item(This,index,pvar)	\
    (This)->lpVtbl -> Item(This,index,pvar)

#define ICollFac_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICollFac_Add(This,pI,key)	\
    (This)->lpVtbl -> Add(This,pI,key)

#define ICollFac_Remove(This,key)	\
    (This)->lpVtbl -> Remove(This,key)

#define ICollFac_GetIndexOfItem(This,pEN,pKey)	\
    (This)->lpVtbl -> GetIndexOfItem(This,pEN,pKey)

#define ICollFac_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][restricted][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFac_get__NewEnum_Proxy( 
    ICollFac __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICollFac_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFac_Item_Proxy( 
    ICollFac __RPC_FAR * This,
    /* [in] */ BSTR index,
    /* [retval][out] */ IFactor __RPC_FAR *__RPC_FAR *pvar);


void __RPC_STUB ICollFac_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICollFac_get_Count_Proxy( 
    ICollFac __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollFac_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFac_Add_Proxy( 
    ICollFac __RPC_FAR * This,
    /* [in] */ IFactor __RPC_FAR *pI,
    /* [in] */ BSTR key);


void __RPC_STUB ICollFac_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFac_Remove_Proxy( 
    ICollFac __RPC_FAR * This,
    /* [in] */ BSTR key);


void __RPC_STUB ICollFac_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICollFac_GetIndexOfItem_Proxy( 
    ICollFac __RPC_FAR * This,
    /* [in] */ IFactor __RPC_FAR *pEN,
    /* [retval][out] */ BSTR __RPC_FAR *pKey);


void __RPC_STUB ICollFac_GetIndexOfItem_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE ICollFac_Clear_Proxy( 
    ICollFac __RPC_FAR * This);


void __RPC_STUB ICollFac_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollFac_INTERFACE_DEFINED__ */


#ifndef __IMGertNet_INTERFACE_DEFINED__
#define __IMGertNet_INTERFACE_DEFINED__

/* interface IMGertNet */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IMGertNet;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A754AAC7-1117-11D4-8EF2-00504E02C39D")
    IMGertNet : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Enumerators( 
            /* [retval][out] */ ICollLingvo __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE get_Factors( 
            /* [retval][out] */ ICollFac __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NotifyWnd( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NotifyWnd( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RunMode( 
            /* [retval][out] */ ModelType __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RunMode( 
            /* [in] */ ModelType newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Run( 
            /* [in] */ long lNExperience,
            /* [in] */ long lNRunInOne,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL bResetModel = 0,
            /* [defaultvalue][optional][in] */ long lRndBase = -1,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL bPrepareOnly = 0) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Cancel( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Counter( 
            /* [in] */ WCHAR cState,
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Dx( 
            /* [in] */ WCHAR cState,
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Mx( 
            /* [in] */ WCHAR cState,
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TimeWork( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NumberStates( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_K( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_K( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_N( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_N( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SampleN( 
            /* [in] */ long lIndex,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrSmpl) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NotifyStep( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NotifyStep( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFactorPresent( 
            /* [in] */ IFactor __RPC_FAR *pF,
            /* [optional][out][in] */ VARIANT __RPC_FAR *strDescription,
            /* [optional][out][in] */ VARIANT __RPC_FAR *strTrustLvl,
            /* [optional][out][in] */ VARIANT __RPC_FAR *dVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFactorPresentForName( 
            /* [in] */ BSTR bsShortName,
            /* [out] */ BSTR __RPC_FAR *pbsDescription,
            /* [out] */ BSTR __RPC_FAR *pbsTrustLvl,
            /* [out] */ double __RPC_FAR *pdVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ApplySF( 
            /* [in] */ ICollSF __RPC_FAR *pSF,
            /* [optional][out] */ VARIANT_BOOL __RPC_FAR *pbOverWrap) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetEnumForFactor( 
            /* [in] */ IFactor __RPC_FAR *pF,
            /* [retval][out] */ ILingvoEnum __RPC_FAR *__RPC_FAR *ppLE) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetInfSampleK( 
            /* [in] */ WCHAR cState,
            /* [out] */ long __RPC_FAR *pMin,
            /* [out] */ long __RPC_FAR *pMax,
            /* [out] */ double __RPC_FAR *pMx,
            /* [out] */ double __RPC_FAR *pDx) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsRunning( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetInfSampleKIdx( 
            /* [in] */ WCHAR cState,
            /* [out] */ double __RPC_FAR *pMin,
            /* [out] */ double __RPC_FAR *pMax,
            /* [out] */ double __RPC_FAR *pMx,
            /* [out] */ double __RPC_FAR *pDx) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ChangeDirty( 
            /* [in] */ VARIANT_BOOL bFlDirty) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CalcDeltaQ( 
            /* [in] */ ICollSF __RPC_FAR *pCollSF,
            /* [in] */ long lK,
            /* [in] */ long lN,
            /* [defaultvalue][optional][in] */ short shMedVal = -1) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CalcSpeed( 
            /* [retval][out] */ ThrdPriority __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CalcSpeed( 
            /* [in] */ ThrdPriority newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_OptimizationType( 
            /* [retval][out] */ OptType __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_OptimizationType( 
            /* [in] */ OptType newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RemovLingvoOverwrap( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RemovLingvoOverwrap( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AverageDamage( 
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_AverageDamage( 
            /* [in] */ CURRENCY newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_MoneyForSF( 
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_MoneyForSF( 
            /* [in] */ CURRENCY newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Snapshot( 
            /* [in] */ VARIANT_BOOL bMake) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Revert( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_UserProp( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_UserProp( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RndBase( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RndBase( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CurrIterStr( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TotalIterStr( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StatOn( 
            /* [retval][out] */ TStatFlag __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StatOn( 
            /* [in] */ TStatFlag newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StatIn( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StatIn( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TstApplyTCH( 
            /* [in] */ IFactor __RPC_FAR *pFac,
            /* [in] */ IFChange __RPC_FAR *pFC,
            /* [retval][out] */ TrustLevel __RPC_FAR *pLvl) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ImitateCount( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_OptCount( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RngCount( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ModuleProgID( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ModuleProgID( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NDiv( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NDiv( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_UnionThreshold( 
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_UnionThreshold( 
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IntegralAccuracy( 
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_IntegralAccuracy( 
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NFormatPr( 
            /* [retval][out] */ NumberFormat __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NFormatPr( 
            /* [in] */ NumberFormat newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NFormatOther( 
            /* [retval][out] */ NumberFormat __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NFormatOther( 
            /* [in] */ NumberFormat newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NDigitsPr( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NDigitsPr( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NDigitsOther( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NDigitsOther( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NFormat( 
            /* [in] */ NFormatTyp nf,
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_VParam( 
            /* [in] */ short Ventil,
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_VParam( 
            /* [in] */ short Ventil,
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_VParamIndistinct( 
            /* [in] */ short Ventil,
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_VParamIndistinct( 
            /* [in] */ short Ventil,
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_VProbability( 
            /* [in] */ short Ventil,
            /* [retval][out] */ double __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_VProbability( 
            /* [in] */ short Ventil,
            /* [in] */ double newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ModuleConfig( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ModuleConfig( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TimeWork2( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OptimizeSelectSF( 
            /* [in] */ OptTask otTask,
            /* [in] */ ICollSF __RPC_FAR *pCollSF,
            /* [in] */ CURRENCY cyMax,
            /* [in] */ double dDeltaQ,
            /* [in] */ double dP0,
            /* [in] */ short shNRetAlternate) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsRunning2( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsRunningOpt( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_OptimizResultsGetAndClear( 
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TimeWork3( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CalibrateModel( 
            /* [in] */ double dP1,
            /* [in] */ double dP2,
            /* [in] */ double dP3,
            /* [in] */ double dP4) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LastCalcError( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetModelStat( 
            /* [in] */ short shIdx,
            /* [out] */ TypANodesStat __RPC_FAR *pNS,
            /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrSmpl,
            /* [in] */ VARIANT_BOOL bFullGet) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsDirty( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IMGertNetVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IMGertNet __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IMGertNet __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IMGertNet __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Enumerators )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ ICollLingvo __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][defaultcollelem][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Factors )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ ICollFac __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NotifyWnd )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NotifyWnd )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RunMode )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ ModelType __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RunMode )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ ModelType newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Run )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long lNExperience,
            /* [in] */ long lNRunInOne,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL bResetModel,
            /* [defaultvalue][optional][in] */ long lRndBase,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL bPrepareOnly);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            IMGertNet __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Counter )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ WCHAR cState,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Dx )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ WCHAR cState,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Mx )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ WCHAR cState,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TimeWork )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NumberStates )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_K )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_K )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_N )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_N )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SampleN )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long lIndex,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrSmpl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NotifyStep )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NotifyStep )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFactorPresent )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ IFactor __RPC_FAR *pF,
            /* [optional][out][in] */ VARIANT __RPC_FAR *strDescription,
            /* [optional][out][in] */ VARIANT __RPC_FAR *strTrustLvl,
            /* [optional][out][in] */ VARIANT __RPC_FAR *dVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFactorPresentForName )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ BSTR bsShortName,
            /* [out] */ BSTR __RPC_FAR *pbsDescription,
            /* [out] */ BSTR __RPC_FAR *pbsTrustLvl,
            /* [out] */ double __RPC_FAR *pdVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ApplySF )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ ICollSF __RPC_FAR *pSF,
            /* [optional][out] */ VARIANT_BOOL __RPC_FAR *pbOverWrap);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetEnumForFactor )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ IFactor __RPC_FAR *pF,
            /* [retval][out] */ ILingvoEnum __RPC_FAR *__RPC_FAR *ppLE);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetInfSampleK )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ WCHAR cState,
            /* [out] */ long __RPC_FAR *pMin,
            /* [out] */ long __RPC_FAR *pMax,
            /* [out] */ double __RPC_FAR *pMx,
            /* [out] */ double __RPC_FAR *pDx);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsRunning )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetInfSampleKIdx )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ WCHAR cState,
            /* [out] */ double __RPC_FAR *pMin,
            /* [out] */ double __RPC_FAR *pMax,
            /* [out] */ double __RPC_FAR *pMx,
            /* [out] */ double __RPC_FAR *pDx);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ChangeDirty )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bFlDirty);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CalcDeltaQ )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ ICollSF __RPC_FAR *pCollSF,
            /* [in] */ long lK,
            /* [in] */ long lN,
            /* [defaultvalue][optional][in] */ short shMedVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CalcSpeed )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ ThrdPriority __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CalcSpeed )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ ThrdPriority newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_OptimizationType )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ OptType __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_OptimizationType )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ OptType newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RemovLingvoOverwrap )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RemovLingvoOverwrap )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AverageDamage )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_AverageDamage )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ CURRENCY newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_MoneyForSF )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ CURRENCY __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_MoneyForSF )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ CURRENCY newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Snapshot )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bMake);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Revert )( 
            IMGertNet __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_UserProp )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_UserProp )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RndBase )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RndBase )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CurrIterStr )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TotalIterStr )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StatOn )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ TStatFlag __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StatOn )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ TStatFlag newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StatIn )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_StatIn )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *TstApplyTCH )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ IFactor __RPC_FAR *pFac,
            /* [in] */ IFChange __RPC_FAR *pFC,
            /* [retval][out] */ TrustLevel __RPC_FAR *pLvl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ImitateCount )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_OptCount )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RngCount )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ModuleProgID )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ModuleProgID )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NDiv )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NDiv )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_UnionThreshold )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_UnionThreshold )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IntegralAccuracy )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_IntegralAccuracy )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NFormatPr )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ NumberFormat __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NFormatPr )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ NumberFormat newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NFormatOther )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ NumberFormat __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NFormatOther )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ NumberFormat newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NDigitsPr )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NDigitsPr )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NDigitsOther )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_NDigitsOther )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_NFormat )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ NFormatTyp nf,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_VParam )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short Ventil,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_VParam )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short Ventil,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_VParamIndistinct )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short Ventil,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_VParamIndistinct )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short Ventil,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_VProbability )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short Ventil,
            /* [retval][out] */ double __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_VProbability )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short Ventil,
            /* [in] */ double newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ModuleConfig )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ModuleConfig )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TimeWork2 )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OptimizeSelectSF )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ OptTask otTask,
            /* [in] */ ICollSF __RPC_FAR *pCollSF,
            /* [in] */ CURRENCY cyMax,
            /* [in] */ double dDeltaQ,
            /* [in] */ double dP0,
            /* [in] */ short shNRetAlternate);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsRunning2 )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsRunningOpt )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_OptimizResultsGetAndClear )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TimeWork3 )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CalibrateModel )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ double dP1,
            /* [in] */ double dP2,
            /* [in] */ double dP3,
            /* [in] */ double dP4);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LastCalcError )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetModelStat )( 
            IMGertNet __RPC_FAR * This,
            /* [in] */ short shIdx,
            /* [out] */ TypANodesStat __RPC_FAR *pNS,
            /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrSmpl,
            /* [in] */ VARIANT_BOOL bFullGet);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsDirty )( 
            IMGertNet __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        END_INTERFACE
    } IMGertNetVtbl;

    interface IMGertNet
    {
        CONST_VTBL struct IMGertNetVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMGertNet_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IMGertNet_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IMGertNet_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IMGertNet_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IMGertNet_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IMGertNet_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IMGertNet_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IMGertNet_get_Enumerators(This,pVal)	\
    (This)->lpVtbl -> get_Enumerators(This,pVal)

#define IMGertNet_get_Factors(This,pVal)	\
    (This)->lpVtbl -> get_Factors(This,pVal)

#define IMGertNet_get_NotifyWnd(This,pVal)	\
    (This)->lpVtbl -> get_NotifyWnd(This,pVal)

#define IMGertNet_put_NotifyWnd(This,newVal)	\
    (This)->lpVtbl -> put_NotifyWnd(This,newVal)

#define IMGertNet_get_RunMode(This,pVal)	\
    (This)->lpVtbl -> get_RunMode(This,pVal)

#define IMGertNet_put_RunMode(This,newVal)	\
    (This)->lpVtbl -> put_RunMode(This,newVal)

#define IMGertNet_Run(This,lNExperience,lNRunInOne,bResetModel,lRndBase,bPrepareOnly)	\
    (This)->lpVtbl -> Run(This,lNExperience,lNRunInOne,bResetModel,lRndBase,bPrepareOnly)

#define IMGertNet_Cancel(This)	\
    (This)->lpVtbl -> Cancel(This)

#define IMGertNet_get_Counter(This,cState,pVal)	\
    (This)->lpVtbl -> get_Counter(This,cState,pVal)

#define IMGertNet_get_Dx(This,cState,pVal)	\
    (This)->lpVtbl -> get_Dx(This,cState,pVal)

#define IMGertNet_get_Mx(This,cState,pVal)	\
    (This)->lpVtbl -> get_Mx(This,cState,pVal)

#define IMGertNet_get_TimeWork(This,pVal)	\
    (This)->lpVtbl -> get_TimeWork(This,pVal)

#define IMGertNet_get_NumberStates(This,pVal)	\
    (This)->lpVtbl -> get_NumberStates(This,pVal)

#define IMGertNet_get_K(This,pVal)	\
    (This)->lpVtbl -> get_K(This,pVal)

#define IMGertNet_put_K(This,newVal)	\
    (This)->lpVtbl -> put_K(This,newVal)

#define IMGertNet_get_N(This,pVal)	\
    (This)->lpVtbl -> get_N(This,pVal)

#define IMGertNet_put_N(This,newVal)	\
    (This)->lpVtbl -> put_N(This,newVal)

#define IMGertNet_get_SampleN(This,lIndex,pparrSmpl)	\
    (This)->lpVtbl -> get_SampleN(This,lIndex,pparrSmpl)

#define IMGertNet_get_NotifyStep(This,pVal)	\
    (This)->lpVtbl -> get_NotifyStep(This,pVal)

#define IMGertNet_put_NotifyStep(This,newVal)	\
    (This)->lpVtbl -> put_NotifyStep(This,newVal)

#define IMGertNet_GetFactorPresent(This,pF,strDescription,strTrustLvl,dVal)	\
    (This)->lpVtbl -> GetFactorPresent(This,pF,strDescription,strTrustLvl,dVal)

#define IMGertNet_GetFactorPresentForName(This,bsShortName,pbsDescription,pbsTrustLvl,pdVal)	\
    (This)->lpVtbl -> GetFactorPresentForName(This,bsShortName,pbsDescription,pbsTrustLvl,pdVal)

#define IMGertNet_ApplySF(This,pSF,pbOverWrap)	\
    (This)->lpVtbl -> ApplySF(This,pSF,pbOverWrap)

#define IMGertNet_GetEnumForFactor(This,pF,ppLE)	\
    (This)->lpVtbl -> GetEnumForFactor(This,pF,ppLE)

#define IMGertNet_GetInfSampleK(This,cState,pMin,pMax,pMx,pDx)	\
    (This)->lpVtbl -> GetInfSampleK(This,cState,pMin,pMax,pMx,pDx)

#define IMGertNet_get_IsRunning(This,pVal)	\
    (This)->lpVtbl -> get_IsRunning(This,pVal)

#define IMGertNet_GetInfSampleKIdx(This,cState,pMin,pMax,pMx,pDx)	\
    (This)->lpVtbl -> GetInfSampleKIdx(This,cState,pMin,pMax,pMx,pDx)

#define IMGertNet_ChangeDirty(This,bFlDirty)	\
    (This)->lpVtbl -> ChangeDirty(This,bFlDirty)

#define IMGertNet_CalcDeltaQ(This,pCollSF,lK,lN,shMedVal)	\
    (This)->lpVtbl -> CalcDeltaQ(This,pCollSF,lK,lN,shMedVal)

#define IMGertNet_get_CalcSpeed(This,pVal)	\
    (This)->lpVtbl -> get_CalcSpeed(This,pVal)

#define IMGertNet_put_CalcSpeed(This,newVal)	\
    (This)->lpVtbl -> put_CalcSpeed(This,newVal)

#define IMGertNet_get_OptimizationType(This,pVal)	\
    (This)->lpVtbl -> get_OptimizationType(This,pVal)

#define IMGertNet_put_OptimizationType(This,newVal)	\
    (This)->lpVtbl -> put_OptimizationType(This,newVal)

#define IMGertNet_get_RemovLingvoOverwrap(This,pVal)	\
    (This)->lpVtbl -> get_RemovLingvoOverwrap(This,pVal)

#define IMGertNet_put_RemovLingvoOverwrap(This,newVal)	\
    (This)->lpVtbl -> put_RemovLingvoOverwrap(This,newVal)

#define IMGertNet_get_AverageDamage(This,pVal)	\
    (This)->lpVtbl -> get_AverageDamage(This,pVal)

#define IMGertNet_put_AverageDamage(This,newVal)	\
    (This)->lpVtbl -> put_AverageDamage(This,newVal)

#define IMGertNet_get_MoneyForSF(This,pVal)	\
    (This)->lpVtbl -> get_MoneyForSF(This,pVal)

#define IMGertNet_put_MoneyForSF(This,newVal)	\
    (This)->lpVtbl -> put_MoneyForSF(This,newVal)

#define IMGertNet_Snapshot(This,bMake)	\
    (This)->lpVtbl -> Snapshot(This,bMake)

#define IMGertNet_Revert(This)	\
    (This)->lpVtbl -> Revert(This)

#define IMGertNet_get_UserProp(This,pVal)	\
    (This)->lpVtbl -> get_UserProp(This,pVal)

#define IMGertNet_put_UserProp(This,newVal)	\
    (This)->lpVtbl -> put_UserProp(This,newVal)

#define IMGertNet_get_RndBase(This,pVal)	\
    (This)->lpVtbl -> get_RndBase(This,pVal)

#define IMGertNet_put_RndBase(This,newVal)	\
    (This)->lpVtbl -> put_RndBase(This,newVal)

#define IMGertNet_get_CurrIterStr(This,pVal)	\
    (This)->lpVtbl -> get_CurrIterStr(This,pVal)

#define IMGertNet_get_TotalIterStr(This,pVal)	\
    (This)->lpVtbl -> get_TotalIterStr(This,pVal)

#define IMGertNet_get_StatOn(This,pVal)	\
    (This)->lpVtbl -> get_StatOn(This,pVal)

#define IMGertNet_put_StatOn(This,newVal)	\
    (This)->lpVtbl -> put_StatOn(This,newVal)

#define IMGertNet_get_StatIn(This,pVal)	\
    (This)->lpVtbl -> get_StatIn(This,pVal)

#define IMGertNet_put_StatIn(This,newVal)	\
    (This)->lpVtbl -> put_StatIn(This,newVal)

#define IMGertNet_TstApplyTCH(This,pFac,pFC,pLvl)	\
    (This)->lpVtbl -> TstApplyTCH(This,pFac,pFC,pLvl)

#define IMGertNet_get_ImitateCount(This,pVal)	\
    (This)->lpVtbl -> get_ImitateCount(This,pVal)

#define IMGertNet_get_OptCount(This,pVal)	\
    (This)->lpVtbl -> get_OptCount(This,pVal)

#define IMGertNet_get_RngCount(This,pVal)	\
    (This)->lpVtbl -> get_RngCount(This,pVal)

#define IMGertNet_get_ModuleProgID(This,pVal)	\
    (This)->lpVtbl -> get_ModuleProgID(This,pVal)

#define IMGertNet_put_ModuleProgID(This,newVal)	\
    (This)->lpVtbl -> put_ModuleProgID(This,newVal)

#define IMGertNet_get_NDiv(This,pVal)	\
    (This)->lpVtbl -> get_NDiv(This,pVal)

#define IMGertNet_put_NDiv(This,newVal)	\
    (This)->lpVtbl -> put_NDiv(This,newVal)

#define IMGertNet_get_UnionThreshold(This,pVal)	\
    (This)->lpVtbl -> get_UnionThreshold(This,pVal)

#define IMGertNet_put_UnionThreshold(This,newVal)	\
    (This)->lpVtbl -> put_UnionThreshold(This,newVal)

#define IMGertNet_get_IntegralAccuracy(This,pVal)	\
    (This)->lpVtbl -> get_IntegralAccuracy(This,pVal)

#define IMGertNet_put_IntegralAccuracy(This,newVal)	\
    (This)->lpVtbl -> put_IntegralAccuracy(This,newVal)

#define IMGertNet_get_NFormatPr(This,pVal)	\
    (This)->lpVtbl -> get_NFormatPr(This,pVal)

#define IMGertNet_put_NFormatPr(This,newVal)	\
    (This)->lpVtbl -> put_NFormatPr(This,newVal)

#define IMGertNet_get_NFormatOther(This,pVal)	\
    (This)->lpVtbl -> get_NFormatOther(This,pVal)

#define IMGertNet_put_NFormatOther(This,newVal)	\
    (This)->lpVtbl -> put_NFormatOther(This,newVal)

#define IMGertNet_get_NDigitsPr(This,pVal)	\
    (This)->lpVtbl -> get_NDigitsPr(This,pVal)

#define IMGertNet_put_NDigitsPr(This,newVal)	\
    (This)->lpVtbl -> put_NDigitsPr(This,newVal)

#define IMGertNet_get_NDigitsOther(This,pVal)	\
    (This)->lpVtbl -> get_NDigitsOther(This,pVal)

#define IMGertNet_put_NDigitsOther(This,newVal)	\
    (This)->lpVtbl -> put_NDigitsOther(This,newVal)

#define IMGertNet_get_NFormat(This,nf,pVal)	\
    (This)->lpVtbl -> get_NFormat(This,nf,pVal)

#define IMGertNet_get_VParam(This,Ventil,pVal)	\
    (This)->lpVtbl -> get_VParam(This,Ventil,pVal)

#define IMGertNet_put_VParam(This,Ventil,newVal)	\
    (This)->lpVtbl -> put_VParam(This,Ventil,newVal)

#define IMGertNet_get_VParamIndistinct(This,Ventil,pVal)	\
    (This)->lpVtbl -> get_VParamIndistinct(This,Ventil,pVal)

#define IMGertNet_put_VParamIndistinct(This,Ventil,newVal)	\
    (This)->lpVtbl -> put_VParamIndistinct(This,Ventil,newVal)

#define IMGertNet_get_VProbability(This,Ventil,pVal)	\
    (This)->lpVtbl -> get_VProbability(This,Ventil,pVal)

#define IMGertNet_put_VProbability(This,Ventil,newVal)	\
    (This)->lpVtbl -> put_VProbability(This,Ventil,newVal)

#define IMGertNet_get_ModuleConfig(This,pVal)	\
    (This)->lpVtbl -> get_ModuleConfig(This,pVal)

#define IMGertNet_put_ModuleConfig(This,newVal)	\
    (This)->lpVtbl -> put_ModuleConfig(This,newVal)

#define IMGertNet_get_TimeWork2(This,pVal)	\
    (This)->lpVtbl -> get_TimeWork2(This,pVal)

#define IMGertNet_OptimizeSelectSF(This,otTask,pCollSF,cyMax,dDeltaQ,dP0,shNRetAlternate)	\
    (This)->lpVtbl -> OptimizeSelectSF(This,otTask,pCollSF,cyMax,dDeltaQ,dP0,shNRetAlternate)

#define IMGertNet_get_IsRunning2(This,pVal)	\
    (This)->lpVtbl -> get_IsRunning2(This,pVal)

#define IMGertNet_get_IsRunningOpt(This,pVal)	\
    (This)->lpVtbl -> get_IsRunningOpt(This,pVal)

#define IMGertNet_get_OptimizResultsGetAndClear(This,ppVal)	\
    (This)->lpVtbl -> get_OptimizResultsGetAndClear(This,ppVal)

#define IMGertNet_get_TimeWork3(This,pVal)	\
    (This)->lpVtbl -> get_TimeWork3(This,pVal)

#define IMGertNet_CalibrateModel(This,dP1,dP2,dP3,dP4)	\
    (This)->lpVtbl -> CalibrateModel(This,dP1,dP2,dP3,dP4)

#define IMGertNet_get_LastCalcError(This,pVal)	\
    (This)->lpVtbl -> get_LastCalcError(This,pVal)

#define IMGertNet_GetModelStat(This,shIdx,pNS,pparrSmpl,bFullGet)	\
    (This)->lpVtbl -> GetModelStat(This,shIdx,pNS,pparrSmpl,bFullGet)

#define IMGertNet_get_IsDirty(This,pVal)	\
    (This)->lpVtbl -> get_IsDirty(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_Enumerators_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ ICollLingvo __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_Enumerators_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][defaultcollelem][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_Factors_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ ICollFac __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_Factors_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NotifyWnd_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NotifyWnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NotifyWnd_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IMGertNet_put_NotifyWnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_RunMode_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ ModelType __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_RunMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_RunMode_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ ModelType newVal);


void __RPC_STUB IMGertNet_put_RunMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_Run_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long lNExperience,
    /* [in] */ long lNRunInOne,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL bResetModel,
    /* [defaultvalue][optional][in] */ long lRndBase,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL bPrepareOnly);


void __RPC_STUB IMGertNet_Run_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_Cancel_Proxy( 
    IMGertNet __RPC_FAR * This);


void __RPC_STUB IMGertNet_Cancel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_Counter_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ WCHAR cState,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_Counter_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_Dx_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ WCHAR cState,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_Dx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_Mx_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ WCHAR cState,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_Mx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_TimeWork_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_TimeWork_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NumberStates_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NumberStates_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_K_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_K_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_K_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IMGertNet_put_K_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_N_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_N_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_N_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IMGertNet_put_N_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_SampleN_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long lIndex,
    /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrSmpl);


void __RPC_STUB IMGertNet_get_SampleN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NotifyStep_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NotifyStep_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NotifyStep_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IMGertNet_put_NotifyStep_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_GetFactorPresent_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ IFactor __RPC_FAR *pF,
    /* [optional][out][in] */ VARIANT __RPC_FAR *strDescription,
    /* [optional][out][in] */ VARIANT __RPC_FAR *strTrustLvl,
    /* [optional][out][in] */ VARIANT __RPC_FAR *dVal);


void __RPC_STUB IMGertNet_GetFactorPresent_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_GetFactorPresentForName_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ BSTR bsShortName,
    /* [out] */ BSTR __RPC_FAR *pbsDescription,
    /* [out] */ BSTR __RPC_FAR *pbsTrustLvl,
    /* [out] */ double __RPC_FAR *pdVal);


void __RPC_STUB IMGertNet_GetFactorPresentForName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_ApplySF_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ ICollSF __RPC_FAR *pSF,
    /* [optional][out] */ VARIANT_BOOL __RPC_FAR *pbOverWrap);


void __RPC_STUB IMGertNet_ApplySF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_GetEnumForFactor_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ IFactor __RPC_FAR *pF,
    /* [retval][out] */ ILingvoEnum __RPC_FAR *__RPC_FAR *ppLE);


void __RPC_STUB IMGertNet_GetEnumForFactor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_GetInfSampleK_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ WCHAR cState,
    /* [out] */ long __RPC_FAR *pMin,
    /* [out] */ long __RPC_FAR *pMax,
    /* [out] */ double __RPC_FAR *pMx,
    /* [out] */ double __RPC_FAR *pDx);


void __RPC_STUB IMGertNet_GetInfSampleK_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_IsRunning_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_IsRunning_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_GetInfSampleKIdx_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ WCHAR cState,
    /* [out] */ double __RPC_FAR *pMin,
    /* [out] */ double __RPC_FAR *pMax,
    /* [out] */ double __RPC_FAR *pMx,
    /* [out] */ double __RPC_FAR *pDx);


void __RPC_STUB IMGertNet_GetInfSampleKIdx_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_ChangeDirty_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bFlDirty);


void __RPC_STUB IMGertNet_ChangeDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_CalcDeltaQ_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ ICollSF __RPC_FAR *pCollSF,
    /* [in] */ long lK,
    /* [in] */ long lN,
    /* [defaultvalue][optional][in] */ short shMedVal);


void __RPC_STUB IMGertNet_CalcDeltaQ_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_CalcSpeed_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ ThrdPriority __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_CalcSpeed_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_CalcSpeed_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ ThrdPriority newVal);


void __RPC_STUB IMGertNet_put_CalcSpeed_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_OptimizationType_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ OptType __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_OptimizationType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_OptimizationType_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ OptType newVal);


void __RPC_STUB IMGertNet_put_OptimizationType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_RemovLingvoOverwrap_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_RemovLingvoOverwrap_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_RemovLingvoOverwrap_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IMGertNet_put_RemovLingvoOverwrap_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_AverageDamage_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ CURRENCY __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_AverageDamage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_AverageDamage_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ CURRENCY newVal);


void __RPC_STUB IMGertNet_put_AverageDamage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_MoneyForSF_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ CURRENCY __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_MoneyForSF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_MoneyForSF_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ CURRENCY newVal);


void __RPC_STUB IMGertNet_put_MoneyForSF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_Snapshot_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bMake);


void __RPC_STUB IMGertNet_Snapshot_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_Revert_Proxy( 
    IMGertNet __RPC_FAR * This);


void __RPC_STUB IMGertNet_Revert_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_UserProp_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_UserProp_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_UserProp_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IMGertNet_put_UserProp_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_RndBase_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_RndBase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_RndBase_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IMGertNet_put_RndBase_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_CurrIterStr_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_CurrIterStr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_TotalIterStr_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_TotalIterStr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_StatOn_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ TStatFlag __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_StatOn_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_StatOn_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ TStatFlag newVal);


void __RPC_STUB IMGertNet_put_StatOn_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_StatIn_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_StatIn_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_StatIn_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IMGertNet_put_StatIn_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_TstApplyTCH_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ IFactor __RPC_FAR *pFac,
    /* [in] */ IFChange __RPC_FAR *pFC,
    /* [retval][out] */ TrustLevel __RPC_FAR *pLvl);


void __RPC_STUB IMGertNet_TstApplyTCH_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_ImitateCount_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_ImitateCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_OptCount_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_OptCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_RngCount_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_RngCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_ModuleProgID_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_ModuleProgID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_ModuleProgID_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IMGertNet_put_ModuleProgID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NDiv_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NDiv_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NDiv_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IMGertNet_put_NDiv_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_UnionThreshold_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_UnionThreshold_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_UnionThreshold_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ double newVal);


void __RPC_STUB IMGertNet_put_UnionThreshold_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_IntegralAccuracy_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_IntegralAccuracy_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_IntegralAccuracy_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ double newVal);


void __RPC_STUB IMGertNet_put_IntegralAccuracy_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NFormatPr_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ NumberFormat __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NFormatPr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NFormatPr_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ NumberFormat newVal);


void __RPC_STUB IMGertNet_put_NFormatPr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NFormatOther_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ NumberFormat __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NFormatOther_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NFormatOther_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ NumberFormat newVal);


void __RPC_STUB IMGertNet_put_NFormatOther_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NDigitsPr_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NDigitsPr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NDigitsPr_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IMGertNet_put_NDigitsPr_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NDigitsOther_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NDigitsOther_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_NDigitsOther_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short newVal);


void __RPC_STUB IMGertNet_put_NDigitsOther_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_NFormat_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ NFormatTyp nf,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_NFormat_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_VParam_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short Ventil,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_VParam_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_VParam_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short Ventil,
    /* [in] */ long newVal);


void __RPC_STUB IMGertNet_put_VParam_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_VParamIndistinct_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short Ventil,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_VParamIndistinct_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_VParamIndistinct_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short Ventil,
    /* [in] */ double newVal);


void __RPC_STUB IMGertNet_put_VParamIndistinct_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_VProbability_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short Ventil,
    /* [retval][out] */ double __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_VProbability_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_VProbability_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short Ventil,
    /* [in] */ double newVal);


void __RPC_STUB IMGertNet_put_VProbability_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_ModuleConfig_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_ModuleConfig_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IMGertNet_put_ModuleConfig_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IMGertNet_put_ModuleConfig_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_TimeWork2_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_TimeWork2_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_OptimizeSelectSF_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ OptTask otTask,
    /* [in] */ ICollSF __RPC_FAR *pCollSF,
    /* [in] */ CURRENCY cyMax,
    /* [in] */ double dDeltaQ,
    /* [in] */ double dP0,
    /* [in] */ short shNRetAlternate);


void __RPC_STUB IMGertNet_OptimizeSelectSF_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_IsRunning2_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_IsRunning2_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_IsRunningOpt_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_IsRunningOpt_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_OptimizResultsGetAndClear_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *ppVal);


void __RPC_STUB IMGertNet_get_OptimizResultsGetAndClear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_TimeWork3_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_TimeWork3_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_CalibrateModel_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ double dP1,
    /* [in] */ double dP2,
    /* [in] */ double dP3,
    /* [in] */ double dP4);


void __RPC_STUB IMGertNet_CalibrateModel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_LastCalcError_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_LastCalcError_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IMGertNet_GetModelStat_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [in] */ short shIdx,
    /* [out] */ TypANodesStat __RPC_FAR *pNS,
    /* [out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *pparrSmpl,
    /* [in] */ VARIANT_BOOL bFullGet);


void __RPC_STUB IMGertNet_GetModelStat_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IMGertNet_get_IsDirty_Proxy( 
    IMGertNet __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IMGertNet_get_IsDirty_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IMGertNet_INTERFACE_DEFINED__ */


#ifndef __IFactorAssign_INTERFACE_DEFINED__
#define __IFactorAssign_INTERFACE_DEFINED__

/* interface IFactorAssign */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IFactorAssign;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("5A049761-6C64-11d4-8FBC-00504E02C39D")
    IFactorAssign : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Description( 
            /* [retval][out] */ BSTR __RPC_FAR *Descr) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AssignFactor( 
            /* [in] */ IDispatch __RPC_FAR *OwnerForm,
            /* [in] */ long HWnd,
            /* [in] */ IMGertNet __RPC_FAR *GertNet,
            /* [in] */ IFactor __RPC_FAR *Factor,
            /* [in] */ BSTR ConfigName,
            /* [in] */ BSTR FileName) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE About( 
            /* [in] */ IDispatch __RPC_FAR *OwnerForm,
            /* [in] */ long HWnd) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ModalResult( 
            /* [retval][out] */ long __RPC_FAR *MResult) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Configure( 
            /* [in] */ IDispatch __RPC_FAR *OwnerForm,
            /* [in] */ long HWnd,
            /* [in] */ BSTR ConfigName,
            /* [in] */ BSTR FileName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IFactorAssignVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IFactorAssign __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IFactorAssign __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IFactorAssign __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Description )( 
            IFactorAssign __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *Descr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AssignFactor )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ IDispatch __RPC_FAR *OwnerForm,
            /* [in] */ long HWnd,
            /* [in] */ IMGertNet __RPC_FAR *GertNet,
            /* [in] */ IFactor __RPC_FAR *Factor,
            /* [in] */ BSTR ConfigName,
            /* [in] */ BSTR FileName);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *About )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ IDispatch __RPC_FAR *OwnerForm,
            /* [in] */ long HWnd);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ModalResult )( 
            IFactorAssign __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *MResult);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Configure )( 
            IFactorAssign __RPC_FAR * This,
            /* [in] */ IDispatch __RPC_FAR *OwnerForm,
            /* [in] */ long HWnd,
            /* [in] */ BSTR ConfigName,
            /* [in] */ BSTR FileName);
        
        END_INTERFACE
    } IFactorAssignVtbl;

    interface IFactorAssign
    {
        CONST_VTBL struct IFactorAssignVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IFactorAssign_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IFactorAssign_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IFactorAssign_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IFactorAssign_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IFactorAssign_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IFactorAssign_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IFactorAssign_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IFactorAssign_get_Description(This,Descr)	\
    (This)->lpVtbl -> get_Description(This,Descr)

#define IFactorAssign_AssignFactor(This,OwnerForm,HWnd,GertNet,Factor,ConfigName,FileName)	\
    (This)->lpVtbl -> AssignFactor(This,OwnerForm,HWnd,GertNet,Factor,ConfigName,FileName)

#define IFactorAssign_About(This,OwnerForm,HWnd)	\
    (This)->lpVtbl -> About(This,OwnerForm,HWnd)

#define IFactorAssign_get_ModalResult(This,MResult)	\
    (This)->lpVtbl -> get_ModalResult(This,MResult)

#define IFactorAssign_Configure(This,OwnerForm,HWnd,ConfigName,FileName)	\
    (This)->lpVtbl -> Configure(This,OwnerForm,HWnd,ConfigName,FileName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id][propget] */ HRESULT STDMETHODCALLTYPE IFactorAssign_get_Description_Proxy( 
    IFactorAssign __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *Descr);


void __RPC_STUB IFactorAssign_get_Description_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IFactorAssign_AssignFactor_Proxy( 
    IFactorAssign __RPC_FAR * This,
    /* [in] */ IDispatch __RPC_FAR *OwnerForm,
    /* [in] */ long HWnd,
    /* [in] */ IMGertNet __RPC_FAR *GertNet,
    /* [in] */ IFactor __RPC_FAR *Factor,
    /* [in] */ BSTR ConfigName,
    /* [in] */ BSTR FileName);


void __RPC_STUB IFactorAssign_AssignFactor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IFactorAssign_About_Proxy( 
    IFactorAssign __RPC_FAR * This,
    /* [in] */ IDispatch __RPC_FAR *OwnerForm,
    /* [in] */ long HWnd);


void __RPC_STUB IFactorAssign_About_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IFactorAssign_get_ModalResult_Proxy( 
    IFactorAssign __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *MResult);


void __RPC_STUB IFactorAssign_get_ModalResult_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IFactorAssign_Configure_Proxy( 
    IFactorAssign __RPC_FAR * This,
    /* [in] */ IDispatch __RPC_FAR *OwnerForm,
    /* [in] */ long HWnd,
    /* [in] */ BSTR ConfigName,
    /* [in] */ BSTR FileName);


void __RPC_STUB IFactorAssign_Configure_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IFactorAssign_INTERFACE_DEFINED__ */



#ifndef __GERTNETLib_LIBRARY_DEFINED__
#define __GERTNETLib_LIBRARY_DEFINED__

/* library GERTNETLib */
/* [helpstring][version][uuid] */ 




EXTERN_C const IID LIBID_GERTNETLib;

#ifndef ___IGertNetEvents_DISPINTERFACE_DEFINED__
#define ___IGertNetEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IGertNetEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IGertNetEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("EE4DF84F-9CE0-11D3-8E11-AA504E02C39A")
    _IGertNetEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IGertNetEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _IGertNetEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _IGertNetEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _IGertNetEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _IGertNetEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _IGertNetEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _IGertNetEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _IGertNetEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _IGertNetEventsVtbl;

    interface _IGertNetEvents
    {
        CONST_VTBL struct _IGertNetEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IGertNetEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IGertNetEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IGertNetEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IGertNetEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IGertNetEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IGertNetEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IGertNetEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IGertNetEvents_DISPINTERFACE_DEFINED__ */


#ifndef ___IFacEvents_DISPINTERFACE_DEFINED__
#define ___IFacEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IFacEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IFacEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("114DF84F-9CE0-91D3-8E10-9A504E02C397")
    _IFacEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IFacEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _IFacEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _IFacEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _IFacEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _IFacEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _IFacEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _IFacEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _IFacEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _IFacEventsVtbl;

    interface _IFacEvents
    {
        CONST_VTBL struct _IFacEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IFacEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IFacEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IFacEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IFacEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IFacEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IFacEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IFacEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IFacEvents_DISPINTERFACE_DEFINED__ */


#ifndef __ICatRegistrar_INTERFACE_DEFINED__
#define __ICatRegistrar_INTERFACE_DEFINED__

/* interface ICatRegistrar */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICatRegistrar;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("840E76B1-6ADC-11D4-8FBA-00504E02C39D")
    ICatRegistrar : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Register( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Unregister( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetProgIDs( 
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *arrProgIDs) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICatRegistrarVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICatRegistrar __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICatRegistrar __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICatRegistrar __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICatRegistrar __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICatRegistrar __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICatRegistrar __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICatRegistrar __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Register )( 
            ICatRegistrar __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Unregister )( 
            ICatRegistrar __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetProgIDs )( 
            ICatRegistrar __RPC_FAR * This,
            /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *arrProgIDs);
        
        END_INTERFACE
    } ICatRegistrarVtbl;

    interface ICatRegistrar
    {
        CONST_VTBL struct ICatRegistrarVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICatRegistrar_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICatRegistrar_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICatRegistrar_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICatRegistrar_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICatRegistrar_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICatRegistrar_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICatRegistrar_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICatRegistrar_Register(This)	\
    (This)->lpVtbl -> Register(This)

#define ICatRegistrar_Unregister(This)	\
    (This)->lpVtbl -> Unregister(This)

#define ICatRegistrar_GetProgIDs(This,arrProgIDs)	\
    (This)->lpVtbl -> GetProgIDs(This,arrProgIDs)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICatRegistrar_Register_Proxy( 
    ICatRegistrar __RPC_FAR * This);


void __RPC_STUB ICatRegistrar_Register_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICatRegistrar_Unregister_Proxy( 
    ICatRegistrar __RPC_FAR * This);


void __RPC_STUB ICatRegistrar_Unregister_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICatRegistrar_GetProgIDs_Proxy( 
    ICatRegistrar __RPC_FAR * This,
    /* [retval][out] */ SAFEARRAY __RPC_FAR * __RPC_FAR *arrProgIDs);


void __RPC_STUB ICatRegistrar_GetProgIDs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICatRegistrar_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_LingvoEnum;

#ifdef __cplusplus

class DECLSPEC_UUID("86822111-0876-11D4-8EE4-00504E02C39D")
LingvoEnum;
#endif

EXTERN_C const CLSID CLSID_CollLingvo;

#ifdef __cplusplus

class DECLSPEC_UUID("86822116-0876-11D4-8EE4-00504E02C39D")
CollLingvo;
#endif

EXTERN_C const CLSID CLSID_CompaundFiles;

#ifdef __cplusplus

class DECLSPEC_UUID("48118E78-0F7C-11D4-8EF1-00504E02C39D")
CompaundFiles;
#endif

EXTERN_C const CLSID CLSID_Factor;

#ifdef __cplusplus

class DECLSPEC_UUID("A754AAC4-1117-11D4-8EF2-00504E02C39D")
Factor;
#endif

EXTERN_C const CLSID CLSID_CollFac;

#ifdef __cplusplus

class DECLSPEC_UUID("A754AAC6-1117-11D4-8EF2-00504E02C39D")
CollFac;
#endif

EXTERN_C const CLSID CLSID_MGertNet;

#ifdef __cplusplus

class DECLSPEC_UUID("A754AAC8-1117-11D4-8EF2-00504E02C39D")
MGertNet;
#endif

EXTERN_C const CLSID CLSID_SafetyPrecaution;

#ifdef __cplusplus

class DECLSPEC_UUID("F4E48354-15DF-11D4-8EFC-00504E02C39D")
SafetyPrecaution;
#endif

EXTERN_C const CLSID CLSID_FChange;

#ifdef __cplusplus

class DECLSPEC_UUID("0482EBD5-15F5-11D4-8EFD-00504E02C39D")
FChange;
#endif

EXTERN_C const CLSID CLSID_ICollFChange;

#ifdef __cplusplus

class DECLSPEC_UUID("0482EBD7-15F5-11D4-8EFD-00504E02C39D")
ICollFChange;
#endif

EXTERN_C const CLSID CLSID_CollSF;

#ifdef __cplusplus

class DECLSPEC_UUID("88B55014-176A-11D4-8EFE-00504E02C39D")
CollSF;
#endif

EXTERN_C const CLSID CLSID_CatRegistrar;

#ifdef __cplusplus

class DECLSPEC_UUID("840E76B2-6ADC-11D4-8FBA-00504E02C39D")
CatRegistrar;
#endif
#endif /* __GERTNETLib_LIBRARY_DEFINED__ */

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
