/* tlb
   https://msdn.microsoft.com/en-us/library/windows/desktop/ms221342(v=vs.85).aspx
   https://msdn.microsoft.com/en-us/library/windows/desktop/ms221442(v=vs.85).aspx
   IDL: https://msdn.microsoft.com/en-us/library/windows/desktop/aa367061%28v=vs.85%29.aspx
*/

#ifndef _LIBTLB_
#define _LIBTLB_

#include <stdio.h>
#include <stdlib.h>
#include <tchar.h>
#include <windows.h>
#include <ole2.h>

// typeinfo type
TCHAR *tlikind[] = {
    _T("enum"),
    _T("struct"),   //record
    _T("module"),
    _T("interface"),
    _T("dispinterface"),
    _T("coclass"),
    _T("typedef"),  //alias
    _T("union"),
    NULL
};

typedef struct _USTC_Array {
    USHORT  v;
    TCHAR   *s;
} USTC_Array;

// var
USTC_Array varflags_array[] = {
    {VARFLAG_FREADONLY, _T("readonly")},
    {VARFLAG_FSOURCE,   _T("source")},
    {VARFLAG_FBINDABLE,     _T("bindable")},
    {VARFLAG_FREQUESTEDIT,  _T("requestedit")},
    {VARFLAG_FDISPLAYBIND,  _T("displaybind")},
    {VARFLAG_FDEFAULTBIND,  _T("defaultbind")},
    {VARFLAG_FHIDDEN,       _T("hidden")},
    {VARFLAG_FRESTRICTED,       _T("restricted")},
    {VARFLAG_FDEFAULTCOLLELEM,  _T("defaultcollelem")},
    {VARFLAG_FUIDEFAULT,        _T("uidefault")},
    {VARFLAG_FNONBROWSABLE,     _T("nonbrowsable")},
    {VARFLAG_FREPLACEABLE,      _T("replaceable")},
    {VARFLAG_FIMMEDIATEBIND,    _T("immediatebind")}
};

USTC_Array varkind_array[] = {
    {VAR_PERINSTANCE,   _T("perinstance")},
    {VAR_STATIC,        _T("static")},
    {VAR_CONST,     _T("const")},
    {VAR_DISPATCH,  _T("dispatch")}
};

// function
USTC_Array funcflags_array[] = {
    {FUNCFLAG_FRESTRICTED,  _T("restricted")},
    {FUNCFLAG_FSOURCE,      _T("source")},
    {FUNCFLAG_FBINDABLE,    _T("bindable")},
    {FUNCFLAG_FREQUESTEDIT, _T("requestedit")},
    {FUNCFLAG_FDISPLAYBIND, _T("displaybind")},
    {FUNCFLAG_FHIDDEN,      _T("hidden")},
    {FUNCFLAG_FUSESGETLASTERROR,_T("usesgetlasterror")},
    {FUNCFLAG_FDEFAULTCOLLELEM, _T("defaultcollelem")},
    {FUNCFLAG_FUIDEFAULT,       _T("uidefault")},
    {FUNCFLAG_FNONBROWSABLE,    _T("nonbrowsable")},
    {FUNCFLAG_FREPLACEABLE,     _T("replaceable")},
    {FUNCFLAG_FIMMEDIATEBIND,   _T("immediatebind")},
};

USTC_Array invkind_array[] = {
    {INVOKE_FUNC,   _T("func")},
    {INVOKE_PROPERTYGET,    _T("propget")},
    {INVOKE_PROPERTYPUT,    _T("propput")},
    {INVOKE_PROPERTYPUTREF, _T("propputref")}
};

USTC_Array paramflags_array[] = {
    {PARAMFLAG_FIN,     _T("in")},
    {PARAMFLAG_FOUT,    _T("out")},
    {PARAMFLAG_FLCID,   _T("lcid")},
    {PARAMFLAG_FRETVAL, _T("retval")},
    {PARAMFLAG_FOPT,    _T("opt")},
    {PARAMFLAG_FHASDEFAULT, _T("hasdefault")},
    {0x40,                  _T("hascustdata")}
};

USTC_Array vt_array[] = {
    // enum VARENUM, https://msdn.microsoft.com/en-us/library/windows/desktop/ms221170%28v=vs.85%29.aspx
    {VT_EMPTY,  _T("")},
    {VT_NULL,   _T("NULL")},
    {VT_I2,     _T("short")},
    {VT_I4,     _T("long")},
    {VT_R4,     _T("float")},
    {VT_R8,     _T("double")},
    {VT_CY,     _T("Currency")},
    {VT_DATE,   _T("DATE")},
    {VT_BSTR,   _T("BSTR")},
    {VT_DISPATCH, _T("IDispatch*")},
    {VT_ERROR,  _T("VT_ERROR")},
    {VT_BOOL,   _T("BOOL")},
    {VT_VARIANT,_T("VARIANT")},
    {VT_UNKNOWN,  _T("IUnknown*")},
    {VT_DECIMAL,_T("DECIMAL")},
    {VT_I1,     _T("char")},
    {VT_UI1,    _T("unsigned char")},
    {VT_UI2,    _T("unsigned short")},
    {VT_UI4,    _T("unsigned long")},
    {VT_I8,     _T("long long")},
    {VT_UI8,    _T("unsigned long long")},
    {VT_INT,    _T("int")},
    {VT_UINT,   _T("unsigned int")},
    {VT_VOID,   _T("void")},
    {VT_HRESULT,    _T("HRESULT")},
    {VT_PTR,        _T("void*")},
    {VT_SAFEARRAY,  _T("SAFEARRAY")},
    {VT_CARRAY,     _T("VT_CARRAY")},
    {VT_USERDEFINED,_T("VT_USERDEFINED")},
    {VT_LPSTR,      _T("LPSTR")},
    {VT_LPWSTR,     _T("LPWSTR")},
    {VT_RECORD,     _T("VT_RECORD")},
    {VT_INT_PTR,    _T("INT_PTR")},
    {VT_UINT_PTR,   _T("UINT_PTR")},
    {VT_FILETIME,   _T("FILETIME")},
    {VT_BLOB,       _T("BLOB")},
    {VT_STREAM,     _T("VT_STREAM")},
    {VT_STORAGE,    _T("VT_STORAGE")},
    {VT_STREAMED_OBJECT,_T("VT_STREAMED_OBJECT")},
    {VT_STORED_OBJECT,  _T("VT_STORED_OBJECT")},
    {VT_BLOB_OBJECT,    _T("VT_BLOB_OBJECT")},
    {VT_CF,             _T("VT_CF")},
    {VT_CLSID,          _T("CLSID")},
    //{VT_VERSIONED_STREAM, _T("VT_VERSIONED_STREAM")},
    {VT_BSTR_BLOB,  _T("VT_BSTR_BLOB")},
    {VT_VECTOR,     _T("VT_VECTOR")},
    {VT_ARRAY,      _T("VT_ARRAY")},
    {VT_BYREF,      _T("VT_BYREF")}
};

//
int check_result( HRESULT ret ) {
    if (ret != S_OK) {
        if ( GetStdHandle(STD_OUTPUT_HANDLE) != NULL) { _tprintf(_T("Error 0x%X\n"), ret); }
        else {}
        exit(1);
    }
}

int parse_CLSID( CLSID *rclsid, int CALLBACK (*callback)(TCHAR *) ) {
    LPOLESTR *lplpsz;
    //
    check_result( StringFromCLSID(rclsid, lplpsz) );
    callback(*lplpsz);
    CoTaskMemFree(*lplpsz);
}

int parse_kindtype2str( USTC_Array *tk_a, UINT n, UINT k, int CALLBACK (*callback)(TCHAR *) ) {
    UINT i;
    for (i = 0; i < n; i++) {
        if (k == tk_a[i].v) {
            callback(tk_a[i].s);
            break;
        }
    }
}

int parse_flags2str( USTC_Array *flag_a, UINT n, USHORT flag, int CALLBACK (*callback)(TCHAR *, int) ) {
    int i, found;
    //
    found = 0;
    for (i = 0; i < n; i++) {
        if ((flag & flag_a[i].v) == flag_a[i].v) {
            found++;
            callback( flag_a[i].s, found );
        }
    }
    if (found > 0) {
        callback( _T(""), 0 );
    }
}

// var of emun/struct/typedef
int parse_varflags( USHORT flag, int CALLBACK (*callback)(TCHAR *, int) ) {
    parse_flags2str( varflags_array, sizeof(varflags_array)/sizeof(USTC_Array), flag, callback );
}

int parse_varkind( UINT k, int CALLBACK (*callback)(TCHAR *) ) {
    parse_kindtype2str(varkind_array, sizeof(varkind_array)/sizeof(USTC_Array), k, callback);
}

int parse_vardesc( ITypeInfo *pTinfo, UINT count, int CALLBACK (*callback)(ITypeInfo *, VARDESC *) ) {
    VARDESC    *pVarDesc;
    UINT        i;
    //
    for (i = 0; i < count; i++) {
        check_result( pTinfo->lpVtbl->GetVarDesc(pTinfo, i, &pVarDesc) );
        callback(pTinfo, pVarDesc);
        pTinfo->lpVtbl->ReleaseVarDesc(pTinfo, pVarDesc);
    }
}

// function of interface/dispinterface/coclass
int parse_funcflags( USHORT flag, int CALLBACK (*callback)(TCHAR *, int) ) {
    parse_flags2str( funcflags_array, sizeof(funcflags_array)/sizeof(USTC_Array), flag, callback );
}

int parse_invkind( UINT k, int CALLBACK (*callback)(TCHAR *) ) {
    parse_kindtype2str(invkind_array, sizeof(invkind_array)/sizeof(USTC_Array), k, callback);
}

int parse_vt( UINT k, int CALLBACK (*callback)(TCHAR *) ) {
    parse_kindtype2str(vt_array, sizeof(vt_array)/sizeof(USTC_Array), k, callback);
}

int parse_paramflags( USHORT flag, int CALLBACK (*callback)(TCHAR *, int) ) {
    if (flag == PARAMFLAG_NONE) { return; }
    parse_flags2str( paramflags_array, sizeof(paramflags_array)/sizeof(USTC_Array), flag, callback );
}

int parse_funcdesc( ITypeInfo *pTinfo, UINT count, int CALLBACK (*callback)(ITypeInfo *, FUNCDESC *) ) {
    FUNCDESC    *pFuncDesc;
    UINT        i;
    //
    for (i = 0; i < count; i++) {
        check_result( pTinfo->lpVtbl->GetFuncDesc(pTinfo, i, &pFuncDesc) );
        callback(pTinfo, pFuncDesc);
        pTinfo->lpVtbl->ReleaseFuncDesc(pTinfo, pFuncDesc);
    }
}

// typeinfo
int get_tli_name( ITypeInfo *pTinfo, int CALLBACK (*callback)(TCHAR *) ) {
    BSTR bStr;
    check_result( pTinfo->lpVtbl->GetDocumentation(pTinfo, MEMBERID_NIL, &bStr, NULL, NULL, NULL) );
    callback( bStr );
    SysFreeString(bStr);
}

int get_tli_help( ITypeInfo *pTinfo, MEMBERID memid, int CALLBACK (*callback)(TCHAR *) ) {
    BSTR bStr;
    check_result( pTinfo->lpVtbl->GetDocumentation(pTinfo, memid, NULL, &bStr, NULL, NULL) );
    callback( bStr );
    SysFreeString(bStr);
}

int get_inherit_tli( ITypeInfo *pTinfo, ITypeInfo **ppTinfo_ih ) {
    HREFTYPE    HRefType;
    //
    check_result( pTinfo->lpVtbl->GetRefTypeOfImplType(pTinfo, -1, &HRefType) );
    check_result( pTinfo->lpVtbl->GetRefTypeInfo(pTinfo, HRefType, ppTinfo_ih) );
}

int process_tli_by_kind( ITypeLib *ptlib, TYPEKIND kind, int CALLBACK (*callback)(ITypeInfo*, TYPEATTR *) ) {
    UINT        NumTI;
    UINT        i;
    //
    NumTI = ptlib->lpVtbl->GetTypeInfoCount(ptlib);
    //
    for (i = 0; i < NumTI; i++) {
        TYPEKIND    TKind;
        TYPEATTR    *pTA;
        ITypeInfo   *pTinfo;
        UINT        tk_can_dual;
        //
        tk_can_dual = 0;
        ptlib->lpVtbl->GetTypeInfoType(ptlib, i, &TKind);
        if ( TKind != kind ) {
            if ( kind == TKIND_INTERFACE && TKind == TKIND_DISPATCH ) { tk_can_dual = 1; }
            else { continue; }
        }
        //
        check_result( ptlib->lpVtbl->GetTypeInfo(ptlib, i, &pTinfo) );
        check_result( pTinfo->lpVtbl->GetTypeAttr(pTinfo, &pTA) );
        //
        if (tk_can_dual) {
            if ((pTA->wTypeFlags & TYPEFLAG_FDUAL) == TYPEFLAG_FDUAL) {
                get_inherit_tli( pTinfo, &pTinfo );
                check_result( pTinfo->lpVtbl->GetTypeAttr(pTinfo, &pTA) );
            }
            else { continue; }
        }
        //
        callback( pTinfo, pTA );
        //
        pTinfo->lpVtbl->ReleaseTypeAttr(pTinfo, pTA);
    }
    return TRUE;
}

// tlb
int get_tlb_name( ITypeLib *ptlib, int CALLBACK (*callback)(TCHAR *) ) {
    BSTR bStr;
    check_result( ptlib->lpVtbl->GetDocumentation(ptlib, MEMBERID_NIL, &bStr, NULL, NULL, NULL) );
    callback( bStr );
    SysFreeString(bStr);
}

int parse_tlb( const TCHAR *pf, int CALLBACK (*callback)(ITypeLib *) ) {
    ITypeLib    *ptlib;
    //
    check_result( LoadTypeLib( pf, &ptlib ) );
    callback( ptlib );
}
// ITypeLib2::GetAllCustData, ITypeInfo2::GetAllCustData, ClearCustData

#endif
