/* tlb2mm
*/

#define _UNICODE
#define UNICODE

#define NONAMELESSUNION

#include "libtlb.c"
#include <shlwapi.h>

//
TCHAR   *trunk_pos;
TCHAR   *pos_l = _T("left");
TCHAR   *pos_r = _T("right");
UINT    fold_node = 0;
TCHAR   *file_name;

//
int gen_node_property_start( TCHAR *s  ) {
    _tprintf(_T(" %s=\""), s);
}

int gen_node_property_end() {
    _tprintf(_T("\""));
}

int gen_node_start_start() {
    _tprintf(_T("<node"));
}

int gen_node_start_end() {
    _tprintf(_T(">\n"));
}

int gen_node_position_start() {
    gen_node_property_start(_T("POSITION"));
}

int gen_node_text_start() {
    gen_node_property_start(_T("TEXT"));
}

int gen_node_folded() {
    gen_node_property_start(_T("FOLDED"));
    _tprintf(_T("true"));
    gen_node_property_end();
}

int gen_node_end() {
    _tprintf(_T("</node>\n"));
}

int CALLBACK gen_trunk_node_start( TCHAR *txt ) {
    gen_node_start_start();
    //
    if (fold_node) { gen_node_folded(); }
    //
    gen_node_position_start();
    _tprintf(_T("%s"), trunk_pos);
    gen_node_property_end();
    //
    gen_node_text_start();
    if (txt) {_tprintf(_T("%s"), txt);}
    gen_node_property_end();
    //
    gen_node_start_end();
}

int CALLBACK gen_node_start( TCHAR *txt ) {
    gen_node_start_start();
    //
    if (fold_node) { gen_node_folded(); }
    //
    gen_node_text_start();
    if (txt) {_tprintf(_T("%s"), txt);}
    gen_node_property_end();
    //
    gen_node_start_end();
}

//
int CALLBACK gen_flags( TCHAR *s, int found ) {
    if (found == 1) { _tprintf(_T("[%s"), s); }
    else if (found > 1) { _tprintf(_T(", %s"), s); }
    else { _tprintf(_T("] ")); }
}

int CALLBACK gen_kind( TCHAR *s ) {
    _tprintf(_T("[%s] "), s);
}

int CALLBACK gen_vartype( TCHAR *s ) {
    _tprintf(_T("%s "), s);
}

int CALLBACK gen_BSTR(TCHAR *s) {
    if (s != NULL) {_tprintf(_T("%s"), s);}
}

int gen_variant( VARTYPE vt, void *pVal ) {
    switch (vt) {
    case VT_I1:
        _tprintf(_T("%d"), *(char*)pVal);
        break;
    case VT_I2:
        _tprintf(_T("%d"), *(short*)pVal);
        break;
    case VT_I4:
        _tprintf(_T("%d"), *(long*)pVal);
        break;
    case VT_I8:
        _tprintf(_T("%lld"), *(long long*)pVal);
        break;
    case VT_INT:
        _tprintf(_T("%d"), *(int*)pVal);
        break;
    case VT_UI1:
        _tprintf(_T("%d"), *(unsigned char*)pVal);
        break;
    case VT_UI2:
        _tprintf(_T("%d"), *(unsigned short*)pVal);
        break;
    case VT_UI4:
        _tprintf(_T("%d"), *(unsigned long*)pVal);
        break;
    case VT_UI8:
        _tprintf(_T("%llu"), *(unsigned long long*)pVal);
        break;
    case VT_UINT:
        _tprintf(_T("%u"), *(unsigned int*)pVal);
        break;
    case VT_BOOL:
        _tprintf(_T("%d"), *(BOOL*)pVal);
        break;
    case VT_R4:
        _tprintf(_T("%f"), *(float*)pVal);
        break;
    case VT_R8:
        _tprintf(_T("%f"), *(double*)pVal);
        break;
    case VT_CY:
        _tprintf(_T("%f"), *(CY*)pVal);
        break;
    case VT_DECIMAL:
        _tprintf(_T("%f"), *(DECIMAL*)pVal);
        break;
    case VT_HRESULT:
        _tprintf(_T("0x%X"), *(HRESULT*)pVal);
        break;
    case VT_BSTR:
        _tprintf(_T("&quot;%s&quot;"), *(BSTR*)pVal);
        break;
    case VT_LPSTR:
        _tprintf(_T("&quot;%s&quot;"), *(LPSTR*)pVal); //?
        break;
    case VT_LPWSTR:
        _tprintf(_T("&quot;%s&quot;"), (LPWSTR)pVal);
        break;
    case VT_CLSID:
        parse_CLSID( (CLSID*)pVal, gen_BSTR );
        break;
    case VT_USERDEFINED:
        //
        break;
    default:
        break;
    }
    /*_tprintf(_T(" [0x%X, 0x%p]"), vt, pVal);*/
}

// var
CALLBACK gen_vars( ITypeInfo *pTinfo, VARDESC *pVarDesc ) {
    BSTR        *pbNames;
    UINT        n;
    VARTYPE     vt;
    //
    vt = pVarDesc->elemdescVar.tdesc.vt;
    //
    gen_node_start_start();
    gen_node_text_start();
    //
    parse_varflags(pVarDesc->wVarFlags, gen_flags);
    parse_varkind(pVarDesc->varkind, gen_kind);
    parse_vt(vt, gen_vartype);
    //
    pbNames = calloc(1, sizeof BSTR); //Must!
    check_result( pTinfo->lpVtbl->GetNames(pTinfo, pVarDesc->memid, pbNames, 1, &n) );
    _tprintf(_T("%s"), pbNames[0]);
    if ( pVarDesc->varkind == VAR_CONST ) {
        _tprintf(_T(" = "));
        gen_variant(vt, &V_I1(pVarDesc->lpvarValue));
    }
    else if ( pVarDesc->varkind == VAR_PERINSTANCE ) {
        //
    }
    //
    gen_node_property_end();
    gen_node_start_end();
    //
    //
    gen_node_end();
}

// function
int CALLBACK gen_functions( ITypeInfo *pTinfo, FUNCDESC *pFuncDesc ) {
    BSTR        *pbNames;
    UINT        n;
    UINT        j;
    //
    gen_node_start_start();
    gen_node_text_start();
    //
    parse_funcflags(pFuncDesc->wFuncFlags, gen_flags);
    parse_invkind(pFuncDesc->invkind, gen_kind);
    parse_vt(pFuncDesc->elemdescFunc.tdesc.vt, gen_vartype);
    //
    n = 1 + pFuncDesc->cParams;
    pbNames = calloc(n, sizeof BSTR); //Must!
    check_result( pTinfo->lpVtbl->GetNames(pTinfo, pFuncDesc->memid, pbNames, n, &n) );
    _tprintf(_T("%s("), pbNames[0]);
    if (pFuncDesc->cParams == 1) {n = 2;} //4 nameless param
    for (j = 1; j < n; j++) {
        ELEMDESC    *pElemDesc;
        PARAMDESC   *pParamDesc;
        VARTYPE     vt;
        // 1 more names
        pElemDesc = &(pFuncDesc->lprgelemdescParam[j-1]);
        pParamDesc = &(pElemDesc->DUMMYUNIONNAME.paramdesc);
        vt = pElemDesc->tdesc.vt;
        //
        if (j > 1) {_tprintf(_T(", "));}
        //
        parse_paramflags( pParamDesc->wParamFlags, gen_flags );
        parse_vt(vt, gen_vartype);
        _tprintf(_T("%s"), pbNames[j]?pbNames[j]:_T("rhs"));
        //others?
        if ((pParamDesc->wParamFlags & PARAMFLAG_FHASDEFAULT) == PARAMFLAG_FHASDEFAULT) {
            _tprintf(_T(" [="));
            gen_variant( vt, &V_I1(&(pParamDesc->pparamdescex->varDefaultValue)) );
            _tprintf(_T("]"));
        }
    }
    _tprintf(_T(")"));
    gen_node_property_end();
    gen_node_start_end();
    //
    gen_node_start_start();
    gen_node_text_start();
    get_tli_help(pTinfo, pFuncDesc->memid, gen_BSTR);
    gen_node_property_end();
    gen_node_start_end();
    gen_node_end();
    //
    gen_node_end();
    //
    SysFreeString(pbNames[0]);
}

int gen_inherit_tli_name( ITypeInfo *pTinfo ) {
    ITypeInfo   *pTinfo2;
    TYPEATTR    *pTA2;
    //
    get_inherit_tli( pTinfo, &pTinfo2 );
    check_result( pTinfo2->lpVtbl->GetTypeAttr(pTinfo2, &pTA2) );
    //
    gen_node_start_start();
    gen_node_folded();
    gen_node_text_start();
    //
    _tprintf(_T("inherited %s: "), tlikind[pTA2->typekind]);
    get_tli_name( pTinfo2, gen_BSTR);
    gen_node_property_end();
    gen_node_start_end();
    //
    gen_node_end();
    //
    pTinfo2->lpVtbl->ReleaseTypeAttr(pTinfo2, pTA2);
}

int list_interfaces( ITypeInfo *pTinfo, TYPEATTR *pTA ) {
    UINT        i;
    //
    for (i = 0; i < pTA->cImplTypes; i++) {
        HREFTYPE    HRefType;
        ITypeInfo   *pTinfo2;
        TYPEATTR    *pTA2;
        check_result( pTinfo->lpVtbl->GetRefTypeOfImplType(pTinfo, i, &HRefType) );
        check_result( pTinfo->lpVtbl->GetRefTypeInfo(pTinfo, HRefType, &pTinfo2) );
        check_result( pTinfo2->lpVtbl->GetTypeAttr(pTinfo2, &pTA2) );
        gen_node_start_start();
        gen_node_text_start();
        _tprintf(_T("%s: "), tlikind[pTA2->typekind]);
        get_tli_name( pTinfo2, gen_BSTR);
        gen_node_property_end();
        gen_node_start_end();
        gen_node_end();
        pTinfo2->lpVtbl->ReleaseTypeAttr(pTinfo2, pTA2);
    }
}

// typeinfo
int CALLBACK ShowTLI( ITypeInfo *pTinfo, TYPEATTR *pTA ) {
    //
    get_tli_name(pTinfo, gen_trunk_node_start);
    //help as sub-node
    gen_node_start(_T("help"));
    get_tli_help(pTinfo, MEMBERID_NIL, gen_node_start);
    gen_node_end();
    gen_node_end();
    //
    switch (pTA->typekind) {
    case TKIND_MODULE:
    case TKIND_INTERFACE:
    case TKIND_DISPATCH:
        //func
        if (pTA->typekind == TKIND_DISPATCH) {
            if ((pTA->wTypeFlags & TYPEFLAG_FDUAL) == TYPEFLAG_FDUAL) {
                gen_inherit_tli_name(pTinfo);
            }
        }
        else if (pTA->typekind == TKIND_INTERFACE) {
            gen_node_start_start();
            gen_node_text_start();
            _tprintf(_T("inherited: %d"), pTA->cImplTypes);
            gen_node_property_end();
            gen_node_start_end();
            //
            list_interfaces( pTinfo, pTA );
            //
            gen_node_end();
        }
        gen_node_start_start();
        gen_node_folded();
        gen_node_text_start();
        //
        _tprintf(_T("func/prop #: %d"), pTA->cFuncs);
        gen_node_property_end();
        gen_node_start_end();
        //
        if (pTA->typekind != TKIND_DISPATCH
        || (pTA->wTypeFlags & TYPEFLAG_FDUAL) != TYPEFLAG_FDUAL) {
            parse_funcdesc( pTinfo, pTA->cFuncs, gen_functions );
        }
        //
        gen_node_end();
        //property!
        gen_node_start_start();
        gen_node_folded();
        gen_node_text_start();
        _tprintf(_T("property #: %d"), pTA->cVars);
        gen_node_property_end();
        gen_node_start_end();
        //
        parse_vardesc( pTinfo, pTA->cVars, gen_vars );
        //
        gen_node_end();
        break;
    case TKIND_COCLASS:
        //
        gen_node_start_start();
        gen_node_folded();
        gen_node_text_start();
        _tprintf(_T("(disp)interface #: %d"), pTA->cImplTypes);
        gen_node_property_end();
        gen_node_start_end();
        //
        list_interfaces( pTinfo, pTA );
        //
        gen_node_end();
        break;
    case TKIND_ALIAS:
        gen_node_start_start();
        gen_node_folded();
        gen_node_text_start();
        parse_vt( pTA->tdescAlias.vt, gen_BSTR );
        gen_node_property_end();
        gen_node_start_end();
        //
        //
        gen_node_end();
        break;
    case TKIND_ENUM:
    case TKIND_UNION:
    case TKIND_RECORD:
        gen_node_start_start();
        gen_node_folded();
        gen_node_text_start();
        _tprintf(_T("# : %d"), pTA->cVars);
        gen_node_property_end();
        gen_node_start_end();
        //
        parse_vardesc( pTinfo, pTA->cVars, gen_vars );
        //
        gen_node_end();
        break;
    default:
        /*_tprintf(_T("vars # : %d\n"), pTA->cVars);*/
        break;
    }
    //
    gen_node_end();
}

int gen_tli_by_kind( ITypeLib *ptlib, TYPEKIND kind ) {
    gen_trunk_node_start( tlikind[kind] );
    process_tli_by_kind( ptlib, kind, ShowTLI );
    gen_node_end();
}

int CALLBACK gen_tlb2mm( ITypeLib *ptlib ) {
    //
    fold_node = ptlib->lpVtbl->GetTypeInfoCount(ptlib) > 50;
    //
    _tprintf(_T("<map>\n"));
    //
    gen_node_start_start();
    gen_node_text_start();
    _tprintf(_T("[%s] "), file_name);
    get_tlb_name(ptlib, gen_BSTR);
    gen_node_property_end();
    gen_node_start_end();
    //
    trunk_pos = pos_l;
    gen_tli_by_kind( ptlib, TKIND_INTERFACE );
    gen_tli_by_kind( ptlib, TKIND_DISPATCH );
    gen_tli_by_kind( ptlib, TKIND_MODULE );
    //
    trunk_pos = pos_r;
    gen_tli_by_kind( ptlib, TKIND_ALIAS );
    gen_tli_by_kind( ptlib, TKIND_ENUM );
    gen_tli_by_kind( ptlib, TKIND_RECORD );
    gen_tli_by_kind( ptlib, TKIND_UNION );
    gen_tli_by_kind( ptlib, TKIND_COCLASS );
    //
    gen_node_end();
    //
    _tprintf(_T("</map>\n"));
}

//
int _tmain(int argc, const TCHAR *argv[])
{
    if (argc <2) {
        _tprintf(_T("Usage: %s <dll-or-tlb-file> [> <out.mm>]\n"), argv[0]);
        return 0;
    }
    file_name = PathFindFileName(argv[1]);
    parse_tlb( (OLECHAR*)argv[1], gen_tlb2mm );
    //
    return 0;
}
