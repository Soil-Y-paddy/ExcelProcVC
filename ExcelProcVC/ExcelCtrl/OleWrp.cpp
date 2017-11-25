#include "OleWrp.h"

OleWrap::OleWrapError OleWrap::m_eErrorState = OleWrap::OleWrapNoError;
HRESULT OleWrap::m_nResult = S_OK;

HRESULT OleWrap::Invoker( int p_nAutoType, VARIANT * p_pVResult, IDispatch * p_pDisp, LPOLESTR p_ptName, int p_nCArgs, VARIANT * p_aryArgs )
{
	// エラー制御
	HRESULT nRetVal = S_OK;
	OleWrap::m_eErrorState = OleWrapNoError;

	// Variables used...
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;


	// Build DISPPARAMS
	dp.cArgs = p_nCArgs;
	dp.rgvarg = p_aryArgs;

	// Handle special-case for property-puts!
	if( p_nAutoType & DISPATCH_PROPERTYPUT )
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}


	if( !p_pDisp )
	{
		OleWrap::m_eErrorState = OleWrapNullPointer;
		nRetVal = E_POINTER;

	}
	// Get DISPID for name passed...
	else if(
		FAILED(
			nRetVal = p_pDisp->GetIDsOfNames( IID_NULL, &p_ptName, 1, LOCALE_USER_DEFAULT, &dispID )
		)
	)
	{
		OleWrap::m_eErrorState = OleWrapGetIdOfName;
	}
	// Make the call!
	else if(
		FAILED(
			nRetVal = p_pDisp->Invoke( dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, p_nAutoType, &dp, p_pVResult, NULL, NULL )
		)
	)
	{
		OleWrap::m_eErrorState = OleWrapInvoke;
	}

	return nRetVal;

}

IDispatch * OleWrap::getObject( IDispatch * p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... )
{
	VARIANT vResult;
	VariantInit( &vResult );

	// Begin variable-argument list...
	va_list marker;
	va_start( marker, p_nCArgs );

	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[ p_nCArgs + 1 ];
	// Extract arguments...
	for( int i = 0; i < p_nCArgs; i++ )
	{
		pArgs[ i ] = va_arg( marker, VARIANT );
	}
	OleWrap::m_nResult =  Invoker( DISPATCH_PROPERTYGET, &vResult, p_pDisp, p_ptName, p_nCArgs, pArgs );
	// End variable-argument section...
	va_end( marker );

	delete[] pArgs;

	return vResult.pdispVal;
}

VARIANT OleWrap::getValue( IDispatch * p_pDisp, LPOLESTR p_ptName )
{
	VARIANT vResult;
	VariantInit( &vResult );

	Invoker( DISPATCH_PROPERTYGET, &vResult, p_pDisp, p_ptName, 0,NULL);

	return vResult;
}

void OleWrap::setValue( IDispatch * p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... )
{
	// Begin variable-argument list...
	va_list marker;
	va_start( marker, p_nCArgs );

	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[ p_nCArgs + 1 ];
	// Extract arguments...
	for( int i = 0; i < p_nCArgs; i++ )
	{
		pArgs[ i ] = va_arg( marker, VARIANT );
	}
	OleWrap::m_nResult = Invoker( DISPATCH_PROPERTYPUT, NULL, p_pDisp, p_ptName, p_nCArgs, pArgs );

	// End variable-argument section...
	va_end( marker );
	delete[] pArgs;

}

VARIANT OleWrap::execMethod( IDispatch * p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... )
{
	VARIANT vResult;
	VariantInit( &vResult );

	// Begin variable-argument list...
	va_list marker;
	va_start( marker, p_nCArgs );

	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[ p_nCArgs + 1 ];
	// Extract arguments...
	for( int i = 0; i < p_nCArgs; i++ )
	{
		pArgs[ i ] = va_arg( marker, VARIANT );
	}
	Invoker( DISPATCH_METHOD, &vResult, p_pDisp, p_ptName, p_nCArgs, pArgs );

	// End variable-argument section...
	va_end( marker );
	delete[] pArgs;

	return vResult;
}

IDispatch * OleWrap::getInstance( LPOLESTR p_ptName )
{

	HRESULT nHresult = S_OK;
	OleWrap::OleWrapError eErrorState = OleWrapNoError;
	CLSID clsid;
	IDispatch *objRetVal = NULL;

	if( FAILED( nHresult = CLSIDFromProgID( p_ptName, &clsid ) ) )
	{
		// Clsの取得に失敗
		eErrorState = OleWrapClsIdNotExist;
	}
	else if( FAILED(
		nHresult = CoCreateInstance( clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, ( void ** ) &objRetVal )
	)
		)
	{
		// Comオブジェクトの取得に失敗
		eErrorState = OleWrapInstanceError;
	}
	OleWrap::m_eErrorState = eErrorState;
	OleWrap::m_nResult = nHresult;
	return objRetVal;
}

void OleWrap::ReleaseObject( IDispatch * p_objObj )
{
	if( p_objObj != NULL )
	{
		p_objObj->Release();
	}
}

void SafeArrayCtrl::Construct( UINT p_nRow, UINT p_nColum )
{
	if( p_nRow > 0UL )
	{
		p_nColum = p_nColum == 0UL ? 1UL : p_nColum;
		this->m_stRowBound.lLbound = 1;	// エクセルのシートは1始まり
		this->m_stRowBound.cElements = p_nRow;
		this->m_stColBound.lLbound = 1;	// エクセルのシートは1始まり
		this->m_stColBound.cElements = p_nColum;
		SAFEARRAYBOUND aryBound[ 2 ] = { this->m_stRowBound, this->m_stColBound };
		this->m_pArray = SafeArrayCreate( VT_VARIANT, 2, aryBound);
	}

}

SafeArrayCtrl::SafeArrayCtrl()
{
	this->Construct( 0, 0 );
}

SafeArrayCtrl::SafeArrayCtrl( VARIANT *p_pVar )
{

	this->m_pArray = p_pVar->parray;
	SafeArrayGetLBound(this->m_pArray, 1, &this->m_stRowBound.lLbound);
	SafeArrayGetUBound( this->m_pArray, 1, (LONG*)(&this->m_stRowBound.cElements));
	SafeArrayGetLBound( this->m_pArray, 2, & (this->m_stColBound.lLbound) );
	SafeArrayGetUBound( this->m_pArray, 2, (LONG*)(&this->m_stColBound.cElements));

	

}

SafeArrayCtrl::SafeArrayCtrl( UINT p_nRow, UINT p_nColum )
{
	this->Construct( p_nRow, p_nColum );

}


VARIANT SafeArrayCtrl::get( UINT p_nRow, UINT p_nColumn )
{
	LONG    indices[ 2 ];
	VARIANT vRetVal;
	indices[ 0 ] = p_nRow;
	indices[ 1 ] = p_nColumn;
	SafeArrayGetElement( this->m_pArray, indices, &vRetVal );

	return vRetVal;
}

void SafeArrayCtrl::set( UINT p_nRow,UINT p_nColumn, VARIANT *p_vVar )
{
	LONG indices[] = { (LONG) p_nRow, (LONG) p_nColumn };
	VARIANT vCopy;
	VariantInit( &vCopy );
	VariantCopy( &vCopy, p_vVar );
	SafeArrayPutElement( this->m_pArray, indices, (void *) &vCopy);

}

SafeArrayCtrl::~SafeArrayCtrl()
{
	VariantClear( &(this->toVariant()) );

	//delete  this->m_pArray ;
}

VARIANT SafeArrayCtrl::toVariant()
{
	VARIANT vRetVal;
	vRetVal.vt = VT_ARRAY | VT_VARIANT;
	vRetVal.parray = this->m_pArray;
	return vRetVal;
}

VARIANT VariantCtrl::fromInteger( INT p_nVal )
{
	VARIANT vRetVal;
	VariantInit( &vRetVal);
	vRetVal.vt = VT_I4;
	vRetVal.lVal = p_nVal;

	return vRetVal;
}

VARIANT VariantCtrl::fromString( const OLECHAR *p_strVal )
{
	VARIANT vRetVal;
	VariantInit( &vRetVal );
	vRetVal.vt = VT_BSTR;
	vRetVal.bstrVal = ::SysAllocString( p_strVal);
	return vRetVal;
}

VARIANT VariantCtrl::fromDouble( DOUBLE p_dVal )
{
	VARIANT vRetVal;
	VariantInit( &vRetVal );
	vRetVal.vt = VT_R8;
	vRetVal.dblVal = p_dVal;
	return vRetVal;
}

int VariantCtrl::toInteger( VARIANT p_vVal )
{
	int nRetVal;
	VARIANT varNew;
	VariantInit( &varNew );
	VariantChangeType( &varNew, &p_vVal, 0, VT_I4 );
	nRetVal = varNew.lVal;
	VariantClear( &varNew );

	return nRetVal;
}

std::wstring VariantCtrl::toString( VARIANT p_vVal)
{
	VARIANT varNew;
	std::wstring strRetVal;
	VariantInit( &varNew );
	if( p_vVal.vt != VT_BSTR )
	{

		VariantChangeType( &varNew, &p_vVal, 0, VT_BSTR );
		strRetVal = std::wstring( (LPWSTR) varNew.bstrVal );
		VariantClear( &varNew );

	}
	else
	{
		strRetVal = std::wstring( (LPWSTR) p_vVal.bstrVal );

	}

	return strRetVal;
}

double VariantCtrl::toDouble( VARIANT p_vVal)
{
	VARIANT varNew;
	double dblRetVal;
	VariantInit( &varNew );
	if( p_vVal.vt != VT_R8 )
	{
		VariantChangeType( &varNew, &p_vVal, 0, VT_R8 );
		dblRetVal = varNew.dblVal;
		VariantClear( &varNew );
	}
	else
	{
		dblRetVal = p_vVal.dblVal;
	}
	return dblRetVal;
}
