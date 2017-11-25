

HRESULT GetRangeValue( IDispatch *pRange )
{
	LONG      i, j;
	LONG      lIndexMin1, lIndexMax1;
	LONG      lIndexMin2, lIndexMax2;
	LONG      indices[ 2 ];
	SAFEARRAY *pArray;
	VARIANT   var;
	VARIANT   varResult;
	HRESULT   hr;
	TCHAR     szBuf[ 256 ];

	VariantInit( &varResult );
	hr = OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &varResult, pRange, L"Value", 0 );

	//hr = Invoke( pRange, L"Value", DISPATCH_PROPERTYGET, NULL, 0, &varResult );
	if( FAILED( hr ) )
		return hr;
	pArray = varResult.parray;

	SafeArrayGetLBound( pArray, 1, &lIndexMin1 );
	SafeArrayGetUBound( pArray, 1, &lIndexMax1 );
	SafeArrayGetLBound( pArray, 2, &lIndexMin2 );
	SafeArrayGetUBound( pArray, 2, &lIndexMax2 );

	for( i = lIndexMin1; i <= lIndexMax1; i++ )
	{
		for( j = lIndexMin2; j <= lIndexMax2; j++ )
		{
			indices[ 0 ] = i;
			indices[ 1 ] = j;
			VariantInit( &var );
			SafeArrayGetElement( pArray, indices, &var );
			if( var.vt == VT_BSTR )
			{
				MessageBoxW( NULL, (LPWSTR) var.bstrVal, L"OK", MB_OK );
				VariantClear( &var );
			}
			else if( var.vt == VT_R8 )
			{
				VARIANT varNew;
				VariantInit( &varNew );
				VariantChangeType( &varNew, &var, 0, VT_BSTR );

				MessageBoxW( NULL, (LPWSTR) varNew.bstrVal, L"OK", MB_OK );
				VariantClear( &varNew );
				double x = var.dblVal;

			}
			else if( var.vt == VT_DATE )
			{
				SYSTEMTIME systemTime;
				VariantTimeToSystemTime( var.date, &systemTime );
				wsprintf( szBuf, TEXT( "%d/%d/%d" ), systemTime.wYear, systemTime.wMonth, systemTime.wDay );
				MessageBoxW( NULL, szBuf, TEXT( "OK" ), MB_OK );
			}
			else if( var.vt == VT_EMPTY )
				MessageBoxW( NULL, TEXT( "" ), TEXT( "OK" ), MB_OK );
			else
			{
				wsprintf( szBuf, TEXT( "予期しないデータ型 %d" ), var.vt );
				MessageBox( NULL, szBuf, TEXT( "OK" ), MB_OK );
			}
		}
	}

	VariantClear( &varResult );

	return hr;
}


int main2( void )
{


	// Initialize COM for this thread...
	CoInitialize( NULL );

	// Get CLSID for our server...
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID( L"Excel.Application", &clsid );

	if( FAILED( hr ) )
	{

		::MessageBoxW( NULL, L"CLSIDFromProgID() failed", L"Error", 0x10010 );
		return -1;
	}

	// Start server and get IDispatch...
	IDispatch *pXlApp;
	hr = CoCreateInstance( clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **) &pXlApp );
	if( FAILED( hr ) )
	{
		::MessageBoxW( NULL, L"Excel not registered properly", L"Error", 0x10010 );
		return -2;
	}

	// Make it visible (nRow.e. app.visible = 1)
	{

		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		OleWrap::AutoWrap( DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x );
	}

	// Get Workbooks collection  ＜＜　Aplication.Workbooks
	IDispatch *pXlBooks;
	{
		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0 );
		pXlBooks = result.pdispVal;
	}

	// Open　＜＜　Applicaiton.Workbooks.Open(fileName)
	IDispatch *pXlBook = NULL;
	{
		VARIANT parm;
		VARIANT result;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString( L"D:\\test.xlsx" );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlBooks, L"Open", 1, parm );
		pXlBook = result.pdispVal;
	}
	// sheet
	// Get ActiveSheet object
	IDispatch *pXlSheet;
	{
		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlBook, L"ActiveSheet", 0 );
		pXlSheet = result.pdispVal;
	}

	{
		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlSheet, L"Index", 0 );
		printf( "%d", result.decVal );
	}


	// get Range
	IDispatch *pXlRange;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString( L"A1:A5" );

		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm );
		VariantClear( &parm );

		pXlRange = result.pdispVal;
	}
	GetRangeValue( pXlRange );


#if 0
	/* Create */
	// Call Workbooks.Add() to get a new workbook...
	IDispatch *pXlBook;
	{
		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlBooks, L"Add", 0 );
		pXlBook = result.pdispVal;
	}

	// Create a 15x15 safearray of variants...
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	{
		SAFEARRAYBOUND sab[ 2 ];
		sab[ 0 ].lLbound = 1; sab[ 0 ].cElements = 15;
		sab[ 1 ].lLbound = 1; sab[ 1 ].cElements = 15;
		arr.parray = SafeArrayCreate( VT_VARIANT, 2, sab );
	}

	// Fill safearray with some values...
	for( int i = 1; i <= 15; i++ )
	{
		for( int j = 1; j <= 15; j++ )
		{
			// Create entry value for (nRow,j)
			VARIANT tmp;
			tmp.vt = VT_I4;
			tmp.lVal = i*j;
			// Add to safearray...
			long indices[] = { i,j };
			SafeArrayPutElement( arr.parray, indices, (void *) &tmp );
		}
	}

	// Get ActiveSheet object
	IDispatch *pXlSheet;
	{
		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0 );
		pXlSheet = result.pdispVal;
	}

	// Get Range object for the Range A1:O15...
	IDispatch *pXlRange;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString( L"A1:O15" );

		VARIANT result;
		VariantInit( &result );
		OleWrap::AutoWrap( DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm );
		VariantClear( &parm );

		pXlRange = result.pdispVal;
	}

	// Set range with our safearray...
	OleWrap::AutoWrap( DISPATCH_PROPERTYPUT, NULL, pXlRange, L"Value", 1, arr );
#endif
	// Wait for user...
	::MessageBoxW( NULL, L"All done.", L"Notice", 0x10000 );

	// Set .Saved property of workbook to TRUE so we aren't prompted
	// to save when we tell Excel to quit...
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		OleWrap::AutoWrap( DISPATCH_PROPERTYPUT, NULL, pXlBook, L"Saved", 1, x );
	}

	// Tell Excel to quit (nRow.e. App.Quit)
	OleWrap::AutoWrap( DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0 );

	// Release references...
	pXlRange->Release();
	pXlSheet->Release();
	pXlBook->Release();
	pXlBooks->Release();
	pXlApp->Release();
	//	VariantClear( &arr );

	// Uninitialize COM for this thread...
	CoUninitialize();
}