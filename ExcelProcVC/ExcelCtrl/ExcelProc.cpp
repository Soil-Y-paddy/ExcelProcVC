#include "ExcelProc.h"

// �R���X�g���N�^
ExcelProc::ExcelProc()
{

	this->m_nError = S_OK;
	this->m_eErrorStep = ExcelProcNoError;
	// Initialize COM for this thread...
	CoInitialize( NULL );

	this->m_pXlBooks = NULL;
	this->m_pXlBook = NULL;
	this->m_pXlSheet = NULL;

	// Com�I�u�W�F�N�g���Ăяo��
	// Get CLSID for our server...
	this->m_pXlApp = OleWrap::getInstance( L"Excel.Application" );
	if( this->m_pXlApp == NULL )
	{
		// �G���[�ʒm
		this->m_nError = OleWrap::m_nResult;
	}
	// Application.visible = 1;
	OleWrap::setValue( this->m_pXlApp, L"Visible", 1, VariantCtrl::fromInteger( 1 ) );

	// m_pXlBooks �� Aplication.Workbooks
	this->m_pXlBooks =  OleWrap::getObject( this->m_pXlApp, L"Workbooks", 0 );
}

// �f�X�g���N�^
ExcelProc::~ExcelProc()
{

	// Application.Quit
	OleWrap::execMethod( this->m_pXlApp, L"Quit", 0 );

	OleWrap::ReleaseObject( this->m_pXlSheet );
	OleWrap::ReleaseObject( this->m_pXlBook );
	OleWrap::ReleaseObject( this->m_pXlBooks );
	OleWrap::ReleaseObject( this->m_pXlApp );

	CoUninitialize();
}

// �u�b�N���J���A�A�N�e�B�u�V�[�g����Ɨp�Ƃ��Ċm�ۂ���
// �����łɊJ���Ă���ꍇ�̓���͖���`
ExcelProc::EXCEL_PROC_ERROR ExcelProc::Open( const OLECHAR * p_strFileName )
{
	EXCEL_PROC_ERROR eError = ExcelProcNoError;
	VARIANT vParam = VariantCtrl::fromString( p_strFileName );

	//m_pXlBook �� Workbooks.Open( fileName )
	IDispatch *objTemp = OleWrap::getObject(this->m_pXlBooks,L"Open", 1, vParam );
	if( objTemp == NULL )
	{
		this->m_eErrorStep = ExcelProcFileNotFound;
		eError = ExcelProcFileNotFound;
	}
	else{
		this->m_pXlBook = objTemp;
		// �V�[�g�����݂��Ȃ��u�b�N�͂Ȃ��̂ŁA�G���[�`�F�b�N���Ȃ��B
		this->m_pXlSheet = OleWrap::getObject(this->m_pXlBook, L"ActiveSheet", 0 );
		// �A�N�e�B�u�ȃV�[�g��ID���擾
		this->m_nOpenSheet = VariantCtrl::toInteger( OleWrap::getValue( this->m_pXlSheet, L"Index" ) );
	}

	return eError;

}

void ExcelProc::SelectSheet( VARIANT p_vValue)
{
	if( this->m_pXlBook == NULL )
	{
		this->m_eErrorStep = ExcelProcFileNotFound;
	}
	else
	{
		IDispatch *objTemp = OleWrap::getObject( this->m_pXlBook, L"Sheets", 1, p_vValue );
		// �V�[�g�����݂���Ƃ������ʂ�
		if(objTemp != NULL )
		{
			this->m_pXlSheet = objTemp;
		}
	}

}




void ExcelProc::SelectSheet( UINT p_nSheetNo )
{
	// �J�������_�̃A�N�e�B�u�V�[�g���Z�b�g����B
	if( p_nSheetNo == 0UL )
	{
		SelectSheet( VariantCtrl::fromInteger( this->m_nOpenSheet ) );
	}
	else
	{
		SelectSheet( VariantCtrl::fromInteger( p_nSheetNo ) );
	}
}

void ExcelProc::SelectSheet( const OLECHAR *p_SheetName )
{
	SelectSheet( VariantCtrl::fromString( p_SheetName ) );
}

std::wstring ExcelProc::ColumnChar( UINT p_nColumnId )
{
	std::wstring strRetVal = std::wstring( L"" );
	if( p_nColumnId > 0 )
	{
		int nMod = p_nColumnId % ALPHABET_COUNT;
		nMod = ( nMod == 0 ) ? ALPHABET_COUNT : nMod;
		std::wstring sTemp = std::wstring( 1,  (char) ( nMod + 'A' - 1 ) );//�@'A'�`'Z'
		if( p_nColumnId == nMod )
		{
			strRetVal = sTemp;
		}else {
			strRetVal = ColumnChar( ( p_nColumnId - nMod ) / ALPHABET_COUNT ) + sTemp; // �ċA�ďo��
		}
	}

	return strRetVal;
}

std::wstring ExcelProc::RangeCode( UINT p_RowStart, UINT p_ColumnStart, UINT p_RowEnd, UINT p_ColumnEnd )
{
	return ColumnChar(p_ColumnStart)+std::to_wstring(p_RowStart)
		+ std::wstring(L":" ) + ColumnChar( p_ColumnEnd ) + std::to_wstring(p_RowEnd);
}

SafeArrayCtrl* ExcelProc::getRange( std::wstring  p_strRangeCode)
{
	// �����W������̃p�����[�^�𐶐� 
	VARIANT vParam = VariantCtrl::fromString(  (p_strRangeCode.c_str()) );
	IDispatch *objRange;
	VARIANT vValue;

	// Sheet.Range(p_strRangeCode).Value
	if( this->m_pXlSheet != NULL )
	{
		objRange = OleWrap::getObject( this->m_pXlSheet, L"Range", 1, vParam );
		vValue = OleWrap::getValue( objRange, L"Value" );
		OleWrap::ReleaseObject( objRange );
	}

	return  new SafeArrayCtrl(&vValue );
}

void ExcelProc::setRange( std::wstring  p_strRangeCode,  SafeArrayCtrl* p_pArrayData ) 
{
	// �����W������̃p�����[�^�𐶐� 
	VARIANT vParam = VariantCtrl::fromString( ( p_strRangeCode.c_str() ) );
	IDispatch *objRange;
	// Sheet.Range(p_strRangeCode).Value
	if( this->m_pXlSheet != NULL )
	{
		objRange = OleWrap::getObject( this->m_pXlSheet, L"Range", 1, vParam );
		OleWrap::setValue( objRange, L"Value" ,1, ( p_pArrayData->toVariant() ) );
		OleWrap::ReleaseObject( objRange );
	}
}

// �V�����u�b�N���J���@(���̑O�ɊJ���Ă����u�b�N�͂��̂܂ܕ��u)
// Excel�I�u�W�F�N�g�{�̂ɂ͎c�����܂܁B
void ExcelProc::NewBook()
{

	// WorkBooks.Add
	this->m_pXlBook = OleWrap::getObject( this->m_pXlBooks, L"Add", 0 );

	// �V�[�g�����݂��Ȃ��u�b�N�͂Ȃ��̂ŁA�G���[�`�F�b�N���Ȃ��B
	this->m_pXlSheet = OleWrap::getObject( this->m_pXlBook, L"ActiveSheet", 0 );
	this->m_nOpenSheet = VariantCtrl::toInteger( OleWrap::getValue( this->m_pXlSheet, L"Index" ) );

}

void ExcelProc::AddSheet()
{
	if( this->m_pXlBook == NULL )
	{
		this->m_eErrorStep = ExcelProcFileNotFound;
	}
	else
	{
		// Workbook.Sheets.Add
		IDispatch *objSheets = OleWrap::getObject( this->m_pXlBook, L"Sheets", 0 );
		OleWrap::execMethod( objSheets, L"Add", 0 );
		// �A�N�e�B�u�ȃV�[�g(=�V�����V�[�g)�����݂̃V�[�g�ɐݒ�
		this->m_pXlSheet = OleWrap::getObject( this->m_pXlBook, L"ActiveSheet", 0 );

	}


}

// �ۑ������ɏI��
void ExcelProc::Close()
{
	OleWrap::execMethod(this->m_pXlBook,L"Saved", 1, VariantCtrl::fromInteger( 1 ) );
	OleWrap::execMethod( this->m_pXlBook, L"Close", 0);
}

void ExcelProc::SaveAs( const OLECHAR * p_strFileName )
{
	OleWrap::execMethod( this->m_pXlBook, L"SaveAs", 1, VariantCtrl::fromString( p_strFileName ) );
}
