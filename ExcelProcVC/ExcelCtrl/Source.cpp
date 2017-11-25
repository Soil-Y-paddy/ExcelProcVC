#include "ExcelProc.h"

#define FILE_TEST L"test.xlsx"
#define MSG_TITLE L"�G�N�Z���e�X�g"
#define MAX_PATH 256

int main( int argc, char *argv[] )
{
	ExcelProc *proc;
	SafeArrayCtrl *Array;
	WCHAR wstrExePath[ MAX_PATH ];
	std::wstring strFilePath;

	setlocale( LC_ALL, ""  );// ���{����g�����߂̐ݒ�

	// ���s�t�@�C���̃f�B���N�g�����擾����
	GetCurrentDirectoryW( MAX_PATH, wstrExePath );

	strFilePath = std::wstring( wstrExePath ) + std::wstring( L"\\" ) + std::wstring( FILE_TEST );

	proc = new ExcelProc();
	if( proc->m_eErrorStep != ExcelProc::ExcelProcNoError )
	{
		MessageBoxExW( NULL, L"Microsoft Excel���C���X�g�[�����Ă��������B�B�B", MSG_TITLE, 0, 0 );

	}
	else
	{
		// �V����Book���J��
		proc->NewBook();

		// Variant�z��𐶐����A�l��ݒ肷��B
		Array = new SafeArrayCtrl( 3, 1 );
		Array->set( 1, 1, &VariantCtrl::fromDouble( 1.2345 ) );
		Array->set( 2, 1, &VariantCtrl::fromInteger( 524569 ) );
		Array->set( 3, 1, &VariantCtrl::fromString( L"���{��" ) );
		// �V�[�g�ɏ����o��MSG_TITLE
		proc->setRange( ExcelProc::RangeCode( 1, 1, 3, 1 ), Array );
		delete Array;

		// �Z�[�u���ĕ���
		proc->SaveAs( strFilePath.c_str() );

		MessageBoxExW( NULL, L"�t�@�C������", MSG_TITLE, 0, 0 );

		proc->Close();

		//	�t�@�C�����J��
		if( proc->Open( strFilePath.c_str() ) != ExcelProc::ExcelProcNoError )
		{
			// �t�@�C��������
			MessageBoxExW( NULL, L"�t�@�C�����Ȃ���", MSG_TITLE, 0, 0 );

		}
		else
		{
			proc->SelectSheet( L"Sheet1" );
			Array = proc->getRange( ExcelProc::RangeCode( 1, 1, 5, 1 ) );
			// �\��
			for( ULONG nRow = (ULONG) Array->m_stRowBound.lLbound; nRow <= Array->m_stRowBound.cElements; nRow++ )
			{
				for( ULONG nColumn = (ULONG) Array->m_stColBound.lLbound; nColumn <= Array->m_stColBound.cElements; nColumn++ )
				{
					VARIANT vCell = Array->get( nRow, nColumn );
					wprintf( L"%d, %s, %f \n",
							 VariantCtrl::toInteger( vCell ),
							 (LPWSTR) VariantCtrl::toString( vCell ).c_str(),
							 VariantCtrl::toDouble( vCell )
					);


				}

			}
			delete Array;
			// �t�@�C��������
			MessageBoxExW( NULL, L"�t�@�C���擾", MSG_TITLE, 0, 0 );

			proc->Close();

		}
		// �g���I�������j�����邱��
		delete proc;

		DeleteFile( strFilePath.c_str() );

	}


	return 0; 
}

