#include "ExcelProc.h"

#define FILE_TEST L"test.xlsx"
#define MSG_TITLE L"エクセルテスト"
#define MAX_PATH 256

int main( int argc, char *argv[] )
{
	ExcelProc *proc;
	SafeArrayCtrl *Array;
	WCHAR wstrExePath[ MAX_PATH ];
	std::wstring strFilePath;

	setlocale( LC_ALL, ""  );// 日本語を使うための設定

	// 実行ファイルのディレクトリを取得する
	GetCurrentDirectoryW( MAX_PATH, wstrExePath );

	strFilePath = std::wstring( wstrExePath ) + std::wstring( L"\\" ) + std::wstring( FILE_TEST );

	proc = new ExcelProc();
	if( proc->m_eErrorStep != ExcelProc::ExcelProcNoError )
	{
		MessageBoxExW( NULL, L"Microsoft Excelをインストールしてください。。。", MSG_TITLE, 0, 0 );

	}
	else
	{
		// 新しいBookを開く
		proc->NewBook();

		// Variant配列を生成し、値を設定する。
		Array = new SafeArrayCtrl( 3, 1 );
		Array->set( 1, 1, &VariantCtrl::fromDouble( 1.2345 ) );
		Array->set( 2, 1, &VariantCtrl::fromInteger( 524569 ) );
		Array->set( 3, 1, &VariantCtrl::fromString( L"日本語" ) );
		// シートに書き出すMSG_TITLE
		proc->setRange( ExcelProc::RangeCode( 1, 1, 3, 1 ), Array );
		delete Array;

		// セーブして閉じる
		proc->SaveAs( strFilePath.c_str() );

		MessageBoxExW( NULL, L"ファイル生成", MSG_TITLE, 0, 0 );

		proc->Close();

		//	ファイルを開く
		if( proc->Open( strFilePath.c_str() ) != ExcelProc::ExcelProcNoError )
		{
			// ファイルが無い
			MessageBoxExW( NULL, L"ファイルがないよ", MSG_TITLE, 0, 0 );

		}
		else
		{
			proc->SelectSheet( L"Sheet1" );
			Array = proc->getRange( ExcelProc::RangeCode( 1, 1, 5, 1 ) );
			// 表示
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
			// ファイルが無い
			MessageBoxExW( NULL, L"ファイル取得", MSG_TITLE, 0, 0 );

			proc->Close();

		}
		// 使い終わったら破棄すること
		delete proc;

		DeleteFile( strFilePath.c_str() );

	}


	return 0; 
}

