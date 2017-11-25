#pragma once
#include <list>
#include "OleWrp.h"

#define ALPHABET_COUNT 26


// エクセル処理クラス
//　※Excelオブジェクトに関しては、Excel VisualBasic for Applications のオブジェクトブラウザを参照
class ExcelProc
{
	private:
	INT m_nOpenSheet;//最初に開いたシート
	void SelectSheet( VARIANT );
	public:

	typedef enum EXCEL_PROC_ERROR
	{
		ExcelProcNoError = 0,
		ExcelProcClsIdUndefined,
		ExcelProcExcelAppUndefined,
		ExcelProcAppError,
		ExcelProcBooksError,
		ExcelProcFileNotFound,

	} tag_EXCEL_PROC_ERROR;

	HRESULT m_nError;
	EXCEL_PROC_ERROR m_eErrorStep;

	IDispatch *m_pXlApp;// ExcelApplicationオブジェクト
	IDispatch *m_pXlBooks;//WorkBookコレクション
	IDispatch *m_pXlBook;// アクティブなWorkBook
	IDispatch *m_pXlSheet;// アクティブなWorkSheet
	// コンストラクタ：
	//   1. Com	オジェクト：ExcelApplication を呼び出し
	//   2. エクセルアプリケーションオブジェクトを生成する。
	//   3. Application.Visible = 1
	//   4. mpXlBooks ← Application.Workbooks
	ExcelProc(); 

	// デストラクタ
	// ブックが開かれている場合、開いたブックを強制的に閉じる
	// Excelアプリケーションを終了する
	~ExcelProc();
	
	//m_pXlBook← Application.Workbooks.Open(fileName)
	EXCEL_PROC_ERROR Open( const OLECHAR* p_strFileName);

	// シートを指定する 0:(開いた時点のアクティブシート)
	void SelectSheet( UINT p_nSheetId);

	void SelectSheet( const OLECHAR* p_strSheetName);

	// エクセル列コード(A,B,C...Z,AA,AB～ZZZ=))を返す(1からカウント)
	static std::wstring ColumnChar( UINT p_ColumnNo);

	// エクセルのレンジ指定文字列を返す 例(1行, 2列 ～ 3行 4列) → "B1:D3"
	static std::wstring RangeCode( UINT p_RowStart, UINT p_ColumnStart, UINT p_RowEnd, UINT p_ColumnEnd );

	// レンジ範囲の値を取得する
	SafeArrayCtrl *getRange( std::wstring p_strRangeCode );

	// レンジ範囲に値を設定する
	void setRange( std::wstring  p_strRangeCode, SafeArrayCtrl* p_pArrayData);

	// 新しいブックを開く
	void NewBook();

	// 新しいシートを現在のシートの左側に追加する
	void AddSheet();

	// ブックを閉じる
	void Close();

	// 保存する
	void SaveAs( const OLECHAR* p_strFileName );
	

};

