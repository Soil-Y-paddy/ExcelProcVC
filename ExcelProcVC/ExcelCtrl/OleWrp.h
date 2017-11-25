#pragma once
#include <ole2.h> // OLE2 Definitions
#include <Windows.h>
#include <stdio.h>
#include <string>

#define OLE_NAME_SIZE 200

// OLE操作用ラッパ
// 参考：https://support.microsoft.com/ja-jp/help/216686/how-to-automate-excel-from-c-without-using-mfc-or-import
class OleWrap
{
	private:
	static HRESULT Invoker( int p_nAutoType, VARIANT *p_pVResult, IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, VARIANT *p_aryArgs );
	public:
	typedef enum
	{
		OleWrapNoError = 0,
		OleWrapClsIdNotExist,
		OleWrapInstanceError,
		OleWrapNullPointer,
		OleWrapGetIdOfName,
		OleWrapInvoke,
	} OleWrapError;
	static OleWrapError m_eErrorState;
	static HRESULT m_nResult;
	// IDispatch取得ラッパー
	// ※複数のパラメータは逆順に渡すこと。
	// p_nAutoType :: DISPATCH_METHOD / DISPATCH_PROPERTYGET / DISPATCH_PROPERTYPUT / DISPATCH_PROPERTYPUTREF
	// p_pVResult :: 結果の値
	// p_pDisp :: 取得元のオブジェクト
	// p_ptName :: オブジェクトのメンバ（プロパティやメソッド)名
	// p_nCArgs :: パラメータ引数の数
	// ... :: パラメータ引数(可変数)
//	static HRESULT AutoWrap( int p_nAutoType, VARIANT *p_pVResult, IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );
	
	// オブジェクトを取得する
	// p_pDisp:: 呼び出し元オブジェクト
	// p_ptName:: プロパティ名 or 関数名
	// p_nCHargs:: パラメータ引数の数
	// ... :: パラメータ引数(可変数)
	static IDispatch *getObject( IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );

	// プロパティを取得する
	// p_pDisp:: 呼び出し元オブジェクト
	// p_ptName:: プロパティ名 or 関数名
	static VARIANT getValue( IDispatch *p_pDisp, LPOLESTR p_ptName );

	// プロパティを設定する
	// p_pDisp:: 呼び出し元オブジェクト
	// p_ptName:: プロパティ名 or 関数名
	// p_nCHargs:: パラメータ引数の数
	// ... :: パラメータ引数(可変数)
	static void setValue( IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );

	// 関数を呼び出す
	// p_pDisp:: 呼び出し元オブジェクト
	// p_ptName:: プロパティ名 or 関数名
	// p_nCHargs:: パラメータ引数の数
	// ... :: パラメータ引数(可変数)
	static VARIANT execMethod( IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );

	// COMオブジェクトを取得する
	// p_ptName :: オブジェクト名
	static IDispatch *getInstance( LPOLESTR p_ptName );

	// オブジェクトを解放する
	// p_objObj :: オブジェクト
	static void ReleaseObject( IDispatch *p_objObj );
};

// Variant型　2次元配列を生成
// 参照：http://officetanaka.net/excel/vba/speed/s11.htm
//       http://eternalwindows.jp/com/auto/auto04.html
class SafeArrayCtrl
{
	private :
	SAFEARRAY *m_pArray;
	void Construct( UINT, UINT );
	public:
	// 行方向の開始位置と個数
	SAFEARRAYBOUND m_stRowBound;
	// 列方向の開始位置と個数
	SAFEARRAYBOUND m_stColBound;
	// 1x1のVariant配列を生成
	SafeArrayCtrl();
	// Variantから取得する
	SafeArrayCtrl( VARIANT* p_vVal);
	// 2次元のVariant配列を生成
	SafeArrayCtrl( UINT p_nRow, UINT p_nColumn);
	// Variantを返す
	VARIANT toVariant();
	// ゲッターとセッター(Excelは1始まりであることに注意！！）
	VARIANT get( UINT p_nRow, UINT p_nColumn);
	void set( UINT p_nRow, UINT p_nColumn, VARIANT* p_vVal );
	// デストラクタ
	~SafeArrayCtrl();

};

class VariantCtrl
{
	public:
	// 整数値が入ったVariant型を生成します。
	static VARIANT fromInteger( INT  p_nVal);
	// 文字列型が入ったVariant型を生成します。
	static VARIANT fromString( const OLECHAR* p_strVal );
	// double値が入ったVariant型を生成します。
	static VARIANT fromDouble( DOUBLE p_dblVal );
	// 他の型は別途定義すよろし

	// 数値を返します（数値ではない場合：0)
	static int toInteger( VARIANT  p_vVal);
	// 文字列を返します(文字列出ない場合空文字列)
	static std::wstring toString( VARIANT p_vVal );
	// double型を返します。（数値ではない場合0)
	static double toDouble( VARIANT p_vVal);

	// 日付型は未実装

};
