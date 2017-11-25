#pragma once
#include <ole2.h> // OLE2 Definitions
#include <Windows.h>
#include <stdio.h>
#include <string>

#define OLE_NAME_SIZE 200

// OLE����p���b�p
// �Q�l�Fhttps://support.microsoft.com/ja-jp/help/216686/how-to-automate-excel-from-c-without-using-mfc-or-import
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
	// IDispatch�擾���b�p�[
	// �������̃p�����[�^�͋t���ɓn�����ƁB
	// p_nAutoType :: DISPATCH_METHOD / DISPATCH_PROPERTYGET / DISPATCH_PROPERTYPUT / DISPATCH_PROPERTYPUTREF
	// p_pVResult :: ���ʂ̒l
	// p_pDisp :: �擾���̃I�u�W�F�N�g
	// p_ptName :: �I�u�W�F�N�g�̃����o�i�v���p�e�B�⃁�\�b�h)��
	// p_nCArgs :: �p�����[�^�����̐�
	// ... :: �p�����[�^����(�ϐ�)
//	static HRESULT AutoWrap( int p_nAutoType, VARIANT *p_pVResult, IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );
	
	// �I�u�W�F�N�g���擾����
	// p_pDisp:: �Ăяo�����I�u�W�F�N�g
	// p_ptName:: �v���p�e�B�� or �֐���
	// p_nCHargs:: �p�����[�^�����̐�
	// ... :: �p�����[�^����(�ϐ�)
	static IDispatch *getObject( IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );

	// �v���p�e�B���擾����
	// p_pDisp:: �Ăяo�����I�u�W�F�N�g
	// p_ptName:: �v���p�e�B�� or �֐���
	static VARIANT getValue( IDispatch *p_pDisp, LPOLESTR p_ptName );

	// �v���p�e�B��ݒ肷��
	// p_pDisp:: �Ăяo�����I�u�W�F�N�g
	// p_ptName:: �v���p�e�B�� or �֐���
	// p_nCHargs:: �p�����[�^�����̐�
	// ... :: �p�����[�^����(�ϐ�)
	static void setValue( IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );

	// �֐����Ăяo��
	// p_pDisp:: �Ăяo�����I�u�W�F�N�g
	// p_ptName:: �v���p�e�B�� or �֐���
	// p_nCHargs:: �p�����[�^�����̐�
	// ... :: �p�����[�^����(�ϐ�)
	static VARIANT execMethod( IDispatch *p_pDisp, LPOLESTR p_ptName, int p_nCArgs, ... );

	// COM�I�u�W�F�N�g���擾����
	// p_ptName :: �I�u�W�F�N�g��
	static IDispatch *getInstance( LPOLESTR p_ptName );

	// �I�u�W�F�N�g���������
	// p_objObj :: �I�u�W�F�N�g
	static void ReleaseObject( IDispatch *p_objObj );
};

// Variant�^�@2�����z��𐶐�
// �Q�ƁFhttp://officetanaka.net/excel/vba/speed/s11.htm
//       http://eternalwindows.jp/com/auto/auto04.html
class SafeArrayCtrl
{
	private :
	SAFEARRAY *m_pArray;
	void Construct( UINT, UINT );
	public:
	// �s�����̊J�n�ʒu�ƌ�
	SAFEARRAYBOUND m_stRowBound;
	// ������̊J�n�ʒu�ƌ�
	SAFEARRAYBOUND m_stColBound;
	// 1x1��Variant�z��𐶐�
	SafeArrayCtrl();
	// Variant����擾����
	SafeArrayCtrl( VARIANT* p_vVal);
	// 2������Variant�z��𐶐�
	SafeArrayCtrl( UINT p_nRow, UINT p_nColumn);
	// Variant��Ԃ�
	VARIANT toVariant();
	// �Q�b�^�[�ƃZ�b�^�[(Excel��1�n�܂�ł��邱�Ƃɒ��ӁI�I�j
	VARIANT get( UINT p_nRow, UINT p_nColumn);
	void set( UINT p_nRow, UINT p_nColumn, VARIANT* p_vVal );
	// �f�X�g���N�^
	~SafeArrayCtrl();

};

class VariantCtrl
{
	public:
	// �����l��������Variant�^�𐶐����܂��B
	static VARIANT fromInteger( INT  p_nVal);
	// ������^��������Variant�^�𐶐����܂��B
	static VARIANT fromString( const OLECHAR* p_strVal );
	// double�l��������Variant�^�𐶐����܂��B
	static VARIANT fromDouble( DOUBLE p_dblVal );
	// ���̌^�͕ʓr��`����낵

	// ���l��Ԃ��܂��i���l�ł͂Ȃ��ꍇ�F0)
	static int toInteger( VARIANT  p_vVal);
	// �������Ԃ��܂�(������o�Ȃ��ꍇ�󕶎���)
	static std::wstring toString( VARIANT p_vVal );
	// double�^��Ԃ��܂��B�i���l�ł͂Ȃ��ꍇ0)
	static double toDouble( VARIANT p_vVal);

	// ���t�^�͖�����

};
