#pragma once
#include <list>
#include "OleWrp.h"

#define ALPHABET_COUNT 26


// �G�N�Z�������N���X
//�@��Excel�I�u�W�F�N�g�Ɋւ��ẮAExcel VisualBasic for Applications �̃I�u�W�F�N�g�u���E�U���Q��
class ExcelProc
{
	private:
	INT m_nOpenSheet;//�ŏ��ɊJ�����V�[�g
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

	IDispatch *m_pXlApp;// ExcelApplication�I�u�W�F�N�g
	IDispatch *m_pXlBooks;//WorkBook�R���N�V����
	IDispatch *m_pXlBook;// �A�N�e�B�u��WorkBook
	IDispatch *m_pXlSheet;// �A�N�e�B�u��WorkSheet
	// �R���X�g���N�^�F
	//   1. Com	�I�W�F�N�g�FExcelApplication ���Ăяo��
	//   2. �G�N�Z���A�v���P�[�V�����I�u�W�F�N�g�𐶐�����B
	//   3. Application.Visible = 1
	//   4. mpXlBooks �� Application.Workbooks
	ExcelProc(); 

	// �f�X�g���N�^
	// �u�b�N���J����Ă���ꍇ�A�J�����u�b�N�������I�ɕ���
	// Excel�A�v���P�[�V�������I������
	~ExcelProc();
	
	//m_pXlBook�� Application.Workbooks.Open(fileName)
	EXCEL_PROC_ERROR Open( const OLECHAR* p_strFileName);

	// �V�[�g���w�肷�� 0:(�J�������_�̃A�N�e�B�u�V�[�g)
	void SelectSheet( UINT p_nSheetId);

	void SelectSheet( const OLECHAR* p_strSheetName);

	// �G�N�Z����R�[�h(A,B,C...Z,AA,AB�`ZZZ=))��Ԃ�(1����J�E���g)
	static std::wstring ColumnChar( UINT p_ColumnNo);

	// �G�N�Z���̃����W�w�蕶�����Ԃ� ��(1�s, 2�� �` 3�s 4��) �� "B1:D3"
	static std::wstring RangeCode( UINT p_RowStart, UINT p_ColumnStart, UINT p_RowEnd, UINT p_ColumnEnd );

	// �����W�͈͂̒l���擾����
	SafeArrayCtrl *getRange( std::wstring p_strRangeCode );

	// �����W�͈͂ɒl��ݒ肷��
	void setRange( std::wstring  p_strRangeCode, SafeArrayCtrl* p_pArrayData);

	// �V�����u�b�N���J��
	void NewBook();

	// �V�����V�[�g�����݂̃V�[�g�̉E���ɒǉ�����
	void AddSheet();

	// �u�b�N�����
	void Close();

	// �ۑ�����
	void SaveAs( const OLECHAR* p_strFileName );
	

};

