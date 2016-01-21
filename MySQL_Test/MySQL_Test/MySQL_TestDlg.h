
// MySQL_TestDlg.h : ͷ�ļ�
//

#pragma once

#include <winsock2.h>
//����socket
#include "mysql.h"
#include "afxcmn.h"

// CMySQL_TestDlg �Ի���
class CMySQL_TestDlg : public CDialog
{
// ����
public:
	CMySQL_TestDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_MYSQL_TEST_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
public:
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnBnClickedBtnSqlOpen();
	afx_msg void OnDestroy();
	afx_msg void OnBnClickedExceltomysql();
	DECLARE_MESSAGE_MAP()	
public:
	BOOL DBIsExist( MYSQL mysql,CString strDB );
	BOOL TableIsExist( MYSQL mysql,CString strDB,CString strTable );
	void WriteInToMySQLDB( int nCount,CString* pStr,int nArrLen );
public:
	MYSQL m_mysql;
	CListCtrl m_ListPeopleInfo;
	CString m_strStatusTips;
	CString m_strExcelFilePath;
};
