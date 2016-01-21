
// MySQL_TestDlg.h : 头文件
//

#pragma once

#include <winsock2.h>
//定义socket
#include "mysql.h"
#include "afxcmn.h"

// CMySQL_TestDlg 对话框
class CMySQL_TestDlg : public CDialog
{
// 构造
public:
	CMySQL_TestDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_MYSQL_TEST_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
