
// MySQL_TestDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "MySQL_Test.h"
#include "MySQL_TestDlg.h"

#include <afxdb.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CMySQL_TestDlg �Ի���




CMySQL_TestDlg::CMySQL_TestDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMySQL_TestDlg::IDD, pParent)
	, m_strStatusTips(_T(""))
	, m_strExcelFilePath( "" )
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMySQL_TestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_DISPLAY, m_ListPeopleInfo);
	DDX_Text(pDX, IDC_EDIT_STATUS_TIPS, m_strStatusTips);
}

BEGIN_MESSAGE_MAP(CMySQL_TestDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BTN_SQL_OPEN, &CMySQL_TestDlg::OnBnClickedBtnSqlOpen)
	ON_WM_DESTROY()
	ON_BN_CLICKED(IDC_EXCELTOMYSQL, &CMySQL_TestDlg::OnBnClickedExceltomysql)
END_MESSAGE_MAP()


// CMySQL_TestDlg ��Ϣ�������

BOOL CMySQL_TestDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	CRect rect;

	//��ʼ������齱��Ա��Ϣ�б�ؼ�
	m_ListPeopleInfo.GetClientRect( rect );
	m_ListPeopleInfo.SetExtendedStyle( m_ListPeopleInfo.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES );

	m_ListPeopleInfo.InsertColumn( 0,_T("���"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 1,_T("����"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 2,_T("���֤��"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 3,_T("�÷�"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 4,_T("ͼƬ·��"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 5,_T("ʡ��"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 6,_T("����"),LVCFMT_CENTER,rect.Width()/7,0 );


	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CMySQL_TestDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CMySQL_TestDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CMySQL_TestDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CMySQL_TestDlg::OnBnClickedBtnSqlOpen()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	//��ʼ��
	mysql_init( &m_mysql );

	//�������ݿ�
	bool isConnected = mysql_real_connect( &m_mysql,"localhost","root","bjgas","",3306,NULL,0 );
	if ( isConnected )
	{
		MessageBox(_T("MySQL Success Connected!"));
		//return;
	}
	else
	{
		int i = mysql_errno( &m_mysql );//���ӳ���
		const char * s = mysql_error( &m_mysql );
		CString strMsg( s );
		//MessageBox(_T("MySQL Failed to connect! Please check for your db service."));
		MessageBox( strMsg );
	}

	//�ж����ݿ��Ƿ����
 	BOOL bRet = DBIsExist( m_mysql,"LotteryDraw" );
 	if ( !bRet )
 	{
 		//AfxMessageBox( "���ݿⲻ���ڣ�" );
		//���ݿⲻ����
		CString strSQL;
		strSQL = "CREATE DATABASE LotteryDraw";

		int nRes = mysql_query( &m_mysql,strSQL );
		if ( !nRes )
		{
			AfxMessageBox( "�ɹ��������ݿ�LotteryDraw��" );
		}
 	}

	//ѡ�����ݿ�
	mysql_select_db( &m_mysql,"LotteryDraw" );

	//���ݿ��ַ������ã������ַ���Ϊgbk,�����޷���ʾ���ĺ���
	mysql_set_character_set( &m_mysql,"gbk" );

	//�ж��Ƿ����PeopleInfo��
	BOOL bReturn = TableIsExist( m_mysql,"LotteryDraw","PeopleInfo" );
	if ( !bReturn )
	{
		//������
		CString strSQL;
		strSQL = "CREATE TABLE PeopleInfo( Sn int unsigned not null auto_increment primary key,PeopleName varchar( 256 ),IDCard varchar( 256 ),Score int unsigned,ImgPath varchar( 1000 ),Provence varchar( 256 ),City varchar( 256 ) )";

		int nRes = mysql_query( &m_mysql,strSQL );
		if ( !nRes )
		{
			AfxMessageBox( "�ɹ��������ݱ�PeopleInfo��" );
		}

		return;
	}

	//�ر��Զ��ύ
	mysql_autocommit( &m_mysql, 0 );

	//MYSQL_RES *result;
	//MYSQL_ROW sql_row;
	//MYSQL_FIELD *fd;

	//CString strTemp;
	//CString strRow;
	//CStringArray strArr;

	//int res = mysql_query( &mysql,"select * from student" );//ִ��sql����
	//if( !res )
	//{
	//	result = mysql_store_result( &mysql );//�����ѯ�������ݵ�result
	//	if( result )
	//	{
	//		int i,j;
	//		for( i = 0;fd = mysql_fetch_field( result );i++ )//��ȡ����
	//		{
	//			//str[ i ].Format( "s",fd->name );
	//			strTemp = fd->name ;
	//			//ForShow = ForShow + str[ i ] + "\t";
	//			strArr.Add( strTemp );
	//		}
	//		j = mysql_num_fields( result );//��ȡ����

	//		while( sql_row = mysql_fetch_row( result ) )//��ȡ���������
	//		{
	//			for( i = 0;i < j;i++ )
	//				//ss.Format( "s",sql_row[ i ] );
	//				strRow = sql_row[ i ];
	//		}

	//		if( result != NULL) 
	//		{
	//			mysql_free_result( result );
	//		}
	//	}
	//}
	//else
	//{
	//	MessageBox(_T("query sql failed!"));
	//}
}

 BOOL CMySQL_TestDlg::DBIsExist( MYSQL mysql,CString strDB )
 {
 	CString strSQL;
 	strSQL.Format( "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '%s'",strDB );
 
 	MYSQL_RES *result;
 
 	int res = mysql_query( &mysql,strSQL );//ִ��sql����
 	if( !res )
 	{
 		result = mysql_store_result( &mysql );//�����ѯ�������ݵ�result
		int nRow = mysql_num_rows( result );
 		if( 0 == nRow )
 		{
 			mysql_free_result( result );
 
 			AfxMessageBox( "DB isn't Exist!" );
 			return FALSE;
 		}
 		else
 		{
 			AfxMessageBox( "DB is Exist!" );
 
 			return TRUE;
 		}
 	}
 
 	return FALSE;
 }

BOOL CMySQL_TestDlg::TableIsExist( MYSQL mysql,CString strDB,CString strTable )
{
	CString strSQL;
	strSQL.Format( "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '%s' AND TABLE_NAME = '%s';",strDB,strTable );

	MYSQL_RES *result;

	int res = mysql_query( &mysql,strSQL );//ִ��sql����
	if( !res )
	{
		result = mysql_store_result( &mysql );//�����ѯ�������ݵ�result
		int nRow = mysql_num_rows( result );
		if( 0 == nRow )
		{
			mysql_free_result( result );

			AfxMessageBox( "Table isn't Exist!" );
			return FALSE;
		}
		else
		{
			AfxMessageBox( "Table is Exist!" );

			return TRUE;
		}
	}

	
	return FALSE;
}
void CMySQL_TestDlg::OnDestroy()
{
	CDialog::OnDestroy();

	// TODO: �ڴ˴������Ϣ����������
	//�ر�sql
	mysql_close( &m_mysql );
}

void CMySQL_TestDlg::OnBnClickedExceltomysql()
{
	//�����¼��
	int nInsertCount = 0;
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CFileDialog dlg( TRUE, //TRUE��FALSE��TRUEΪ���ļ���FALSEΪ�����ļ�
		_T("xls"), //Ϊȱʡ����չ��
		_T("FileList"), //Ϊ��ʾ���ļ�����Ͽ�ı༭����ļ�����һ���ѡNULL 
		OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT,//Ϊ�Ի�����һ��ΪOFN_HIDEREADONLY|OFN_OVERWRITEPROMPT,������ֻ��ѡ��͸��������ļ�ǰ��ʾ�� 
		_T("Excel �ļ�(*.xls)|*.xls||")//Ϊ�����б�������ʾ�ļ�����
		);
	dlg.m_ofn.lpstrTitle = _T("��������");

	if ( dlg.DoModal() != IDOK )
	{
		m_strStatusTips = "�򿪶Ի���ʧ�ܣ�";
		UpdateData( FALSE );

		return;
	}

	//����ļ�·����
	m_strExcelFilePath = dlg.GetPathName();
	//�ж��ļ��Ƿ��Ѿ����ڣ���������ļ�
	DWORD dwRe = GetFileAttributes( m_strExcelFilePath );
	if ( dwRe != ( DWORD ) - 1 )
	{
		//ShellExecute(NULL, NULL, m_strExcelFilePath, NULL, NULL, SW_RESTORE);//�˾�Ϊ��Excel��� 

		CDatabase db;//����Ҫ�İ����ļ������ݿ����Ҫ����ͷ�ļ� #include <afxdb.h>
		CString strDriver = _T("MICROSOFT EXCEL DRIVER (*.XLS)"); // Excel����
		CString strSQL,strArr[ 7 ];

		strSQL.Format( _T("DRIVER={%s};DSN='';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s"),strDriver, m_strExcelFilePath, m_strExcelFilePath);
		if( !db.OpenEx( strSQL,CDatabase::noOdbcDialog ) )//��������ԴDJB��xls
		{
			m_strStatusTips = "��EXCEL�ļ�ʧ��!";
			UpdateData( FALSE );

			return;
		}
		//��EXCEL��
		CRecordset pset( &db );
		m_ListPeopleInfo.DeleteAllItems();
		//�����е���Ͽ��е��б���ȫ��ɾ��

		strSQL.Format( _T("SELECT ���,����,���֤��,�÷�,ͼƬ·��,ʡ��,���� FROM [Sheet1$]") );//ע�����[Sheet1$]��д��
		pset.Open( CRecordset::forwardOnly,strSQL,CRecordset::readOnly );

		while( !pset.IsEOF() )
		{
			pset.GetFieldValue( _T("���"),strArr[ 0 ] );//ǰ���ֶα�������е���ͬ���������
			pset.GetFieldValue( _T("����"),strArr[ 1 ] );
			pset.GetFieldValue( _T("���֤��"),strArr[ 2 ] );
			pset.GetFieldValue( _T("�÷�"),strArr[ 3 ] );
			pset.GetFieldValue( _T("ͼƬ·��"),strArr[ 4 ] );
			pset.GetFieldValue( _T("ʡ��"),strArr[ 5 ] );
			pset.GetFieldValue( _T("����"),strArr[ 6 ] );

			//·�������"/"���滻��"\\"
			if ( -1 != strArr[ 4 ].Find( _T('/') ) )
			{
				strArr[ 4 ].Replace( "/","\\" );
			}

			//���뵽ListCtrl��
			int count = m_ListPeopleInfo.GetItemCount();

			m_ListPeopleInfo.InsertItem( count,strArr[ 0 ] );
			m_ListPeopleInfo.SetItemText( count,1,strArr[ 1 ] );
			m_ListPeopleInfo.SetItemText( count,2,strArr[ 2 ] );
			m_ListPeopleInfo.SetItemText( count,3,strArr[ 3 ] );
			m_ListPeopleInfo.SetItemText( count,4,strArr[ 4 ] );
			m_ListPeopleInfo.SetItemText( count,5,strArr[ 5 ] );
			m_ListPeopleInfo.SetItemText( count,6,strArr[ 6 ] );

			//ͬʱ����MySQL���ݿ�
			WriteInToMySQLDB( count,strArr,7 );

			if ( 20000 == nInsertCount )
			{
				mysql_commit( &m_mysql );
				nInsertCount = 0;
			}
			else
			{
				nInsertCount++;
			}

			pset.MoveNext();

			//ͳ���ܷ�
			//nSumScore += atoi( strArr[ 3 ].GetBuffer( strArr[ 3 ].GetLength() ) );
		}
		mysql_commit( &m_mysql );
		//�ָ��Զ��ύ
		mysql_autocommit( &m_mysql, 1 );

// 		m_nPeoCount = m_ListPeopleInfo.GetItemCount();
// 		m_nAverageScore = ( int )( nSumScore / m_nPeoCount );//ƽ����
		m_strStatusTips = "Excel���ݵ���ɹ���";

		UpdateData( FALSE );

		db.Close();

		//���б�ؼ��е����ݵ���ACCESS���ݿ�
		//AfxBeginThread( AfxListCtlToAccessDBThread,this );
	}
	else 
	{
		m_strStatusTips = "����Excel�ļ��Ƿ����!";
		UpdateData( FALSE );

		return;
	}
}

void CMySQL_TestDlg::WriteInToMySQLDB( int nCount,CString* pStr,int nArrLen )
{
	int nTemp = 0;
	CString strSQL;


	nTemp = atoi( pStr[ 3 ].GetBuffer( pStr[ 3 ].GetLength() ) );

	if ( 0 == nCount )
	{
		strSQL = "DELETE FROM PeopleInfo";
		//m_ADOSet.ExecSql( strSQL );
		mysql_query( &m_mysql,strSQL );

		strSQL.Format( "INSERT INTO PeopleInfo (Sn,PeopleName,IDCard,Score,ImgPath,Provence,City) VALUES ('%d','%s','%s','%d','%s','%s','%s')",nCount + 1,pStr[ 1 ],pStr[ 2 ],nTemp,pStr[ 4 ],pStr[ 5 ],pStr[ 6 ] );
		//m_ADOSet.ExecSql( strSQL );
		mysql_query( &m_mysql,strSQL );
	}
	else
	{
		strSQL.Format( "INSERT INTO PeopleInfo (Sn,PeopleName,IDCard,Score,ImgPath,Provence,City) VALUES ('%d','%s','%s','%d','%s','%s','%s')",nCount + 1,pStr[ 1 ],pStr[ 2 ],nTemp,pStr[ 4 ],pStr[ 5 ],pStr[ 6 ] );
		//m_ADOSet.ExecSql( strSQL );
		mysql_query( &m_mysql,strSQL );
	}
}