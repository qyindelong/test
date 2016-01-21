
// MySQL_TestDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "MySQL_Test.h"
#include "MySQL_TestDlg.h"

#include <afxdb.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CMySQL_TestDlg 对话框




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


// CMySQL_TestDlg 消息处理程序

BOOL CMySQL_TestDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	CRect rect;

	//初始化参与抽奖人员信息列表控件
	m_ListPeopleInfo.GetClientRect( rect );
	m_ListPeopleInfo.SetExtendedStyle( m_ListPeopleInfo.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES );

	m_ListPeopleInfo.InsertColumn( 0,_T("序号"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 1,_T("姓名"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 2,_T("身份证号"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 3,_T("得分"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 4,_T("图片路径"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 5,_T("省份"),LVCFMT_CENTER,rect.Width()/7,0 );
	m_ListPeopleInfo.InsertColumn( 6,_T("地区"),LVCFMT_CENTER,rect.Width()/7,0 );


	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CMySQL_TestDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CMySQL_TestDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CMySQL_TestDlg::OnBnClickedBtnSqlOpen()
{
	// TODO: 在此添加控件通知处理程序代码

	//初始化
	mysql_init( &m_mysql );

	//连接数据库
	bool isConnected = mysql_real_connect( &m_mysql,"localhost","root","bjgas","",3306,NULL,0 );
	if ( isConnected )
	{
		MessageBox(_T("MySQL Success Connected!"));
		//return;
	}
	else
	{
		int i = mysql_errno( &m_mysql );//连接出错
		const char * s = mysql_error( &m_mysql );
		CString strMsg( s );
		//MessageBox(_T("MySQL Failed to connect! Please check for your db service."));
		MessageBox( strMsg );
	}

	//判断数据库是否存在
 	BOOL bRet = DBIsExist( m_mysql,"LotteryDraw" );
 	if ( !bRet )
 	{
 		//AfxMessageBox( "数据库不存在！" );
		//数据库不存在
		CString strSQL;
		strSQL = "CREATE DATABASE LotteryDraw";

		int nRes = mysql_query( &m_mysql,strSQL );
		if ( !nRes )
		{
			AfxMessageBox( "成功创建数据库LotteryDraw！" );
		}
 	}

	//选择数据库
	mysql_select_db( &m_mysql,"LotteryDraw" );

	//数据库字符集设置，设置字符库为gbk,否则无法显示中文汉字
	mysql_set_character_set( &m_mysql,"gbk" );

	//判断是否存在PeopleInfo表
	BOOL bReturn = TableIsExist( m_mysql,"LotteryDraw","PeopleInfo" );
	if ( !bReturn )
	{
		//表不存在
		CString strSQL;
		strSQL = "CREATE TABLE PeopleInfo( Sn int unsigned not null auto_increment primary key,PeopleName varchar( 256 ),IDCard varchar( 256 ),Score int unsigned,ImgPath varchar( 1000 ),Provence varchar( 256 ),City varchar( 256 ) )";

		int nRes = mysql_query( &m_mysql,strSQL );
		if ( !nRes )
		{
			AfxMessageBox( "成功创建数据表PeopleInfo！" );
		}

		return;
	}

	//关闭自动提交
	mysql_autocommit( &m_mysql, 0 );

	//MYSQL_RES *result;
	//MYSQL_ROW sql_row;
	//MYSQL_FIELD *fd;

	//CString strTemp;
	//CString strRow;
	//CStringArray strArr;

	//int res = mysql_query( &mysql,"select * from student" );//执行sql命令
	//if( !res )
	//{
	//	result = mysql_store_result( &mysql );//保存查询到的数据到result
	//	if( result )
	//	{
	//		int i,j;
	//		for( i = 0;fd = mysql_fetch_field( result );i++ )//获取列名
	//		{
	//			//str[ i ].Format( "s",fd->name );
	//			strTemp = fd->name ;
	//			//ForShow = ForShow + str[ i ] + "\t";
	//			strArr.Add( strTemp );
	//		}
	//		j = mysql_num_fields( result );//获取列数

	//		while( sql_row = mysql_fetch_row( result ) )//获取具体的数据
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
 
 	int res = mysql_query( &mysql,strSQL );//执行sql命令
 	if( !res )
 	{
 		result = mysql_store_result( &mysql );//保存查询到的数据到result
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

	int res = mysql_query( &mysql,strSQL );//执行sql命令
	if( !res )
	{
		result = mysql_store_result( &mysql );//保存查询到的数据到result
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

	// TODO: 在此处添加消息处理程序代码
	//关闭sql
	mysql_close( &m_mysql );
}

void CMySQL_TestDlg::OnBnClickedExceltomysql()
{
	//插入记录数
	int nInsertCount = 0;
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog dlg( TRUE, //TRUE或FALSE。TRUE为打开文件；FALSE为保存文件
		_T("xls"), //为缺省的扩展名
		_T("FileList"), //为显示在文件名组合框的编辑框的文件名，一般可选NULL 
		OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT,//为对话框风格，一般为OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT,即隐藏只读选项和覆盖已有文件前提示。 
		_T("Excel 文件(*.xls)|*.xls||")//为下拉列表枢中显示文件类型
		);
	dlg.m_ofn.lpstrTitle = _T("导入数据");

	if ( dlg.DoModal() != IDOK )
	{
		m_strStatusTips = "打开对话框失败！";
		UpdateData( FALSE );

		return;
	}

	//获得文件路径名
	m_strExcelFilePath = dlg.GetPathName();
	//判断文件是否已经存在，存在则打开文件
	DWORD dwRe = GetFileAttributes( m_strExcelFilePath );
	if ( dwRe != ( DWORD ) - 1 )
	{
		//ShellExecute(NULL, NULL, m_strExcelFilePath, NULL, NULL, SW_RESTORE);//此句为打开Excel表格 

		CDatabase db;//（重要的包含文件）数据库库需要包含头文件 #include <afxdb.h>
		CString strDriver = _T("MICROSOFT EXCEL DRIVER (*.XLS)"); // Excel驱动
		CString strSQL,strArr[ 7 ];

		strSQL.Format( _T("DRIVER={%s};DSN='';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s"),strDriver, m_strExcelFilePath, m_strExcelFilePath);
		if( !db.OpenEx( strSQL,CDatabase::noOdbcDialog ) )//连接数据源DJB．xls
		{
			m_strStatusTips = "打开EXCEL文件失败!";
			UpdateData( FALSE );

			return;
		}
		//打开EXCEL表
		CRecordset pset( &db );
		m_ListPeopleInfo.DeleteAllItems();
		//将已有的组合框中的列表项全部删除

		strSQL.Format( _T("SELECT 序号,姓名,身份证号,得分,图片路径,省份,地区 FROM [Sheet1$]") );//注意表名[Sheet1$]的写法
		pset.Open( CRecordset::forwardOnly,strSQL,CRecordset::readOnly );

		while( !pset.IsEOF() )
		{
			pset.GetFieldValue( _T("序号"),strArr[ 0 ] );//前面字段必须与表中的相同，否则出错。
			pset.GetFieldValue( _T("姓名"),strArr[ 1 ] );
			pset.GetFieldValue( _T("身份证号"),strArr[ 2 ] );
			pset.GetFieldValue( _T("得分"),strArr[ 3 ] );
			pset.GetFieldValue( _T("图片路径"),strArr[ 4 ] );
			pset.GetFieldValue( _T("省份"),strArr[ 5 ] );
			pset.GetFieldValue( _T("地区"),strArr[ 6 ] );

			//路径如果是"/"则替换成"\\"
			if ( -1 != strArr[ 4 ].Find( _T('/') ) )
			{
				strArr[ 4 ].Replace( "/","\\" );
			}

			//插入到ListCtrl中
			int count = m_ListPeopleInfo.GetItemCount();

			m_ListPeopleInfo.InsertItem( count,strArr[ 0 ] );
			m_ListPeopleInfo.SetItemText( count,1,strArr[ 1 ] );
			m_ListPeopleInfo.SetItemText( count,2,strArr[ 2 ] );
			m_ListPeopleInfo.SetItemText( count,3,strArr[ 3 ] );
			m_ListPeopleInfo.SetItemText( count,4,strArr[ 4 ] );
			m_ListPeopleInfo.SetItemText( count,5,strArr[ 5 ] );
			m_ListPeopleInfo.SetItemText( count,6,strArr[ 6 ] );

			//同时导入MySQL数据库
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

			//统计总分
			//nSumScore += atoi( strArr[ 3 ].GetBuffer( strArr[ 3 ].GetLength() ) );
		}
		mysql_commit( &m_mysql );
		//恢复自动提交
		mysql_autocommit( &m_mysql, 1 );

// 		m_nPeoCount = m_ListPeopleInfo.GetItemCount();
// 		m_nAverageScore = ( int )( nSumScore / m_nPeoCount );//平均分
		m_strStatusTips = "Excel数据导入成功！";

		UpdateData( FALSE );

		db.Close();

		//将列表控件中的数据导入ACCESS数据库
		//AfxBeginThread( AfxListCtlToAccessDBThread,this );
	}
	else 
	{
		m_strStatusTips = "请检查Excel文件是否存在!";
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