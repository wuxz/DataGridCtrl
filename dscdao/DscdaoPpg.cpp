// DscdaoPpg.cpp : Implementation of the CDscdaoPropPage property page class.

/*
Copyright Adrian Roman, aroman@medanet.ro

You can use the code free of charge, you can modify it, but the author (Adrian Roman) 
is not responsible of any kind of damage or loss of data or loss of profit, 
incidental or consequential, occurred using this code. 
You cannot claim that the code is written by yourself, even the code is modified.
If you use this control, you must make a notice (in About box and/or startup splash screen 
and/or help file) that the program contains code developed by Adrian Roman, 
e-mail: aroman@medanet.ro.
*/



#include "stdafx.h"
#include "dscdao.h"
#include "DscdaoPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CDscdaoPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CDscdaoPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CDscdaoPropPage)
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_CBN_SELCHANGE(IDC_COMBO3, OnSelchangeCombo3)
	ON_CBN_SELCHANGE(IDC_COMBO4, OnSelchangeCombo4)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CDscdaoPropPage, "DSCDAO.DscdaoPropPage.1",
	0x64d3f3d7, 0x494d, 0x11d2, 0xbc, 0x7d, 0, 0, 0x21, 0x6a, 0x6, 0xc9)


/////////////////////////////////////////////////////////////////////////////
// CDscdaoPropPage::CDscdaoPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CDscdaoPropPage

BOOL CDscdaoPropPage::CDscdaoPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),m_clsid,IDS_DSCDAO_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoPropPage::CDscdaoPropPage - Constructor

CDscdaoPropPage::CDscdaoPropPage() :
	COlePropertyPage(IDD, IDS_DSCDAO_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CDscdaoPropPage)
	m_databasePath = _T("");
	m_Table = _T("");
	m_indexName = _T("");
	m_ReadOnly = FALSE;
	m_openType = -1;
	m_sqlType = -1;
	m_SqlString = _T("");
	m_lockingMode = FALSE;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoPropPage::DoDataExchange - Moves data between page and properties

void CDscdaoPropPage::DoDataExchange(CDataExchange* pDX)
{
   //{{AFX_DATA_MAP(CDscdaoPropPage)
	DDX_Control(pDX, IDC_EDIT2, m_Edit2);
	DDX_Control(pDX, IDC_COMBO4, m_Combo4);
	DDX_Control(pDX, IDC_COMBO3, m_Combo3);
	DDX_Control(pDX, IDC_CHECK1, m_Check1);
	DDX_Control(pDX, IDC_COMBO2, m_Index);
	DDX_Control(pDX, IDC_COMBO1, m_Combo);
	DDP_Text(pDX, IDC_EDIT1, m_databasePath, _T("DatabasePath") );
	DDX_Text(pDX, IDC_EDIT1, m_databasePath);
	DDP_CBString(pDX, IDC_COMBO1, m_Table, _T("TableName") );
	DDX_CBString(pDX, IDC_COMBO1, m_Table);
	DDP_CBString(pDX, IDC_COMBO2, m_indexName, _T("IndexName") );
	DDX_CBString(pDX, IDC_COMBO2, m_indexName);
	DDP_Check(pDX, IDC_CHECK1, m_ReadOnly, _T("ReadOnly") );
	DDX_Check(pDX, IDC_CHECK1, m_ReadOnly);
	DDP_CBIndex(pDX, IDC_COMBO3, m_openType, _T("OpenType") );
	DDX_CBIndex(pDX, IDC_COMBO3, m_openType);
	DDP_CBIndex(pDX, IDC_COMBO4, m_sqlType, _T("SqlType") );
	DDX_CBIndex(pDX, IDC_COMBO4, m_sqlType);
	DDP_Text(pDX, IDC_EDIT2, m_SqlString, _T("SqlString") );
	DDX_Text(pDX, IDC_EDIT2, m_SqlString);
	DDP_Check(pDX, IDC_CHECK2, m_lockingMode, _T("LockingMode") );
	DDX_Check(pDX, IDC_CHECK2, m_lockingMode);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoPropPage message handlers

void CDscdaoPropPage::OnButton1() 
{
	// TODO: Add your control notification handler code here
	CFileDialog dlg(TRUE,NULL,m_databasePath);
	dlg.m_ofn.lpstrFilter="MDB files\0*.mdb\0\0";
	dlg.DoModal();
	m_databasePath=dlg.GetPathName();
	m_Combo.ResetContent();
	CDaoDatabase database;
	try{
		database.Open(m_databasePath,FALSE,TRUE,_T(";"));
	}catch(CDaoException *e){
		e->Delete();
		return;
	}
	BOOL first_time=TRUE;
	for(int i=0;i<database.GetTableDefCount();i++){
		CDaoTableDefInfo def;
		database.GetTableDefInfo(i,def);
		if(!(def.m_lAttributes & dbSystemObject) && !(def.m_lAttributes & dbHiddenObject)){
			m_Combo.AddString(def.m_strName);
			if(first_time){
				m_Table=def.m_strName;
				first_time=FALSE;
			}
		}
	}
	SetDlgItemText(IDC_EDIT1,m_databasePath);
	m_Combo.SetCurSel(0);
	SetDlgItemText(IDC_COMBO1,m_Table);
	SetControlStatus(IDC_COMBO1,TRUE);
   OnSelchangeCombo4();
}


BOOL CDscdaoPropPage::OnInitDialog() 
{
	COlePropertyPage::OnInitDialog();
	
	// TODO: Add extra initialization here
   OnSelchangeCombo4();
   return FALSE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}



void CDscdaoPropPage::OnSelchangeCombo3() 
{
	// TODO: Add your control notification handler code here
int show,show_index;
show=SW_SHOW;
show_index=SW_HIDE;
int cur=m_Combo3.GetCurSel();
if(cur==0 && m_Combo3.GetCount()==3){//table
   show_index=SW_SHOW;
}else if((cur==2 && m_Combo3.GetCount()==3) || (cur==1 && m_Combo3.GetCount()==2)){//snapshot
   show=SW_HIDE;  
}

if(show==SW_HIDE){
   m_ReadOnly=TRUE;
   m_Check1.SetCheck(1);
}

if(show_index==SW_HIDE){
   m_indexName=_T("None");
   m_Index.SetCurSel(0);
}

m_Check1.ShowWindow(show);
m_Index.ShowWindow(show_index);
}

void CDscdaoPropPage::OnSelchangeCombo4() 
{
	// TODO: Add your control notification handler code here
int cur=m_Combo4.GetCurSel();
if(cur==0){//table
  m_Edit2.ShowWindow(SW_HIDE);
  m_Index.ShowWindow(SW_SHOW);
  m_Combo.ShowWindow(SW_SHOW);
   if(m_Combo3.GetCount()==2){
     m_Combo3.InsertString(0,_T("Table"));
	  m_Combo3.SetCurSel(0);	
   }
}else if(cur==1){//Stored query
  m_Edit2.ShowWindow(SW_HIDE);
  m_Index.ShowWindow(SW_HIDE);
  m_Combo.ShowWindow(SW_SHOW);
  if(m_Combo3.GetCount()==3){
     m_Combo3.DeleteString(0);
	  m_Combo3.SetCurSel(0);	
  }
}else{//local query
  m_Index.ShowWindow(SW_HIDE);
  if(m_Combo3.GetCount()==3){
     m_Combo3.DeleteString(0);
	  m_Combo3.SetCurSel(0);	
  }
  m_Edit2.ShowWindow(SW_SHOW);
  m_Combo.ShowWindow(SW_HIDE);
}
if(cur<2){
m_Combo.ResetContent();
CDaoDatabase database;
try{
	database.Open(m_databasePath,FALSE,TRUE,_T(";"));
}catch(CDaoException *e){
	e->Delete();
	return;
}
BOOL first_time=TRUE;
   if(cur==0){
	for(int i=0;i<database.GetTableDefCount();i++){
		CDaoTableDefInfo def;
		database.GetTableDefInfo(i,def);
		if(!(def.m_lAttributes & dbSystemObject) && !(def.m_lAttributes & dbHiddenObject)){
			m_Combo.AddString(def.m_strName);
			if(first_time){
				m_Table=def.m_strName;
				first_time=FALSE;
			}
		}
	}
   CDaoTableDef tableDef(&database);
	try{
		tableDef.Open(m_Table);
	}catch(CDaoException *e){
		e->Delete();
		return;
	}
   m_Index.ResetContent();
	m_Index.AddString("None");
	for(i=0;i<tableDef.GetIndexCount();i++){
		CDaoIndexInfo info;
		tableDef.GetIndexInfo(i,info);
		m_Index.AddString(info.m_strName);
	}
	m_Index.SetCurSel(0);	
	SetDlgItemText(IDC_COMBO2,"None");
	SetControlStatus(IDC_COMBO2,TRUE);

   m_Combo.SetCurSel(0);
   SetDlgItemText(IDC_COMBO1,m_Table);
   SetControlStatus(IDC_COMBO1,TRUE);

   }else if(cur==1){
	for(int i=0;i<database.GetQueryDefCount();i++){
		CDaoQueryDefInfo def;
		database.GetQueryDefInfo(i,def);
      if(!(def.m_nType & dbQSelect)){
         CDaoQueryDef query(&database);
         query.Open(def.m_strName);
         if(!query.GetParameterCount()){
            m_Combo.AddString(def.m_strName);
			   if(first_time){
				   m_Table=def.m_strName;
				   first_time=FALSE;
			   }
         }
         query.Close();
      }
	}
   m_Combo.SetCurSel(0);
   SetDlgItemText(IDC_COMBO1,m_Table);
   SetControlStatus(IDC_COMBO1,TRUE);
   }
}
OnSelchangeCombo3();
}


