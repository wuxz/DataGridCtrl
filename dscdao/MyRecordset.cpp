// MyRecordset.cpp: implementation of the CMyRecordset class.
//
//////////////////////////////////////////////////////////////////////

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
#include "MyRecordset.h"

#include "dscdaoctl.h"



IMPLEMENT_DYNAMIC(CMyDaoRecordset, CDaoRecordset)

#ifdef _DEBUG
long int CMyDaoRecordset::count=0;
#endif
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMyDaoRecordset::CMyDaoRecordset(CDscdaoCtrl *ctrl,CDaoDatabase* pDatabase)
	: ctl(ctrl), CDaoRecordset(pDatabase)
{
#ifdef _DEBUG
count++;
#endif
}

CMyDaoRecordset::~CMyDaoRecordset()
{
#ifdef _DEBUG
count--;
TRACE("\n%d recordsets left\n",count);
#endif
}


CString CMyDaoRecordset::GetDefaultDBName()
{
return ctl->m_databasePath;
}


CString  CMyDaoRecordset::GetDefaultSQL()
{
return CString("[")+ctl->m_Table+CString("]");
}




COleVariant CMyDaoRecordset::GetBookmark()
{
if(GetType()!=dbOpenSnapshot)return CDaoRecordset::GetBookmark();
long pos=GetAbsolutePosition();
char str[34];
ltoa(pos,str,10);
VARIANT vari;
VariantInit(&vari);
vari.vt=VT_UI1|VT_ARRAY;
SAFEARRAY FAR *psaTmp;
SAFEARRAYBOUND sab[1];
VOID HUGEP *pData;
sab[0].cElements=strlen(str)+1;
sab[0].lLbound=0L;
psaTmp=SafeArrayCreate(VT_UI1,1,sab);
SafeArrayAccessData(psaTmp, &pData);
strcpy((char*)pData,str);
SafeArrayUnaccessData(psaTmp);
vari.parray=psaTmp;
COleVariant var;
var.Attach(vari);
return var;
}

void CMyDaoRecordset::SetBookmark(COleVariant bok)
{
if(GetType()!=dbOpenSnapshot){
	CDaoRecordset::SetBookmark(bok);
	return;
}
VOID HUGEP *pData;
SafeArrayAccessData(bok.parray,&pData);
try{
   SetAbsolutePosition(atol((char*)pData));
}catch(CDaoException *){
   SafeArrayUnaccessData(bok.parray);   
   throw;
}
SafeArrayUnaccessData(bok.parray);
}

CMyDaoRecordset* CMyDaoRecordset::Clone()
{
if(!m_pDAORecordset || !IsOpen())AfxThrowDaoException();
CMyDaoRecordset *set=new CMyDaoRecordset(ctl,m_pDatabase);

DAO_CHECK(m_pDAORecordset->Clone(&set->m_pDAORecordset));

set->m_bOpen=m_bOpen;
set->m_strFilter=m_strFilter;    // Filter string used when constructing SQL
set->m_strSort=m_strSort;      // Sort string used when constructing SQL
set->m_nFields=m_nFields;
set->m_nParams=m_nParams;
set->m_bCheckCacheForDirtyFields=m_bCheckCacheForDirtyFields;   // Switch for dirty field checking.
set->m_strSQL=m_strSQL;
set->m_cbFixedLengthFields=m_cbFixedLengthFields;
set->m_nStatus=m_nStatus;
set->m_bAppendable=m_bAppendable;
set->m_bScrollable=m_bScrollable;
set->m_bDeleted=m_bDeleted;
set->m_nOpenType=m_nOpenType;
set->m_nDefaultType=m_nDefaultType;
set->m_nOptions=m_nOptions;
set->m_strRequerySQL=m_strRequerySQL;
set->m_strRequeryFilter=m_strRequeryFilter;
set->m_strRequerySort=m_strRequerySort;


//set->m_DaoFetchRows=m_DaoFetchRows;

//if(set->m_nStatus & AFX_DAO_IMPLICIT_TD)set->m_pQueryDef=m_pQueryDef;  // Source query for this result set
//if(set->m_nStatus & AFX_DAO_IMPLICIT_QD)set->m_pTableDef=m_pTableDef;

set->m_pDatabase->m_mapRecordsets.SetAt(set, set);
/*
	TRY
	{
		set->BindFields();   //no bind fields in this case
		set->GetDataAndFixupNulls();  //no bind fields in our case
		set->SetCursorAttributes();
	}
	CATCH_ALL(e)
	{
		set->Close();
		THROW_LAST();
	}
	END_CATCH_ALL
*/
try{
   set->SetCurrentIndex(GetCurrentIndex());
}catch(CDaoException *e){
   e->Delete();
}
try{
   set->SetBookmark(GetBookmark());
}catch(CDaoException *e){
   e->Delete();
}
if(set->IsOpen())set->SetLockingMode(ctl->m_lockingMode);
return set;
}

void CMyDaoRecordset::Reclone(CMyDaoRecordset * set)
{
if(!set->m_pDAORecordset || !set->IsOpen())AfxThrowDaoException();
if(IsOpen())Close();
DAO_CHECK(set->m_pDAORecordset->Clone(&m_pDAORecordset));
m_bOpen=set->m_bOpen;


m_strFilter=set->m_strFilter;    // Filter string used when constructing SQL
m_strSort=set->m_strSort;      // Sort string used when constructing SQL
m_nFields=set->m_nFields;
m_nParams=set->m_nParams;
m_bCheckCacheForDirtyFields=set->m_bCheckCacheForDirtyFields;   // Switch for dirty field checking.
m_strSQL=set->m_strSQL;
m_cbFixedLengthFields=set->m_cbFixedLengthFields;
m_nStatus=set->m_nStatus;
m_bAppendable=set->m_bAppendable;
m_bScrollable=set->m_bScrollable;
m_bDeleted=set->m_bDeleted;
m_nOpenType=set->m_nOpenType;
m_nDefaultType=set->m_nDefaultType;
m_nOptions=set->m_nOptions;
m_strRequerySQL=set->m_strRequerySQL;
m_strRequeryFilter=set->m_strRequeryFilter;
m_strRequerySort=set->m_strRequerySort;

//m_DaoFetchRows=set->m_DaoFetchRows;

//if(m_nStatus & AFX_DAO_IMPLICIT_TD)m_pQueryDef=set->m_pQueryDef;  // Source query for this result set
//if(m_nStatus & AFX_DAO_IMPLICIT_QD)m_pTableDef=set->m_pTableDef;



	// Add the recordset to map of Open CDaoRecordsets
//	m_pDatabase->m_mapRecordsets.SetAt(this, this);
/*
	TRY
	{
		BindFields();  //no bind fields in our case
		GetDataAndFixupNulls();  //no bind fields in our case
		SetCursorAttributes();
	}
	CATCH_ALL(e)
	{
		Close();
		THROW_LAST();
	}
	END_CATCH_ALL
*/
//	ICDAORecordset* m_pICDAORecordsetGetRows;
//	DAOFields* m_pDAOFields;
//	DAOIndexes* m_pDAOIndexes;

//	DAOCOLUMNBINDING* m_prgDaoColBindInfo;
//	DWORD* m_pulColumnLengths;
	
//	BYTE* m_pbFieldFlags;
//	BYTE* m_pbParamFlags;

//	CMapPtrToPtr* m_pMapFieldCache;
//	CMapPtrToPtr* m_pMapFieldIndex;


try{
   SetCurrentIndex(set->GetCurrentIndex());
}catch(CDaoException *e){
   e->Delete();
}
try{
   SetBookmark(set->GetBookmark());
}catch(CDaoException *e){
   e->Delete();
}
if(IsOpen())SetLockingMode(ctl->m_lockingMode);
}


void CMyDaoRecordset::Open( int nOpenType, LPCTSTR lpszSQL, int nOptions )
{
CDaoRecordset::Open( nOpenType, lpszSQL, nOptions );
if(IsOpen())SetLockingMode(ctl->m_lockingMode);
}

void CMyDaoRecordset::Open( CDaoTableDef* pTableDef, int nOpenType, int nOptions)
{
CDaoRecordset::Open( pTableDef, nOpenType, nOptions);
if(IsOpen())SetLockingMode(ctl->m_lockingMode);
}

void CMyDaoRecordset::Open( CDaoQueryDef* pQueryDef, int nOpenType, int nOptions)
{
CDaoRecordset::Open( pQueryDef, nOpenType, nOptions);
if(IsOpen())SetLockingMode(ctl->m_lockingMode);
}

void CMyDaoRecordset::SetLockingMode(BOOL bPessimistic)
{
try{
   CDaoRecordset::SetLockingMode(bPessimistic);
}catch(CDaoException *e){
   e->Delete();
}
}
