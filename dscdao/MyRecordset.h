// MyRecordset.h: interface for the CMyRecordset class.
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


#if !defined(AFX_MYRECORDSET_H__887E70D0_3854_11D2_BBF0_0000216A06C9__INCLUDED_)
#define AFX_MYRECORDSET_H__887E70D0_3854_11D2_BBF0_0000216A06C9__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class CDscdaoCtrl;

class CMyDaoRecordset : public CDaoRecordset
{
public:
	inline void SetLockingMode( BOOL bPessimistic );
   virtual void Open( int nOpenType = AFX_DAO_USE_DEFAULT_TYPE, LPCTSTR lpszSQL = NULL, int nOptions = 0 );
   virtual void Open( CDaoTableDef* pTableDef, int nOpenType = dbOpenTable, int nOptions = 0 );
   virtual void Open( CDaoQueryDef* pQueryDef, int nOpenType = dbOpenDynaset, int nOptions = 0 );

	void Reclone(CMyDaoRecordset *set);
	CMyDaoRecordset* Clone();
	void SetBookmark(COleVariant bok);
	COleVariant GetBookmark();
	CDscdaoCtrl * ctl;
	CMyDaoRecordset(CDscdaoCtrl *ctrl,CDaoDatabase* pDatabase = NULL);
   virtual ~CMyDaoRecordset();
   DECLARE_DYNAMIC(CMyDaoRecordset)

// Field/Param Data
	//{{AFX_FIELD(CMyDaoRecordset, CDaoRecordset)
	//}}AFX_FIELD

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMyDaoRecordset)
	public:
	virtual CString GetDefaultDBName();		// Default database name
	virtual CString GetDefaultSQL( );
	//}}AFX_VIRTUAL

// Implementation
#ifdef _DEBUG
   static long int count;
#endif
};

#endif // !defined(AFX_MYRECORDSET_H__887E70D0_3854_11D2_BBF0_0000216A06C9__INCLUDED_)
