#if !defined(AFX_DSCDAOCTL_H__64D3F3E4_494D_11D2_BC7D_0000216A06C9__INCLUDED_)
#define AFX_DSCDAOCTL_H__64D3F3E4_494D_11D2_BC7D_0000216A06C9__INCLUDED_

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


#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "MyRecordset.h"

#include <ocdbid.h>
#include <ocdb.h>

#include "vbdsc.h"

// structure (columns) of each row of the meta data cursor
typedef struct tagDBMETADATA {
      DBCOLUMNID    ColumnID;
	   LPWSTR		szColName;
      DWORD			dwColType;
      DWORD			dwBindType;
      BOOL			bDataCol;
      ULONG			nColMaxLength;
      ULONG			nEntryIDMaxLen;
      DWORD			dwUpdatable;
} DBMETADATA, *LPDBMETADATA;

// DscdaoCtl.h : Declaration of the CDscdaoCtrl ActiveX Control class.

// Forward declaration for member variable
class CRowCursor;

// Forward declaration for friend class
class CColCursor;

/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl : See DscdaoCtl.cpp for implementation.

class CDscdaoCtrl : public COleControl
{
	friend CRowCursor;
	friend CColCursor;
	DECLARE_DYNCREATE(CDscdaoCtrl)

// Constructor
public:
	CDscdaoCtrl();


	HRESULT BindColumns(DBCOLUMNBINDING FAR **pBoundColumns, ULONG *nExistingBoundCols, ULONG cCol, DBCOLUMNBINDING rgBoundColumns[], DWORD dwFlags);
	HRESULT FindDataRowLength();


	// GUI Functions
	void DrawButtons(CDC * pdc, const CRect& rcBounds);
	void DrawButtonAt(CDC * pdc,const CRect& rcBounds,UINT btn,UINT state);
	void DrawBorder(CDC * pdc, const CRect& rcBounds);
	void DrawButtonShadow(CDC * pdc, const CRect& rcBounds, CRect rcShadow, int state);
	int GetExtraButtonCount();
	int GetHitButton(CPoint point, CRect rcBounds);
	HRESULT GetMetaData();
	HRESULT ValidateColumnBindings(ULONG cCol, DBCOLUMNBINDING rgBoundColumns[]);

	int m_CurrentButton;
	int m_nBottomBorderWidth;
	int m_nTopBorderWidth;
	CRect m_Buttons[7];

	ULONG m_nColumns;			// total # of columns in the current Resultset
	ULONG m_nBookmarks;			// total # of bookmark columns in the current Resultset
	ULONG m_cbRowSize;			// size in bytes of each row of data in the current resultset

public:
//private:
	LPUNKNOWN GetInterfaceHook(const void* piid);
	LPCONNECTIONPOINT GetConnectionHook(REFIID iid);
	
	CRowCursor *m_pRowCursor;	// ptr to the Row Cursor Object

	BEGIN_INTERFACE_PART(VBDSC,IVBDSC)
		INIT_INTERFACE_PART(CDscdaoCtrl, VBDSC)
		STDMETHOD(CancelUnload)(THIS_ BOOL FAR *pfCancel);
		STDMETHOD(Error)(THIS_ DWORD dwErr, BOOL FAR *pfShowError);
		STDMETHOD(CreateCursor)(THIS_ ICursor FAR * FAR *ppCursor);
   END_INTERFACE_PART(VBDSC)

public:

	CDaoDatabase *m_pdatabase;
	// ptr to an array of offsets to each column of data within each row of the
	// current Resultset, indexed according to that column's ordinal position
	// in that ResultSet.
	ULONG *m_pColOffsetInRow;	
	LPDBMETADATA m_pMetaData;	// ptr to meta data rows for the current ResultSet

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDscdaoCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL OnSetExtent(LPSIZEL lpSizeL);
	virtual void OnAppearanceChanged();
	//}}AFX_VIRTUAL

// helper functions
protected:


// Implementation
protected:
	~CDscdaoCtrl();


	DECLARE_OLECREATE_EX(CDscdaoCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CDscdaoCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CDscdaoCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CDscdaoCtrl)		// Type name and misc status


// Message maps
	//{{AFX_MSG(CDscdaoCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	


		 

	DECLARE_INTERFACE_MAP()
public:
// Dispatch maps
	//{{AFX_DISPATCH(CDscdaoCtrl)
	BOOL m_showAddnewButton;
	afx_msg void OnShowAddnewButtonChanged();
	BOOL m_showCancelButton;
	afx_msg void OnShowCancelButtonChanged();
	BOOL m_showUpdateButton;
	afx_msg void OnShowUpdateButtonChanged();
	CString m_databasePath;
	afx_msg void OnDatabasePathChanged();
	CString m_Table;
	afx_msg void OnTableNameChanged();
	CString m_indexName;
	afx_msg void OnIndexNameChanged();
	BOOL m_readOnly;
	afx_msg void OnReadOnlyChanged();
	short m_openType;
	afx_msg void OnOpenTypeChanged();
	short m_sqlType;
	afx_msg void OnSqlTypeChanged();
	CString m_sqlString;
	afx_msg void OnSqlStringChanged();
	OLE_COLOR m_drawColor;
	afx_msg void OnDrawColorChanged();
	BOOL m_openExclusive;
	afx_msg void OnOpenExclusiveChanged();
	CString m_password;
	afx_msg void OnPasswordChanged();
	BOOL m_lockingMode;
	afx_msg void OnLockingModeChanged();
	afx_msg void Next();
	afx_msg void Previous();
	afx_msg void Beginning();
	afx_msg void End();
	afx_msg void AddNew();
	afx_msg void Delete();
	afx_msg void CancelUpdate();
	afx_msg void Update();
	afx_msg void BeginTransaction();
	afx_msg void CommitTransaction();
	afx_msg void Rollback();
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

//	afx_msg void AboutBox();

// Event maps
	//{{AFX_EVENT(CDscdaoCtrl)
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	BOOL reset;
	enum {
	//{{AFX_DISP_ID(CDscdaoCtrl)
	dispidShowAddnewButton = 1L,
	dispidShowCancelButton = 2L,
	dispidShowUpdateButton = 3L,
	dispidDatabasePath = 4L,
	dispidTableName = 5L,
	dispidIndexName = 6L,
	dispidReadOnly = 7L,
	dispidOpenType = 8L,
	dispidSqlType = 9L,
	dispidSqlString = 10L,
	dispidDrawColor = 11L,
	dispidOpenExclusive = 12L,
	dispidPassword = 13L,
	dispidLockingMode = 14L,
	dispidNext = 15L,
	dispidPrevious = 16L,
	dispidBeginning = 17L,
	dispidEnd = 18L,
	dispidAddNew = 19L,
	dispidDelete = 20L,
	dispidCancelUpdate = 21L,
	dispidUpdate = 22L,
	dispidBeginTransaction = 23L,
	dispidCommitTransaction = 24L,
	dispidRollback = 25L,
	//}}AFX_DISP_ID
	};
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DSCDAOCTL_H__64D3F3E4_494D_11D2_BC7D_0000216A06C9__INCLUDED)
