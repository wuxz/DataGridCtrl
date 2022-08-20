// ColCrsr.h : header file
//

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



#if !defined( _COLCRSR_H_ )
#define _COLCRSR_H_

// Forward declarations for back pointers
class CDscdaoCtrl;
class CRowCursor;

/////////////////////////////////////////////////////////////////////////////
// CColCursor command target

class CColCursor : public CCmdTarget
{
	friend CRowCursor;
	
	DECLARE_DYNCREATE(CColCursor)

	// protected constructor used by dynamic creation
	CColCursor(CDscdaoCtrl *m_pControl = NULL, CRowCursor *m_pRowCursor = NULL);

// Attributes
public:
	
	virtual ~CColCursor();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CColCursor)
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CColCursor)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

public:

	BEGIN_INTERFACE_PART(COLUMNSCURSOR, ICursor)
		INIT_INTERFACE_PART(CColCursor, COLUMNSCURSOR)
		STDMETHOD(GetColumnsCursor)(THIS_ REFIID riid,IUnknown **ppvColumnsCursor,ULONG *pcRows);
		STDMETHOD(SetBindings)(THIS_ ULONG cCol,	DBCOLUMNBINDING rgBoundColumns[],ULONG cbRowLength,DWORD dwFlags);
		STDMETHOD(GetBindings)(THIS_ ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength);
		STDMETHOD(GetNextRows)(THIS_ LARGE_INTEGER udlRowsToSkip,	DBFETCHROWS *pFetchParams);
		STDMETHOD(Requery)(THIS_);
	END_INTERFACE_PART(COLUMNSCURSOR)

	BEGIN_INTERFACE_PART(COLUMNSCURSORMOVE, ICursorMove)
		INIT_INTERFACE_PART(CColCursor, COLUMNSCURSORMOVE)
		STDMETHOD(GetColumnsCursor)(THIS_ REFIID riid,IUnknown **ppvColumnsCursor,ULONG *pcRows);
		STDMETHOD(SetBindings)(THIS_ ULONG cCol,	DBCOLUMNBINDING rgBoundColumns[],ULONG cbRowLength,DWORD dwFlags);
		STDMETHOD(GetBindings)(THIS_ ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength);
		STDMETHOD(GetNextRows)(THIS_ LARGE_INTEGER udlRowsToSkip,	DBFETCHROWS *pFetchParams);
		STDMETHOD(Requery)(THIS_);
		STDMETHOD(Move)(THIS_ ULONG cbBookmark, void *pBookmark, LARGE_INTEGER dlOffset, DBFETCHROWS *pFetchParams);
		STDMETHOD(GetBookmark)(THIS_ DBCOLUMNID *pBookmarkType, ULONG cbMaxSize, ULONG *pcbBookmark, void *pBookmark);
		STDMETHOD(Clone)(THIS_ DWORD dwFlags, REFIID riid, IUnknown **ppvClonedCursor);
	END_INTERFACE_PART(COLUMNSCURSORMOVE)

	DECLARE_INTERFACE_MAP()

// Helper Functions
public:
	void Close();
	
	HRESULT FetchRows(DBFETCHROWS *pFetchParams);

private:

	CDscdaoCtrl *m_pControl;				// ptr to the Control Class object
	CRowCursor *m_pRowCursor;				// ptr to the Row Cursor Object
	
	ULONG m_lMetaDataRowLength;				// Size in bytes of each row of meta data
	ULONG m_CurrentMetaDataRow;				// Current Row Position in meta data cursor

	DBCOLUMNBINDING *m_pMetaDataColumns;	// ptr to an array of Column Bindings for meta data
	ULONG m_nExistingMetaDataCols;			// number of existing bound meta data columns
};

#endif
