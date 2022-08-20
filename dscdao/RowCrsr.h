// RowCrsr.h : header file
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


#if !defined( _ROWCRSR_H_ )
#define _ROWCRSR_H_

// macro helpful for coercing data types
#define CVTFIXEDTYPE(FROMTYPE, TOTYPE, VARFIELD)										\
																						\
		if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_VARIANT)					\
		{																				\
			((DBVARIANT *)pDestColVal)->vt = (VARTYPE)dwRequestedType;					\
			((DBVARIANT *)pDestColVal)->VARFIELD = (TOTYPE)(*(FROMTYPE *)pSrcColVal);	\
		}																				\
		else if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_DEFAULT)				\
			*(TOTYPE *)pDestColVal = (TOTYPE)(*(FROMTYPE *)pSrcColVal);

// Forward declarations for back pointers
class CDscdaoCtrl;
class CColCursor;

/////////////////////////////////////////////////////////////////////////////
// CRowCursor command target

class CRowCursor : public CCmdTarget
{
	friend CColCursor;
	friend CDscdaoCtrl;
	
	DECLARE_DYNCREATE(CRowCursor)

public:

	// protected constructor used by dynamic creation
	CRowCursor(CDscdaoCtrl *pControl = NULL,BOOL create_cursor=TRUE);           

// Attributes
public:
	LPCONNECTIONPOINT GetConnectionHook(REFIID iid);

// Operations
public:
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CRowCursor)
	//}}AFX_VIRTUAL

// Implementation
protected:
	
	virtual ~CRowCursor();

	// Generated message map functions
	//{{AFX_MSG(CRowCursor)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

public:

	BEGIN_INTERFACE_PART(CURSOR, ICursor)
		INIT_INTERFACE_PART(CRowCursor, CURSOR)
		STDMETHOD(GetColumnsCursor)(THIS_ REFIID riid,IUnknown **ppvColumnsCursor,ULONG *pcRows);
		STDMETHOD(SetBindings)(THIS_ ULONG cCol, DBCOLUMNBINDING rgBoundColumns[],ULONG cbRowLength,DWORD dwFlags);
		STDMETHOD(GetBindings)(THIS_ ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength);
		STDMETHOD(GetNextRows)(THIS_ LARGE_INTEGER udlRowsToSkip, DBFETCHROWS *pFetchParams);
		STDMETHOD(Requery)(THIS_);
	END_INTERFACE_PART(CURSOR)


	BEGIN_INTERFACE_PART(CURSORMOVE, ICursorMove)
		INIT_INTERFACE_PART(CRowCursor, CURSORMOVE)
		STDMETHOD(GetColumnsCursor)(THIS_ REFIID riid,IUnknown **ppvColumnsCursor,ULONG *pcRows);
		STDMETHOD(SetBindings)(THIS_ ULONG cCol, DBCOLUMNBINDING rgBoundColumns[],ULONG cbRowLength,DWORD dwFlags);
		STDMETHOD(GetBindings)(THIS_ ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength);
		STDMETHOD(GetNextRows)(THIS_ LARGE_INTEGER udlRowsToSkip, DBFETCHROWS *pFetchParams);
		STDMETHOD(Requery)(THIS);
		STDMETHOD(Move)(THIS_ ULONG cbBookmark, void *pBookmark, LARGE_INTEGER dlOffset, DBFETCHROWS *pFetchParams);
		STDMETHOD(GetBookmark)(THIS_ DBCOLUMNID *pBookmarkType, ULONG cbMaxSize, ULONG *pcbBookmark, void *pBookmark);
		STDMETHOD(Clone)(THIS_ DWORD dwFlags, REFIID riid, IUnknown **ppvClonedCursor);
	END_INTERFACE_PART(CURSORMOVE)

	BEGIN_INTERFACE_PART(CursorUpdateARow, ICursorUpdateARow)
	     INIT_INTERFACE_PART(CRowCursor, CursorUpdateARow)
        STDMETHOD (BeginUpdate)(THIS_ DWORD dwFlags);
        STDMETHOD (SetColumn)(THIS_ DBCOLUMNID * pcid, DBBINDPARAMS * pBindParams);
        STDMETHOD (GetColumn)(THIS_ DBCOLUMNID * pcid, DBBINDPARAMS * pBindParams, DWORD * pdwFlags);
        STDMETHOD (GetEditMode)(THIS_ DWORD * pdwState);
        STDMETHOD (Update)(THIS_ DBCOLUMNID * pBookmarkType, ULONG * pcbBookmark, void ** ppBookmark);
        STDMETHOD (Cancel)(THIS);
        STDMETHOD (Delete)(THIS);
    END_INTERFACE_PART(CursorUpdateARow)

	// DBGRID needs to see ICursorScroll and ICursorFind

	BEGIN_INTERFACE_PART(CURSORSCROLL, ICursorScroll)
		INIT_INTERFACE_PART(CRowCursor, CURSORSCROLL)
		STDMETHOD(GetColumnsCursor)(THIS_ REFIID riid,IUnknown **ppvColumnsCursor,ULONG *pcRows);
		STDMETHOD(SetBindings)(THIS_ ULONG cCol, DBCOLUMNBINDING rgBoundColumns[],ULONG cbRowLength,DWORD dwFlags);
		STDMETHOD(GetBindings)(THIS_ ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength);
		STDMETHOD(GetNextRows)(THIS_ LARGE_INTEGER udlRowsToSkip, DBFETCHROWS *pFetchParams);
		STDMETHOD(Requery)(THIS);
		STDMETHOD(Move)(THIS_ ULONG cbBookmark, void *pBookmark, LARGE_INTEGER dlOffset, DBFETCHROWS *pFetchParams);
		STDMETHOD(GetBookmark)(THIS_ DBCOLUMNID *pBookmarkType, ULONG cbMaxSize, ULONG *pcbBookmark, void *pBookmark);
		STDMETHOD(Clone)(THIS_ DWORD dwFlags, REFIID riid, IUnknown **ppvClonedCursor);
		STDMETHOD (Scroll)(THIS_ ULONG, ULONG, DBFETCHROWS *);
		STDMETHOD (GetApproximatePosition)(THIS_ ULONG, void  *, ULONG *, ULONG *);
		STDMETHOD (GetApproximateCount)(THIS_ LARGE_INTEGER *, DWORD *);
	END_INTERFACE_PART(CURSORSCROLL)

	BEGIN_INTERFACE_PART(CURSORFIND, ICursorFind)
		INIT_INTERFACE_PART(CRowCursor, CURSORFIND)
		STDMETHOD (FindByValues)(THIS_ ULONG cbBookmark, LPVOID pBookmark, DWORD dwFindFlags, ULONG cValues, DBCOLUMNID rgColumns[], DBVARIANT rgValues[], DWORD rgdwSeekFlags[], DBFETCHROWS * pFetchParams);
	END_INTERFACE_PART(CURSORFIND)


	DECLARE_INTERFACE_MAP()


	BEGIN_CONNECTION_PART(CRowCursor, ConnectionPoint)
        CONNECTION_IID(IID_INotifyDBEvents)
   END_CONNECTION_PART(ConnectionPoint)


	DECLARE_CONNECTION_MAP()
	
	// Helper Functions
public:
	CString ConvertToString(DBVARIANT value);
	BOOL IsOpen();
	void Close();
	BOOL is_clone;
	CRowCursor* clone;
	CRowCursor* parent;
	CMyDaoRecordset* m_pRecordSet;
	
	HRESULT FetchRows(DBFETCHROWS *pFetchParams);
	
private:

	HRESULT CvtToString(DWORD dwRequestedType, BYTE *pDestColVal, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol, DBFETCHROWS *pFetchParams, ULONG *pcbCumulativeOutOfLineMemSize);
	HRESULT CvtToVariant(DWORD dwRequestedType, BYTE *pDestColVal, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol);
	HRESULT GetBlobData(BYTE *pDestBytePtr, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol, DBFETCHROWS *pFetchParams, ULONG *pcbCumulativeOutOfLineMemSize,char *add=NULL);
	HRESULT GetFixedLengthInLineMemTypes(DWORD dwRequestedType, BYTE *pDestBytePtr, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol);
	HRESULT DoNotifyBefore(DWORD dwEventWhat, ULONG cReasons, DBNOTIFYREASON rgReasons[]);
	HRESULT DoNotifyAfter(DWORD dwEventWhat, ULONG cReasons, DBNOTIFYREASON rgReasons[]);
	HRESULT CallFailedToDo(DWORD dwEventWhat, ULONG cReasons, DBNOTIFYREASON rgReasons[]);
	HRESULT MakeSafearrayFromBuf(SAFEARRAY FAR * FAR *ppsa, LPVOID p, ULONG cb);
	void MakeSafearrayFromBok(SAFEARRAY FAR * FAR *ppbBookmark);

private:

	CDscdaoCtrl *m_pControl;			// ptr to the Control Class object
	CColCursor *m_pColCursor;			// ptr to the Columns Cursor Object
	
	ULONG m_lRowDataLength;				// Size in bytes of each row of fetched data

	DBCOLUMNBINDING *m_pDataColumns;	// ptr to an array of Column Bindings for actual data
	ULONG m_nExistingDataCols;			// number of existing bound data columns

	DWORD m_dwEditMode;					// Indicates if Cursor is in Edit mode or AddNew mode
	BOOL m_bUpdateInProgress;			// Flag to prevent Re-entrancy of Updates, once begun
};

#endif
