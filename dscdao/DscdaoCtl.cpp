// DscdaoCtl.cpp : Implementation of the CDscdaoCtrl ActiveX Control class.

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


#include <initguid.h>



#include "dscdao.h"
#include "DscdaoCtl.h"
#include "DscdaoPpg.h"
#include "ColCrsr.h"
#include "RowCrsr.h"
#include "DbasePropPage.h"
#include "GenPropPage.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



#define GetInterfacePtr(pTarget, pEntry) \
	((LPUNKNOWN)((BYTE*)pTarget + pEntry->nOffset))


// UI Constants
#define DSC_ARROWHEIGHT 10
#define DSC_ARROWWIDTH 5
#define DSC_BUTTONWIDTH 18
#define DSC_BUTTONHEIGHT 19
#define DSC_STOPWIDTH 2

enum tagButtons
{
	DSC_MOVEFIRST,
	DSC_MOVENEXT,
	DSC_MOVEPREVIOUS,
	DSC_MOVELAST,
	DSC_ADDNEW,		
	DSC_CANCEL,
	DSC_UPDATE
};

enum tagState
{
	DSC_BUTTONDOWN,
	DSC_BUTTONUP,
	DSC_BUTTONDISABLED
};



IMPLEMENT_DYNCREATE(CDscdaoCtrl, COleControl)

BEGIN_INTERFACE_MAP(CDscdaoCtrl, COleControl)
	INTERFACE_PART(CDscdaoCtrl, IID_IVBDSC, VBDSC)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CDscdaoCtrl, COleControl)
	//{{AFX_MSG_MAP(CDscdaoCtrl)
	ON_WM_CREATE()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	//}}AFX_MSG_MAP
//	ON_WM_DESTROY()
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CDscdaoCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CDscdaoCtrl)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "ShowAddnewButton", m_showAddnewButton, OnShowAddnewButtonChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "ShowCancelButton", m_showCancelButton, OnShowCancelButtonChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "ShowUpdateButton", m_showUpdateButton, OnShowUpdateButtonChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "DatabasePath", m_databasePath, OnDatabasePathChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "TableName", m_Table, OnTableNameChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "IndexName", m_indexName, OnIndexNameChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "ReadOnly", m_readOnly, OnReadOnlyChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "OpenType", m_openType, OnOpenTypeChanged, VT_I2)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "SqlType", m_sqlType, OnSqlTypeChanged, VT_I2)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "SqlString", m_sqlString, OnSqlStringChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "DrawColor", m_drawColor, OnDrawColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "OpenExclusive", m_openExclusive, OnOpenExclusiveChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "Password", m_password, OnPasswordChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CDscdaoCtrl, "LockingMode", m_lockingMode, OnLockingModeChanged, VT_BOOL)
	DISP_FUNCTION(CDscdaoCtrl, "Next", Next, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "Previous", Previous, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "Beginning", Beginning, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "End", End, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "AddNew", AddNew, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "Delete", Delete, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "CancelUpdate", CancelUpdate, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "Update", Update, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "BeginTransaction", BeginTransaction, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "CommitTransaction", CommitTransaction, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CDscdaoCtrl, "Rollback", Rollback, VT_EMPTY, VTS_NONE)
	DISP_STOCKPROP_APPEARANCE()
	DISP_STOCKPROP_BACKCOLOR()
	DISP_STOCKPROP_FONT()
	DISP_STOCKPROP_CAPTION()
	DISP_STOCKPROP_FORECOLOR()
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CDscdaoCtrl, COleControl)
	//{{AFX_EVENT_MAP(CDscdaoCtrl)
	// NOTE - ClassWizard will add and remove event map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CDscdaoCtrl, 5)
	PROPPAGEID(CDscdaoPropPage::guid)
	PROPPAGEID(CDbasePropPage::guid)
	PROPPAGEID(CGenPropPage::guid)
   PROPPAGEID(CLSID_CFontPropPage)
   PROPPAGEID(CLSID_CColorPropPage)
END_PROPPAGEIDS(CDscdaoCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CDscdaoCtrl, "DSCDAO.DscdaoCtrl.1",
	0x64d3f3d6, 0x494d, 0x11d2, 0xbc, 0x7d, 0, 0, 0x21, 0x6a, 0x6, 0xc9)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CDscdaoCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DDscdao =
		{ 0x64d3f3d4, 0x494d, 0x11d2, { 0xbc, 0x7d, 0, 0, 0x21, 0x6a, 0x6, 0xc9 } };
const IID BASED_CODE IID_DDscdaoEvents =
		{ 0x64d3f3d5, 0x494d, 0x11d2, { 0xbc, 0x7d, 0, 0, 0x21, 0x6a, 0x6, 0xc9 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwDscdaoOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CDscdaoCtrl, IDS_DSCDAO, _dwDscdaoOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::CDscdaoCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CDscdaoCtrl

BOOL CDscdaoCtrl::CDscdaoCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_DSCDAO,
			IDB_DSCDAO,
			afxRegApartmentThreading,
			_dwDscdaoOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::CDscdaoCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

//BOOL CDscdaoCtrl::CDscdaoCtrlFactory::VerifyUserLicense()
//{
//	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
//		_szLicString);
//}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::CDscdaoCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

//BOOL CDscdaoCtrl::CDscdaoCtrlFactory::GetLicenseKey(DWORD dwReserved,
//	BSTR FAR* pbstrKey)
//{
//	if (pbstrKey == NULL)
//		return FALSE;

//	*pbstrKey = SysAllocString(_szLicString);
//	return (*pbstrKey != NULL);
//}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::CDscdaoCtrl - Constructor


CDscdaoCtrl::CDscdaoCtrl()
{
   TRACE("\nControl class constructor\n");
   reset=FALSE;
   m_pdatabase=new CDaoDatabase;
	m_pRowCursor = NULL;


	// TODO: Initialize your control's instance data here.
	
	m_pColOffsetInRow = NULL;
	m_nColumns = 0;
	m_nBookmarks = 0;
	m_cbRowSize = 0;
	m_pMetaData = NULL;	
  	m_sAppearance = 0; //no 3D
	m_showCancelButton=FALSE;
	m_showUpdateButton=FALSE;
	m_showAddnewButton=FALSE;
	m_databasePath=_T("");
	m_Table=_T("");
	m_indexName=_T("");
	m_readOnly=FALSE;
	m_openType=0;
	m_sqlType=0;

   InitializeIIDs(&IID_DDscdao, &IID_DDscdaoEvents);
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::~CDscdaoCtrl - Destructor

CDscdaoCtrl::~CDscdaoCtrl()
{
	// TODO: Cleanup your control's instance data here.
   TRACE("\nControl class destructor\n");
   if(m_pRowCursor){
         m_pRowCursor->ExternalRelease();
   }
   m_pdatabase->Close();
   delete m_pdatabase;
   if (m_pColOffsetInRow)CoTaskMemFree((LPVOID)m_pColOffsetInRow);
   if (m_pMetaData){
	   for(unsigned long int i=0;i<m_nColumns + m_nBookmarks;i++){
		   if(m_pMetaData[i].szColName){
			   ::SysFreeString((BSTR)m_pMetaData[i].szColName);		
		   }
      }
	   CoTaskMemFree ((LPVOID)m_pMetaData);
   }
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::OnDraw - Drawing function

void CDscdaoCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /*rcInvalid*/)
{
	// TODO: Replace the following code with your own drawing code.
	DrawButtons(pdc, rcBounds);
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::DoPropExchange - Persistence support

void CDscdaoCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
	PX_Bool(pPX,_T("ShowAddnewButton"),m_showAddnewButton);
	PX_Bool(pPX,_T("ShowCancelButton"),m_showCancelButton);
	PX_Bool(pPX,_T("ShowUpdateButton"),m_showUpdateButton);
	PX_String(pPX,_T("TableName"),m_Table,_T(""));
	PX_String(pPX,_T("IndexName"),m_indexName,_T("None"));
	PX_String(pPX,_T("DatabasePath"),m_databasePath,_T(""));
	PX_Bool(pPX,_T("ReadOnly"),m_readOnly,FALSE);
	PX_Short(pPX,_T("OpenType"),m_openType,0);
	PX_Short(pPX,_T("SqlType"),m_sqlType,0);
	PX_String(pPX,_T("SqlString"),m_sqlString,_T(""));
	PX_Color(pPX,_T("DrawColor"),m_drawColor);
	PX_String(pPX,_T("Password"),m_password,_T(""));
	PX_Bool(pPX,_T("OpenExclusive"),m_openExclusive,FALSE);
	PX_Bool(pPX,_T("LockingMode"),m_lockingMode,FALSE);
   
   if(pPX->IsLoading() && !reset && !m_pRowCursor){
      m_pRowCursor = new CRowCursor(this);
//      InterlockedDecrement(&m_pRowCursor->m_dwRef);
   }
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::OnResetState - Reset control to default state

void CDscdaoCtrl::OnResetState()
{
	reset=TRUE;
   COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
	m_sAppearance = 0; //no 3D
	m_showCancelButton=FALSE;
	m_showUpdateButton=FALSE;
	m_showAddnewButton=FALSE;
   m_lockingMode=FALSE;
	m_databasePath=_T("");
	m_password=m_sqlString=m_Table=_T("");
   m_openExclusive=FALSE;
	m_indexName=_T("");
	m_readOnly=FALSE;
	m_openType=0;
	m_sqlType=0;
	SetText(AmbientDisplayName());
	SetBackColor(RGB(255,255,255));
   m_drawColor=RGB(0,0,0);
   reset=FALSE;
}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl::AboutBox - Display an "About" box to the user

//void CDscdaoCtrl::AboutBox()
//{
//	CDialog dlgAbout(IDD_ABOUTBOX_DSCDAO);
//	dlgAbout.DoModal();
//}


/////////////////////////////////////////////////////////////////////////////
// CDscdaoCtrl message handlers

void CDscdaoCtrl::Next() 
{
	// TODO: Add your dispatch handler code here
	LARGE_INTEGER dlOffset;
	LISet32 (dlOffset, (ULONG)1L);
   if(m_pRowCursor && m_pRowCursor->IsOpen()){
	if(!m_pRowCursor->m_pRecordSet->IsEOF())
		m_pRowCursor->m_xCURSORMOVE.Move(DB_BMK_SIZE, (LPVOID)&DBBMK_CURRENT, dlOffset, NULL);
   }
}

void CDscdaoCtrl::Previous() 
{
	// TODO: Add your dispatch handler code here
	LARGE_INTEGER dlOffset;
	LISet32 (dlOffset, (ULONG)-1L);
   if(m_pRowCursor && m_pRowCursor->IsOpen()){
      if(!m_pRowCursor->m_pRecordSet->IsBOF())
	      m_pRowCursor->m_xCURSORMOVE.Move(DB_BMK_SIZE, (LPVOID)&DBBMK_CURRENT, dlOffset, NULL);
   }
}

void CDscdaoCtrl::Beginning() 
{
	// TODO: Add your dispatch handler code here
	LARGE_INTEGER dlOffset;
	LISet32 (dlOffset, (ULONG)1L);
	if(m_pRowCursor && m_pRowCursor->IsOpen())m_pRowCursor->m_xCURSORMOVE.Move(DB_BMK_SIZE, (LPVOID)&DBBMK_BEGINNING, dlOffset, NULL);
}

void CDscdaoCtrl::End() 
{
	// TODO: Add your dispatch handler code here
	LARGE_INTEGER dlOffset;
	LISet32 (dlOffset, (ULONG)-1L);	
   if(m_pRowCursor && m_pRowCursor->IsOpen())m_pRowCursor->m_xCURSORMOVE.Move(DB_BMK_SIZE, (LPVOID)&DBBMK_END, dlOffset, NULL);
}

void CDscdaoCtrl::AddNew() 
{
	// TODO: Add your dispatch handler code here
	DWORD dwState;
	TRACE("Addnew");
   if(!m_pRowCursor || !m_pRowCursor->IsOpen())return;
	m_pRowCursor->m_xCursorUpdateARow.GetEditMode(&dwState);
	if (dwState == DBEDITMODE_NONE)
		m_pRowCursor->m_xCursorUpdateARow.BeginUpdate(DBROWACTION_ADD);
}

void CDscdaoCtrl::Delete() 
{
	// TODO: Add your dispatch handler code here
	DWORD dwState;
	TRACE("Delete");
   if(!m_pRowCursor || !m_pRowCursor->IsOpen())return;
	m_pRowCursor->m_xCursorUpdateARow.GetEditMode(&dwState);
	if (dwState == DBEDITMODE_NONE)
		m_pRowCursor->m_xCursorUpdateARow.Delete();
}

void CDscdaoCtrl::CancelUpdate() 
{
	// TODO: Add your dispatch handler code here
	DWORD dwState;
	TRACE("CancelUpdate");
   if(!m_pRowCursor || !m_pRowCursor->IsOpen())return;
	m_pRowCursor->m_xCursorUpdateARow.GetEditMode(&dwState);
	if (dwState != DBEDITMODE_NONE)
		m_pRowCursor->m_xCursorUpdateARow.Cancel();
}

void CDscdaoCtrl::Update() 
{
	// TODO: Add your dispatch handler code here
	DWORD dwState;
	TRACE("Update");
   if(!m_pRowCursor || !m_pRowCursor->IsOpen())return;
	m_pRowCursor->m_xCursorUpdateARow.GetEditMode(&dwState);
	if (dwState != DBEDITMODE_NONE)
		m_pRowCursor->m_xCursorUpdateARow.Update(NULL, NULL, NULL);
}

void CDscdaoCtrl::OnShowAddnewButtonChanged() 
{
	// TODO: Add notification handler code
	int cx, cy;
 
	GetControlSize(&cx, &cy);
	SetControlSize(cx, cy);
	
	SetModifiedFlag();	
	InvalidateControl();
}

void CDscdaoCtrl::OnShowCancelButtonChanged() 
{
	// TODO: Add notification handler code
	int cx, cy;
 
	GetControlSize(&cx, &cy);
	SetControlSize(cx, cy);
	
	SetModifiedFlag();	
	InvalidateControl();
}

void CDscdaoCtrl::OnShowUpdateButtonChanged() 
{
	// TODO: Add notification handler code
	
	int cx, cy;
 
	GetControlSize(&cx, &cy);
	SetControlSize(cx, cy);
	
	SetModifiedFlag();	
	InvalidateControl();
}



LPUNKNOWN CDscdaoCtrl::GetInterfaceHook(const void* piid)
{
    void* pObj;

    pObj=COleControl::GetInterfaceHook(piid);
 	 
    if(!pObj){
      if (IsEqualIID(*(IID*)piid, IID_ICursor))
		   m_pRowCursor->m_xCURSOR.QueryInterface(IID_ICursor,(void **)&pObj);
	   else if (IsEqualIID(*(IID*)piid, IID_ICursorMove))
		   m_pRowCursor->m_xCURSORMOVE.QueryInterface(IID_ICursorMove,(void **)&pObj);
      else if (IsEqualIID(*(IID*)piid, IID_ICursorFind))
		   m_pRowCursor->m_xCURSORFIND.QueryInterface(IID_ICursorFind,(void **)&pObj);
      else if (IsEqualIID(*(IID*)piid, IID_ICursorScroll))
		   m_pRowCursor->m_xCURSORSCROLL.QueryInterface(IID_ICursorScroll,(void **)&pObj);
      else if (!m_readOnly && IsEqualIID(*(IID*)piid, IID_ICursorUpdateARow))
		   m_pRowCursor->m_xCursorUpdateARow.QueryInterface(IID_ICursorUpdateARow,(void **)&pObj);
    }
    
    return (LPUNKNOWN)pObj;
}

HRESULT CDscdaoCtrl::BindColumns(DBCOLUMNBINDING FAR **pBoundColumns, ULONG *nExistingBoundCols, ULONG cCol, DBCOLUMNBINDING rgBoundColumns[], DWORD dwFlags)
{
	if (!cCol){
		if (dwFlags == DBCOLUMNBINDOPTS_REPLACE){
			// remove existing column bindings
         if (*pBoundColumns){
            CoTaskMemFree ((LPVOID)(*pBoundColumns));
            *pBoundColumns=NULL;
         }
			return S_OK;
		}else if (dwFlags == DBCOLUMNBINDOPTS_ADD)
			// do nothing -- no columns to add!
			return S_OK;
		else return E_INVALIDARG;
	}else if (!rgBoundColumns && dwFlags == DBCOLUMNBINDOPTS_REPLACE)return E_INVALIDARG;
	else if (dwFlags != DBCOLUMNBINDOPTS_REPLACE && dwFlags != DBCOLUMNBINDOPTS_REPLACE)return E_INVALIDARG;
	
	if (dwFlags == DBCOLUMNBINDOPTS_REPLACE){
		// remove existing bound columns
		if (*pBoundColumns)CoTaskMemFree ((LPVOID)(*pBoundColumns));
		*pBoundColumns = (DBCOLUMNBINDING *)CoTaskMemAlloc(sizeof(DBCOLUMNBINDING) * cCol);
		if (!(*pBoundColumns)){
			*nExistingBoundCols = 0;
			return E_OUTOFMEMORY;
		}
		
		*nExistingBoundCols = cCol;
		for (ULONG i = 0; i < cCol; i++){
			// copy column binding info into allocated memory.
			(*pBoundColumns)[i] = rgBoundColumns[i];

			// for COLUMNIDs with name fields, the string member, lpdbsz will have to be copied 
			// over as well also!
			if (rgBoundColumns[i].columnID.dwKind == DBCOLKIND_GUID_NAME || rgBoundColumns[i].columnID.dwKind == DBCOLKIND_NAME){
				(*pBoundColumns)[i].columnID.lpdbsz = (LPDBSTR)CoTaskMemAlloc(sizeof(DBCHAR) * (ldbstrlen(rgBoundColumns[i].columnID.lpdbsz) + 1));
				if (!((*pBoundColumns)[i].columnID.lpdbsz)) return E_OUTOFMEMORY;
				ldbstrcpy((*pBoundColumns)[i].columnID.lpdbsz, rgBoundColumns[i].columnID.lpdbsz);
			}
		}

    }else if (dwFlags == DBCOLUMNBINDOPTS_ADD){
		// save existing bound columns
		DBCOLUMNBINDING *pExistingBoundColumns = *pBoundColumns;
 		
		*pBoundColumns = (DBCOLUMNBINDING *) CoTaskMemAlloc(sizeof(DBCOLUMNBINDING) * (cCol + *nExistingBoundCols));
		if (!(*pBoundColumns)){
			// restore existing bound columns, in case of failure
			*pBoundColumns = pExistingBoundColumns;
			return E_OUTOFMEMORY;
		}

		// copy the existing bound columns into the newly allocated memory, before adding new columns
		for (ULONG i = 0; i < *nExistingBoundCols; i++)(*pBoundColumns)[i] = pExistingBoundColumns[i];

		// delete the old copy of bound columns as it is no longer needed.
		CoTaskMemFree ((LPVOID)(pExistingBoundColumns));
		
		for (ULONG j = i; j < i+cCol; j++){
			// copy the new column binding info into allocated memory.
			(*pBoundColumns)[j] = rgBoundColumns[j-i];

			// for COLUMNIDs with name fields, the string member, lpdbsz will have to be copied 
			// over as well also!
			if (rgBoundColumns[j-i].columnID.dwKind == DBCOLKIND_GUID_NAME || rgBoundColumns[j-i].columnID.dwKind == DBCOLKIND_NAME){
				(*pBoundColumns)[j].columnID.lpdbsz = (LPDBSTR)CoTaskMemAlloc(sizeof(DBCHAR) * (ldbstrlen(rgBoundColumns[j-i].columnID.lpdbsz) + 1));
				if (!((*pBoundColumns)[j].columnID.lpdbsz)) return E_OUTOFMEMORY;
				ldbstrcpy((*pBoundColumns)[j].columnID.lpdbsz, rgBoundColumns[j-i].columnID.lpdbsz);
			}
		}
		*nExistingBoundCols += cCol;
	}
	return S_OK;
}


// Draws all buttons on control including the extra buttons if their Show 
// properties are selected
void CDscdaoCtrl::DrawButtons(CDC * pdc,const CRect& rcBounds)
{
	CFont* pOldFont;
   const CString& strName = InternalGetText();
	
	CBrush hbBackColor(TranslateColor(GetBackColor()));
	CBrush hb(GetSysColor(COLOR_BTNTEXT));
	CPen hp(PS_SOLID, 1, GetSysColor(COLOR_BTNTEXT));
	CPen hpFace(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
	CPen *hpOld, *hpFaceOld;

	CRect rectText;

	if (m_sAppearance){// 3D
		m_nBottomBorderWidth = 3;
		m_nTopBorderWidth = 2;
	}else{
		m_nBottomBorderWidth = 1;
		m_nTopBorderWidth = 1;
	}

	// How many extra buttons are we going to have
	int nExtraButtons=GetExtraButtonCount();
	
	// For the Client area of the control for printing text
	rectText = rcBounds;
	rectText.left += (DSC_BUTTONWIDTH * 2 + m_nTopBorderWidth + 2);
	rectText.right -= (DSC_BUTTONWIDTH * (2 + nExtraButtons) + m_nBottomBorderWidth + 2 + nExtraButtons);
	rectText.top += m_nTopBorderWidth;
	rectText.bottom -= m_nBottomBorderWidth;
	
	// Draw the border around the outside of all the buttons
	DrawBorder(pdc, rcBounds);
	
	hpOld =(CPen *) pdc->SelectObject(&hp);		
	
	if (m_sAppearance)
		rectText.bottom += 1;

	// Fills area behind the control name.
	pdc->FillRect(rectText, &hbBackColor);

	pOldFont = SelectStockFont(pdc);
	pdc->SetBkMode(TRANSPARENT);
	pdc->SetTextColor(TranslateColor(GetForeColor()));

	// Moves the text to the right to look like VB's data control
	rectText.left +=2;

	// Draw the control name on the small client area
	pdc->DrawText(strName, strName.GetLength(), &rectText, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
		
	if(pOldFont)pdc->SelectObject(pOldFont);

	// Draw each individual button. 
	DrawButtonAt(pdc, rcBounds,DSC_MOVEFIRST, DSC_BUTTONUP);
	DrawButtonAt(pdc, rcBounds,DSC_MOVEPREVIOUS,DSC_BUTTONUP);
	DrawButtonAt(pdc, rcBounds,DSC_MOVENEXT,DSC_BUTTONUP);
	DrawButtonAt(pdc, rcBounds,DSC_MOVELAST,DSC_BUTTONUP);

	// Draw each extra button if Show property is set
	if(m_showAddnewButton)DrawButtonAt(pdc, rcBounds,DSC_ADDNEW, DSC_BUTTONUP);
	if(m_showCancelButton)DrawButtonAt(pdc, rcBounds,DSC_CANCEL, DSC_BUTTONUP);
	if(m_showUpdateButton)DrawButtonAt(pdc, rcBounds,DSC_UPDATE, DSC_BUTTONUP);

	// Draw lines in between buttons

	// First
	pdc->MoveTo(rcBounds.left + DSC_BUTTONWIDTH + m_nTopBorderWidth, rcBounds.top + m_nTopBorderWidth - 1);
	pdc->LineTo(rcBounds.left + DSC_BUTTONWIDTH + m_nTopBorderWidth, rcBounds.bottom - m_nBottomBorderWidth + 1);

	// Second
	pdc->MoveTo(rcBounds.left + DSC_BUTTONWIDTH*2 + m_nTopBorderWidth*2 - (m_nTopBorderWidth - 1), rcBounds.top + m_nTopBorderWidth - 1);
	pdc->LineTo(rcBounds.left + DSC_BUTTONWIDTH*2 + m_nTopBorderWidth*2 - (m_nTopBorderWidth - 1), rcBounds.bottom - m_nBottomBorderWidth + 1);
	
	// Third
	if (m_sAppearance){	

		hpFaceOld = (CPen *)pdc->SelectObject(&hpFace);
		pdc->MoveTo(rcBounds.right - DSC_BUTTONWIDTH*(2 + nExtraButtons) - m_nBottomBorderWidth * 2 - 1 + m_nBottomBorderWidth -1 - nExtraButtons, rcBounds.top + m_nTopBorderWidth);
		pdc->LineTo(rcBounds.right - DSC_BUTTONWIDTH*(2 + nExtraButtons) - m_nBottomBorderWidth * 2 - 1 + m_nBottomBorderWidth -1 - nExtraButtons, rcBounds.bottom - m_nBottomBorderWidth + 1);	
		pdc->SelectObject(hpFaceOld);
	}else{
		pdc->MoveTo(rcBounds.right - DSC_BUTTONWIDTH*(2 + nExtraButtons) - m_nBottomBorderWidth*2 - 1+ m_nBottomBorderWidth -1 - nExtraButtons, rcBounds.top + m_nTopBorderWidth - 1);
		pdc->LineTo(rcBounds.right - DSC_BUTTONWIDTH*(2 + nExtraButtons) - m_nBottomBorderWidth*2 - 1+ m_nBottomBorderWidth -1 - nExtraButtons, rcBounds.bottom - m_nBottomBorderWidth + 1);	
	}

	pdc->MoveTo(rcBounds.right - DSC_BUTTONWIDTH - m_nBottomBorderWidth - 1, rcBounds.top + m_nTopBorderWidth - 1);
	pdc->LineTo(rcBounds.right - DSC_BUTTONWIDTH - m_nBottomBorderWidth - 1, rcBounds.bottom - m_nBottomBorderWidth + 1);

	// Draw lines for Extra Buttons
	if(nExtraButtons){
		for (int i=1;i<=nExtraButtons;i++){
			pdc->MoveTo(rcBounds.right - DSC_BUTTONWIDTH*(i+1) - m_nBottomBorderWidth - 1 - i, rcBounds.top + m_nTopBorderWidth - 1);
			pdc->LineTo(rcBounds.right - DSC_BUTTONWIDTH*(i+1) - m_nBottomBorderWidth - 1 - i, rcBounds.bottom - m_nBottomBorderWidth + 1);
		}

	}

	if(hpOld)pdc->SelectObject(hpOld);
}


// Draw each individual button.
void CDscdaoCtrl::DrawButtonAt(CDC * pdc,const CRect& rcBounds, UINT btn,UINT State)
{
	
	POINT ptOffset;
	POINT pt[3];

	
	int nInnerButtonSpace;
	int nOuterButtonSpace;
	int DownOffset;
	
	nOuterButtonSpace =  (DSC_BUTTONWIDTH - (DSC_ARROWWIDTH + DSC_STOPWIDTH + 2))/2; 
	nInnerButtonSpace = (DSC_BUTTONWIDTH - DSC_ARROWWIDTH) /2;

	CRect rectStop;
	CRect rcShadow;

	CBrush hb(/*GetSysColor(COLOR_BTNTEXT)*/TranslateColor(m_drawColor));
	CBrush *hbOld;
	CPen hp(PS_SOLID, 1, /*GetSysColor(COLOR_BTNTEXT)*/TranslateColor(m_drawColor));
	CPen hpWidePen(PS_SOLID, 2, /*GetSysColor(COLOR_BTNTEXT)*/TranslateColor(m_drawColor));
	CPen *hpOld, *hpWidePenOld;

	if(State == DSC_BUTTONDOWN)DownOffset = 1;
	else DownOffset = 0;
	
	hbOld =(CBrush *) pdc->SelectObject(&hb);
	hpOld =(CPen *) pdc->SelectObject(&hp);
	ptOffset.y = (rcBounds.Height()- m_nTopBorderWidth - m_nBottomBorderWidth - DSC_ARROWHEIGHT) / 2;
	
	// Draw Button
	switch (btn){
		case DSC_MOVEFIRST:
			ptOffset.x = m_nTopBorderWidth;
			
			m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
				rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);

			DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);
			
			// Upper Right
			pt[0].x = m_Buttons[btn].left + DSC_ARROWWIDTH + DSC_STOPWIDTH + 2 + nOuterButtonSpace + DownOffset;
			pt[0].y = m_Buttons[btn].top + ptOffset.y + DownOffset;
			// Left
			pt[1].x	= m_Buttons[btn].left + DSC_STOPWIDTH + 2 + nOuterButtonSpace + DownOffset;
			pt[1].y = m_Buttons[btn].top + DSC_ARROWHEIGHT/2 + ptOffset.y + DownOffset;
			// Lower Right
			pt[2].x = m_Buttons[btn].left + DSC_ARROWWIDTH + DSC_STOPWIDTH + 2 + nOuterButtonSpace + DownOffset;
			pt[2].y = m_Buttons[btn].top + DSC_ARROWHEIGHT + ptOffset.y + DownOffset;

			// Draw Stop
			rectStop.SetRect(m_Buttons[btn].left + nOuterButtonSpace + 1,
				m_Buttons[btn].top + ptOffset.y + DownOffset,
				m_Buttons[btn].left + DSC_STOPWIDTH + nOuterButtonSpace + 1,
				m_Buttons[btn].top + ptOffset.y + DSC_ARROWHEIGHT + DownOffset + 1);

			rectStop.OffsetRect(DownOffset, DownOffset);
			pdc->Rectangle(rectStop);
			pdc->Polygon(pt, 3);
			break;
		case DSC_MOVEPREVIOUS:

			ptOffset.x = DSC_BUTTONWIDTH + 2 + m_nTopBorderWidth - 1;
			
			m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
				rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);

			DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);

			// Upper Right
			pt[0].x = m_Buttons[btn].left + DSC_ARROWWIDTH + nInnerButtonSpace + DownOffset;
			pt[0].y = m_Buttons[btn].top + ptOffset.y + DownOffset;
			// Left
			pt[1].x	= m_Buttons[btn].left + nInnerButtonSpace + DownOffset;
			pt[1].y = m_Buttons[btn].top + DSC_ARROWHEIGHT/2 + ptOffset.y + DownOffset;
			// Lower Right
			pt[2].x = m_Buttons[btn].left + DSC_ARROWWIDTH + nInnerButtonSpace + DownOffset;
			pt[2].y = m_Buttons[btn].top + DSC_ARROWHEIGHT + ptOffset.y + DownOffset;
			pdc->Polygon(pt, 3);
			break;
		case DSC_MOVENEXT:
			ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 2 - m_nBottomBorderWidth -1; 
						
			m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
				rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);

			DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);

			// Upper Left
			pt[0].x = m_Buttons[btn].left + nInnerButtonSpace + DownOffset;
			pt[0].y = m_Buttons[btn].top + ptOffset.y + DownOffset;
			// Right
			pt[1].x	= m_Buttons[btn].left + DSC_ARROWWIDTH + nInnerButtonSpace + DownOffset;
			pt[1].y = m_Buttons[btn].top + DSC_ARROWHEIGHT/2 + ptOffset.y + DownOffset;
			// Lower Left
			pt[2].x = m_Buttons[btn].left + nInnerButtonSpace + DownOffset;
			pt[2].y = m_Buttons[btn].top + DSC_ARROWHEIGHT + ptOffset.y + DownOffset;
			pdc->Polygon(pt, 3);
			break;
		case DSC_MOVELAST:
			ptOffset.x = rcBounds.Width() - DSC_BUTTONWIDTH - m_nBottomBorderWidth;
			
			m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
				rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);

			DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);

			// Upper Left
			pt[0].x = m_Buttons[btn].left + nOuterButtonSpace + DownOffset + 1;
			pt[0].y = m_Buttons[btn].top + ptOffset.y + DownOffset;
			// Right
			pt[1].x	= m_Buttons[btn].left + DSC_ARROWWIDTH + nOuterButtonSpace + DownOffset + 1;
			pt[1].y = m_Buttons[btn].top + DSC_ARROWHEIGHT/2 + ptOffset.y + DownOffset;
			// Lower Left
			pt[2].x = m_Buttons[btn].left + nOuterButtonSpace + DownOffset + 1;
			pt[2].y = m_Buttons[btn].top + DSC_ARROWHEIGHT + ptOffset.y + DownOffset;

			// DrawStop
			rectStop.SetRect(m_Buttons[btn].left + DSC_ARROWWIDTH + nOuterButtonSpace + 3,
				m_Buttons[btn].top + ptOffset.y,
				m_Buttons[btn].left + DSC_ARROWWIDTH + DSC_STOPWIDTH + nOuterButtonSpace + 3,
				m_Buttons[btn].top + DSC_ARROWHEIGHT + ptOffset.y + 1);

			rectStop.OffsetRect(DownOffset, DownOffset);
			
			pdc->Rectangle(rectStop);
			pdc->Polygon(pt, 3);
			break;
		case DSC_UPDATE:
		{
			if(m_showUpdateButton){
				// This sets up the first position for extra buttons
				ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 3 - m_nBottomBorderWidth - 2; 
				m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
						rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);
				DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);

				// Beginning of Update Button Code
				hpWidePenOld = (CPen *) pdc->SelectObject(&hpWidePen);
				int RightOffset = 6;
				int LongSide = 9;
				int ShortSide = 3;
				pdc->MoveTo(m_Buttons[btn].left + RightOffset - 4 + DownOffset, m_Buttons[btn].top + DownOffset + ptOffset.y + 6);
				pdc->LineTo(m_Buttons[btn].left + RightOffset + DownOffset, m_Buttons[btn].top + DownOffset + ptOffset.y + LongSide+1);
				pdc->MoveTo(m_Buttons[btn].left + RightOffset + DownOffset + LongSide, m_Buttons[btn].top + DownOffset + ptOffset.y);
				pdc->LineTo(m_Buttons[btn].left + RightOffset + DownOffset, m_Buttons[btn].top + DownOffset + ptOffset.y + LongSide);
				pdc->SelectObject(hpWidePenOld);
				// End of Update Button Code
			}
			break;
		}
		case DSC_CANCEL:
		{
			if (m_showCancelButton){
				if (m_showUpdateButton){
					// Second Position
					ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 4 - m_nBottomBorderWidth - 3; 
						m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
					rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);
				}else{
					// Back at First Position
					ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 3 - m_nBottomBorderWidth - 2; 
					m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
						rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);
				}
				DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);
				
				// Beginning of Cancel Code
				hpWidePenOld = (CPen *) pdc->SelectObject(&hpWidePen);
							
				int RightOffset = 4;
				int CrossSize = 9;

				pdc->MoveTo(m_Buttons[btn].left + RightOffset + DownOffset, m_Buttons[btn].top + DownOffset + ptOffset.y);
				pdc->LineTo(m_Buttons[btn].left + RightOffset + DownOffset + CrossSize , m_Buttons[btn].top + DownOffset + ptOffset.y + CrossSize);

				pdc->MoveTo(m_Buttons[btn].left + RightOffset + DownOffset + CrossSize, m_Buttons[btn].top + DownOffset + ptOffset.y);
				pdc->LineTo(m_Buttons[btn].left + RightOffset + DownOffset, m_Buttons[btn].top + DownOffset + ptOffset.y + CrossSize);

				pdc->SelectObject(hpWidePenOld);
				// End of Cancel Code
			}
			break;
		}

		case DSC_ADDNEW:
		{
			if(m_showAddnewButton){
				if ((m_showUpdateButton && !m_showCancelButton) || (!m_showUpdateButton && m_showCancelButton)){
					// Second Position
					ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 4 - m_nBottomBorderWidth - 3; 
						m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
					rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);
				}else if(!m_showUpdateButton && !m_showCancelButton){
					// Back at First Position
					ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 3 - m_nBottomBorderWidth - 2; 
					m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
						rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);
				}else{
					// Third Position
					ptOffset.x=rcBounds.Width() - DSC_BUTTONWIDTH * 5 - m_nBottomBorderWidth - 4; 
					m_Buttons[btn].SetRect(rcBounds.left + ptOffset.x, rcBounds.top + m_nTopBorderWidth,
						rcBounds.left + DSC_BUTTONWIDTH + ptOffset.x, rcBounds.bottom - m_nBottomBorderWidth);
				}
				
				DrawButtonShadow(pdc, rcBounds, m_Buttons[btn], State);
			
				// Beginning of Addnew Code
				int RightOffset=3;

				// Left
				pt[0].x = m_Buttons[btn].left + RightOffset + DownOffset;
				pt[0].y = m_Buttons[btn].top + ptOffset.y + DownOffset + (DSC_ARROWHEIGHT - DSC_ARROWWIDTH)/2 + DSC_ARROWWIDTH;
				// Center
				pt[1].x	= m_Buttons[btn].left + RightOffset + (DSC_ARROWHEIGHT/2) + DownOffset;
				pt[1].y = m_Buttons[btn].top + ptOffset.y + DownOffset + (DSC_ARROWHEIGHT - DSC_ARROWWIDTH)/2;
				// Right
				pt[2].x = m_Buttons[btn].left + RightOffset + DownOffset + DSC_ARROWHEIGHT;
				pt[2].y = m_Buttons[btn].top + ptOffset.y + DownOffset + (DSC_ARROWHEIGHT - DSC_ARROWWIDTH)/2 + DSC_ARROWWIDTH;
				pdc->Polygon(pt, 3);

				// End of Addnew Code
			}
			break;
		}

	}
	
	if (hpOld)pdc->SelectObject(hpOld);
	if (hbOld)pdc->SelectObject(hbOld);
}

// Draw border around all controls
void CDscdaoCtrl::DrawBorder(CDC * pdc,const CRect& rcBounds)
{
	CRect ClientRect;
	
	CPen hpHighlight, hpShadow, hpFace;
	
	CPen *hpHighlightOld, *hpShadowOld, *hpFaceOld, *hpBlackOld;

	ClientRect = rcBounds;

	ClientRect.top += m_nTopBorderWidth - 1;
	ClientRect.left += m_nTopBorderWidth - 1; 
	
	ClientRect.bottom -= m_nBottomBorderWidth - 1;
	ClientRect.right -= m_nBottomBorderWidth - 1;

	hpBlackOld = pdc->SelectObject(CPen::FromHandle((HPEN)GetStockObject(BLACK_PEN)));

	pdc->MoveTo(ClientRect.TopLeft());
	pdc->LineTo(ClientRect.right, ClientRect.top);	
	pdc->MoveTo(ClientRect.TopLeft());	
	pdc->LineTo(ClientRect.left, ClientRect.bottom);
	
	pdc->MoveTo(ClientRect.left, ClientRect.bottom-1);
	pdc->LineTo(ClientRect.right-1, ClientRect.bottom-1);	
	pdc->LineTo(ClientRect.right-1, ClientRect.top);	

	pdc->SelectObject(hpBlackOld);

	if (m_sAppearance)
	{
		hpShadow.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW));
		hpHighlight.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT));
		hpFace.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
		
		hpShadowOld = (CPen *)pdc->SelectObject(&hpShadow);
		
		// Draw Top in Shadow Color
		pdc->MoveTo(rcBounds.TopLeft());
		pdc->LineTo(rcBounds.right, rcBounds.top);	
		pdc->MoveTo(rcBounds.TopLeft());	
		pdc->LineTo(rcBounds.left, rcBounds.bottom);
		pdc->SelectObject(hpShadowOld);
	
		// Draw Bottom
		hpHighlightOld = (CPen *)pdc->SelectObject(&hpHighlight);
		pdc->MoveTo(rcBounds.left, rcBounds.bottom-1);
		pdc->LineTo(rcBounds.right-1, rcBounds.bottom-1);	
		pdc->LineTo(rcBounds.right-1, rcBounds.top - 1);	
		pdc->SelectObject(hpHighlightOld);

		// Draw Extra
		hpFaceOld = (CPen *)pdc->SelectObject(&hpFace);
		pdc->MoveTo(rcBounds.left + 1, rcBounds.bottom - 2);
		pdc->LineTo(rcBounds.right - 2 , rcBounds.bottom - 2);	
		pdc->LineTo(rcBounds.right-2, rcBounds.top);	
		
		pdc->SelectObject(hpFaceOld);
			
	}
}


// Draw shadow around each individual control
void CDscdaoCtrl::DrawButtonShadow(CDC * pdc, const CRect& /*rcBounds*/, CRect rcShadow, int State)
{
	CPen hpHighlight;
	CPen hpShadow;
	CPen hpInsideShadow;
	CBrush hbFace(GetSysColor(COLOR_BTNFACE));
	
	CPen *hpHighlightOld, *hpShadowOld, *hpInsideShadowOld;
	CBrush *hbFaceOld;

	if (State == DSC_BUTTONDOWN)
	{
		hpHighlightOld = pdc->SelectObject(CPen::FromHandle((HPEN)GetStockObject(BLACK_PEN)));
		hpShadow.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
	}
	else
	{
		hpShadow.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW));
		hpHighlight.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT));
		hpHighlightOld = (CPen *)pdc->SelectObject(&hpHighlight);
	}

	hbFaceOld = (CBrush *) pdc->SelectObject(&hbFace);
	
	pdc->Rectangle(rcShadow);
	pdc->SelectObject(hbFaceOld);
	
	pdc->MoveTo(rcShadow.TopLeft());
	pdc->LineTo(rcShadow.right,rcShadow.top);
	pdc->MoveTo(rcShadow.TopLeft());
	pdc->LineTo(rcShadow.left,rcShadow.bottom);

	pdc->SelectObject(hpHighlightOld);
	
	hpShadowOld = (CPen *) pdc->SelectObject(&hpShadow);
	
	pdc->MoveTo(rcShadow.left + 1, rcShadow.bottom - 1);
	pdc->LineTo(rcShadow.right, rcShadow.bottom - 1);
	pdc->MoveTo(rcShadow.right - 1, rcShadow.top + 1);
	pdc->LineTo(rcShadow.right - 1, rcShadow.bottom);

	pdc->SelectObject(hpShadowOld);
		
	if (State == DSC_BUTTONDOWN)
	{
		hpInsideShadow.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW));
		hpInsideShadowOld = (CPen *)pdc->SelectObject(&hpInsideShadow);
		
		pdc->MoveTo(rcShadow.left + 1, rcShadow.top+1);
		pdc->LineTo(rcShadow.right-1,rcShadow.top+1);
		pdc->MoveTo(rcShadow.left + 1, rcShadow.top+1);
		pdc->LineTo(rcShadow.left + 1,rcShadow.bottom - 1);

		pdc->SelectObject(hpInsideShadowOld);
	}

}



// called by the SetBindings() function of both the Row and ColumnsCursor to determine the
// minimum number of bytes of in-line memory required by the new column bindings instead of
// cbRowLength, in case this parameter is 0. This function is not implemented and is merely
// a placeholder. You will need to implement this function according the documentation for
// the SetBinding Errors: (DB_E_BADBINDINFO, DB_E_COLUMNUNAVAILABLE, and DB_E_ROWTOOSHORT)
// in the Appendix A of the Interface Refernce of the Data Binding SDK.

HRESULT CDscdaoCtrl::FindDataRowLength()
{   
   if(m_cbRowSize && m_pRowCursor)m_pRowCursor->m_lRowDataLength=m_cbRowSize;
	return S_OK;
}

int CDscdaoCtrl::GetExtraButtonCount()
{
	int count=0;
	if(m_showAddnewButton)count++;
	if(m_showCancelButton)count++;
	if(m_showUpdateButton)count++;
	return count;
}


// Find out which button was selected with the mouse if any
int CDscdaoCtrl::GetHitButton(CPoint point, CRect /*rcBounds*/)
{
	CRect rcButton;
	CPoint ptOffset;
	if (m_Buttons[DSC_MOVEFIRST].PtInRect(point))return DSC_MOVEFIRST;
	else if(m_Buttons[DSC_MOVEPREVIOUS].PtInRect(point))return DSC_MOVEPREVIOUS;
	else if(m_Buttons[DSC_MOVENEXT].PtInRect(point))return DSC_MOVENEXT;
	else if(m_Buttons[DSC_MOVELAST].PtInRect(point))return DSC_MOVELAST;
	else if(m_showAddnewButton && m_Buttons[DSC_ADDNEW].PtInRect(point))return DSC_ADDNEW;
	else if(m_showCancelButton && m_Buttons[DSC_CANCEL].PtInRect(point))return DSC_CANCEL;
	else if(m_showUpdateButton && m_Buttons[DSC_UPDATE].PtInRect(point))return DSC_UPDATE;
	else	return -1; // No Hit
}


// called by the GetColumnsCursor() function of the RowCursor to create the meta data cursor rows
// As static data is used in this sample, this function directly fills in the values for the
// appropriate member variables of the Control Class that hold the meta data info. You will need
// to re-implement this function according to the backend that this data source control will be
// used for. All these meta data related variables are members of the Control class and not
// CRowCursor or CColCursor, because these are the same for the current resultset even for clones
// of the Row Cursor.

HRESULT CDscdaoCtrl::GetMetaData()
{
	// no need to create meta data again if already created the first time GetColumnsCursor() is
	// called on a Row Cursor!
	if (m_pMetaData){
		for(unsigned long int i=0;i<m_nColumns + m_nBookmarks;i++){
			if(m_pMetaData[i].szColName){
				::SysFreeString((BSTR)m_pMetaData[i].szColName);		
				m_pMetaData[i].szColName=NULL;
			}
		}
		CoTaskMemFree ((LPVOID)m_pMetaData);
		m_pMetaData=NULL;
	}
	if(m_pColOffsetInRow){
		CoTaskMemFree(m_pColOffsetInRow);
		m_pColOffsetInRow=NULL;
	}
	m_nColumns = 0;
	m_nBookmarks = 0;
	m_cbRowSize = 0;

   CMyDaoRecordset *m_pXRecordSet=new CMyDaoRecordset(this,m_pdatabase);
	if (!m_pXRecordSet)return E_OUTOFMEMORY;
   if(m_sqlType==0){
	CString mquery(_T("select * from "));
	mquery=mquery+m_Table;
	try{
      m_pXRecordSet->Open(
			(m_openType==0 ? dbOpenTable : (m_openType==1 ? dbOpenDynaset : dbOpenSnapshot)),
			(m_openType ? (LPCSTR)mquery : NULL),
			((m_openType==2 || m_readOnly) ? dbReadOnly : 0));
   }catch(CDaoException *e){
		e->Delete();
      delete m_pXRecordSet;
		return E_FAIL;
	}
   }else if(m_sqlType==1){
      CDaoQueryDef tquery(m_pdatabase);
	try{
      tquery.Open(m_Table);
      m_pXRecordSet->Open(&tquery,(m_openType==0 ? dbOpenDynaset : dbOpenSnapshot),((m_openType==2 || m_readOnly) ? dbReadOnly : 0));
   }catch(CDaoException *e){
		e->Delete();
      delete m_pXRecordSet;
		return E_FAIL;
	}
   }else{
      CDaoQueryDef tquery(m_pdatabase);
	try{
      tquery.Create(NULL,m_sqlString);
      m_pXRecordSet->Open(&tquery,(m_openType==0 ? dbOpenDynaset : dbOpenSnapshot),((m_openType==2 || m_readOnly) ? dbReadOnly : 0));
   }catch(CDaoException *e){
		e->Delete();
      delete m_pXRecordSet;
		return E_FAIL;
	}
   }
   CDaoFieldInfo info;
	unsigned long int i;
	if (m_pMetaData)return S_OK;
	m_nBookmarks = 1;
	if(m_pXRecordSet)m_nColumns = m_pXRecordSet->GetFieldCount();
	


	m_pColOffsetInRow = (ULONG *)CoTaskMemAlloc(sizeof(ULONG) * (m_nColumns + m_nBookmarks+1));

	if (!m_pColOffsetInRow){
		delete m_pXRecordSet;
		return E_OUTOFMEMORY;
	}
	
	m_pMetaData = (LPDBMETADATA)CoTaskMemAlloc(sizeof(DBMETADATA) * (m_nColumns + m_nBookmarks));

	if (!m_pMetaData){
		delete m_pXRecordSet;
		return E_OUTOFMEMORY;
	}
	GUID gnumonly = GUID_NUMBERONLY;

	i=0;
	for(;i<m_nColumns;i++){
	// Set up the meta data columns
		m_pXRecordSet->GetFieldInfo( i, info);
		m_pMetaData[i].ColumnID.dwKind = DBCOLKIND_GUID_NUMBER;
		m_pMetaData[i].ColumnID.guid = gnumonly;
		m_pMetaData[i].ColumnID.lNumber = i;
		m_pMetaData[i].szColName=info.m_strName.AllocSysString();
		switch(info.m_nType){
			case dbBoolean:
				m_pMetaData[i].dwColType = VT_BOOL;
				m_pMetaData[i].nColMaxLength=1;
				break;
			case dbByte:
				m_pMetaData[i].dwColType = VT_UI1;
				m_pMetaData[i].nColMaxLength=1;
				break;
			case dbInteger:
				m_pMetaData[i].dwColType = VT_I2;
				m_pMetaData[i].nColMaxLength=2;
				break;
			case dbLong:
				m_pMetaData[i].dwColType = VT_I4;
				m_pMetaData[i].nColMaxLength=4;
				break;
			case dbCurrency:
				m_pMetaData[i].dwColType = VT_CY;
				m_pMetaData[i].nColMaxLength=8;
				break;
			case dbSingle:
				m_pMetaData[i].dwColType = VT_R4;
				m_pMetaData[i].nColMaxLength=4;
				break;
			case dbDouble:
				m_pMetaData[i].dwColType = VT_R8;
				m_pMetaData[i].nColMaxLength=8;
				break;
			case dbDate:
				m_pMetaData[i].dwColType = VT_DATE;
				m_pMetaData[i].nColMaxLength=8;
				break;
			case dbText:
				m_pMetaData[i].dwColType = VT_LPSTR;
				m_pMetaData[i].nColMaxLength=info.m_lSize;
				break;
			case dbLongBinary:
			case dbMemo:
				m_pMetaData[i].dwColType = VT_BLOB;
				m_pMetaData[i].nColMaxLength=DB_NOMAXLENGTH;
				break;
			case dbGUID:
				m_pMetaData[i].dwColType = DBTYPE_BYTES;
				m_pMetaData[i].nColMaxLength=16;
				break;
		}
		m_pMetaData[i].dwBindType = DBBINDTYPE_DATA;
		m_pMetaData[i].bDataCol = TRUE;
		m_pMetaData[i].nEntryIDMaxLen = 0;
		m_pMetaData[i].dwUpdatable =((info.m_lAttributes & dbUpdatableField) ? DBUPDATEABLE_UPDATEABLE : DBUPDATEABLE_NOTUPDATEABLE);
		m_pColOffsetInRow[i]=m_cbRowSize;
		if(info.m_nType==dbLongBinary || info.m_nType==dbMemo)m_cbRowSize+=sizeof(BLOB);
		else m_cbRowSize+=((info.m_nType!=dbText/*&& info.m_nType!=dbGUID*/) ? m_pMetaData[i].nColMaxLength : 4L);
		}

	for(;i<m_nColumns+m_nBookmarks;i++){
		m_pMetaData[i].ColumnID = COLUMN_BMKTEMPORARY;
		m_pMetaData[i].szColName = NULL;
		m_pMetaData[i].dwColType = VT_BLOB;
		m_pMetaData[i].dwBindType = DBBINDTYPE_DATA;
		m_pMetaData[i].bDataCol = FALSE;
		m_pMetaData[i].nColMaxLength = DB_NOMAXLENGTH;
		m_pMetaData[i].nEntryIDMaxLen = 0;
		m_pMetaData[i].dwUpdatable = DBUPDATEABLE_NOTUPDATEABLE;
		m_pColOffsetInRow[i]=m_cbRowSize;
		m_cbRowSize+=sizeof(BLOB);
	}
	m_pColOffsetInRow[i]=m_cbRowSize;
	delete m_pXRecordSet;
	return S_OK;
}

int CDscdaoCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	// TODO: Add your specialized creation code here
	return 0;
}

void CDscdaoCtrl::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	// TODO: Add your message handler code here and/or call default
	
	OnLButtonDown(nFlags,point); 
}

void CDscdaoCtrl::OnLButtonDown(UINT /*nFlags*/, CPoint point) 
{
	// TODO: Add your message handler code here and/or call default
	CRect rcBounds;
	CDC * pdc;
	CRect bRect;
	
	GetClientRect(&rcBounds);
	
	m_CurrentButton = GetHitButton(point, rcBounds);

	if (m_CurrentButton >= 0)
	{
		pdc=GetDC();
		DrawButtonAt(pdc, rcBounds,m_CurrentButton, DSC_BUTTONDOWN);
		ReleaseDC(pdc);
		SetCapture();
	}
}

void CDscdaoCtrl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	// TODO: Add your message handler code here and/or call default
	CRect rcBounds;
	CDC * pdc;
	CRect bRect;
	int tmphit;
	
	if(m_CurrentButton >=0)
	{
		GetClientRect(&rcBounds);
		pdc=GetDC();
		tmphit = GetHitButton(point, rcBounds);
		switch(m_CurrentButton)
		{
		case DSC_MOVEFIRST:
			if(tmphit == DSC_MOVEFIRST)
				Beginning();
			break;
		case DSC_MOVEPREVIOUS:
			if(tmphit == DSC_MOVEPREVIOUS)
				Previous();
			break;
		case DSC_MOVENEXT:
			if(tmphit == DSC_MOVENEXT)
				Next();
			break;
		case DSC_MOVELAST:
			if(tmphit == DSC_MOVELAST)
				End();
			break;
		case DSC_ADDNEW:
			if(tmphit == DSC_ADDNEW)
				AddNew();
			break;
		case DSC_CANCEL:
			if(tmphit == DSC_CANCEL)
				CancelUpdate();
			break;
		case DSC_UPDATE:
			if(tmphit == DSC_UPDATE)
				Update();
			break;
		}

		DrawButtonAt(pdc, rcBounds, m_CurrentButton, DSC_BUTTONUP);
		ReleaseDC(pdc);
		ReleaseCapture();
		m_CurrentButton = -1;
	}
	COleControl::OnLButtonUp(nFlags, point);
}

void CDscdaoCtrl::OnMouseMove(UINT nFlags, CPoint point) 
{
	// TODO: Add your message handler code here and/or call default
	CRect rcBounds;
	CDC * pdc;
	CRect bRect;
	int tmphit;
	static int State;
	
	if(m_CurrentButton >= 0) // MouseDown
	{
		GetClientRect(&rcBounds);
		tmphit = GetHitButton(point, rcBounds);
		if (tmphit<0)
		{
			if (State != DSC_BUTTONUP)
			{
				pdc=GetDC();
				DrawButtonAt(pdc, rcBounds, m_CurrentButton, DSC_BUTTONUP);
				State = DSC_BUTTONUP;
				ReleaseDC(pdc);
			}
		}
		else if(tmphit == m_CurrentButton) 
		{
			if (State != DSC_BUTTONDOWN)
			{
				pdc=GetDC();
				DrawButtonAt(pdc, rcBounds, m_CurrentButton, DSC_BUTTONDOWN);
				State = DSC_BUTTONDOWN;
				ReleaseDC(pdc);
			}
		}
		else if(State!=DSC_BUTTONUP)
		{
				pdc=GetDC();
				DrawButtonAt(pdc, rcBounds, m_CurrentButton, DSC_BUTTONUP);
				State = DSC_BUTTONUP;
				ReleaseDC(pdc);
		}


	}
	COleControl::OnMouseMove(nFlags, point);
}

BOOL CDscdaoCtrl::OnSetExtent(LPSIZEL lpSizeL) 
{
	// TODO: Add your specialized code here and/or call the base class
	TEXTMETRIC tm;
	HDC hdc = ::GetDC(NULL);
	CDC dc;
	
	if (m_sAppearance)
	{
		m_nBottomBorderWidth = 3;
		m_nTopBorderWidth = 2;
	}
	else
	{
		m_nBottomBorderWidth = 1;
		m_nTopBorderWidth = 1;
	}
	
	int nExtraButtons=GetExtraButtonCount();

	dc.Attach(hdc);
 	dc.HIMETRICtoDP(lpSizeL);
	
	GetStockTextMetrics(&tm);

	if(m_sAppearance)
	{
		lpSizeL->cy = max(lpSizeL->cy, tm.tmHeight + 8);
		lpSizeL->cx = max(lpSizeL->cx, DSC_BUTTONWIDTH*(4 + nExtraButtons) + 2 + m_nBottomBorderWidth + m_nTopBorderWidth + nExtraButtons);
    }
	else
	{	
    	lpSizeL->cy = max(lpSizeL->cy, tm.tmHeight + 6);
		lpSizeL->cx = max(lpSizeL->cx, DSC_BUTTONWIDTH*(4 + nExtraButtons) + 3 + m_nBottomBorderWidth + m_nTopBorderWidth + nExtraButtons);
    }

	dc.DPtoHIMETRIC(lpSizeL);
	
	dc.Detach();
	::ReleaseDC(NULL, hdc);

	return COleControl::OnSetExtent(lpSizeL);
}


/////////////////////////////////////////////////////////////////////////////
// Helper function implementations


// called by the SetBindings() function of both the Row and ColumnsCursor to see if the entries
// in the rgBoundColumns are valid, before actually setting the bindings. This function is not
// implemented and is merely a placeholder. You will need to implement this function according
// the documentation for the SetBinding Errors: (DB_E_BADBINDINFO, DB_E_COLUMNUNAVAILABLE, and
// DB_E_ROWTOOSHORT) in the Appendix A of the Interface Reference of the Data Binding SDK.

HRESULT CDscdaoCtrl::ValidateColumnBindings(ULONG cCol, DBCOLUMNBINDING rgBoundColumns[])
{
  if(cCol && !rgBoundColumns)return DB_E_BADBINDINFO;

	for (ULONG i = 0; i < cCol; i++){
         if(!(
						rgBoundColumns[i].dwDataType==VT_BSTR 
						|| rgBoundColumns[i].dwDataType==VT_LPSTR	
						|| rgBoundColumns[i].dwDataType==VT_LPWSTR
						|| rgBoundColumns[i].dwDataType==VT_VARIANT
						|| rgBoundColumns[i].dwDataType==VT_BLOB
						|| rgBoundColumns[i].dwDataType==DBTYPE_BYTES
						|| rgBoundColumns[i].dwDataType==DBTYPE_CHARS
						|| rgBoundColumns[i].dwDataType==DBTYPE_WCHARS
                  || rgBoundColumns[i].dwDataType==VT_I1
                  || rgBoundColumns[i].dwDataType==VT_UI1
                  || rgBoundColumns[i].dwDataType==VT_I2
                  || rgBoundColumns[i].dwDataType==VT_UI2
						|| rgBoundColumns[i].dwDataType==VT_I4
                  || rgBoundColumns[i].dwDataType==VT_UI4
                  || rgBoundColumns[i].dwDataType==VT_INT
                  || rgBoundColumns[i].dwDataType==VT_UINT
						|| rgBoundColumns[i].dwDataType==VT_BOOL
						|| rgBoundColumns[i].dwDataType==VT_R4
						|| rgBoundColumns[i].dwDataType==VT_R8
						|| rgBoundColumns[i].dwDataType==VT_DATE
						|| rgBoundColumns[i].dwDataType==VT_CY
                  || rgBoundColumns[i].dwDataType==VT_CY
                  || rgBoundColumns[i].dwDataType==DBTYPE_ANYVARIANT
                  || rgBoundColumns[i].dwDataType==DBTYPE_COLUMNID
                  || rgBoundColumns[i].dwDataType==VT_PTR
             ))return DB_E_BADBINDINFO;

         if(rgBoundColumns[i].cbMaxLen==DB_NOMAXLENGTH && rgBoundColumns[i].dwBinding==DBBINDING_DEFAULT
            && (rgBoundColumns[i].dwDataType==DBTYPE_CHARS || rgBoundColumns[i].dwDataType==DBTYPE_WCHARS || rgBoundColumns[i].dwDataType==DBTYPE_BYTES))
            return DB_E_BADBINDINFO;

//         if(rgBoundColumns[i].dwBinding!=DBBINDING_DEFAULT && ((rgBoundColumns[i].dwBinding | DBBINDING_VARIANT) ||
//            (rgBoundColumns[i].dwBinding | DBBINDING_ENTRYID)))
//            return DB_E_BADBINDINFO;


         if(rgBoundColumns[i].dwBinding==DBBINDING_VARIANT && (rgBoundColumns[i].dwDataType==DBTYPE_BYTES
            || rgBoundColumns[i].dwDataType==DBTYPE_CHARS || rgBoundColumns[i].dwDataType==DBTYPE_WCHARS))
            return DB_E_BADBINDINFO;

         if(rgBoundColumns[i].dwBinding!=DBBINDING_VARIANT && rgBoundColumns[i].dwDataType==DBTYPE_ANYVARIANT)
            return DB_E_BADBINDINFO;
   }

   return S_OK;
}

void CDscdaoCtrl::OnAppearanceChanged() 
{
	// TODO: Add your specialized code here and/or call the base class
	InvalidateControl();
	SetModifiedFlag();
}


/////////////////////////////////////////////////////////////////////////////
// IVBDSC implementation

// This function is not implemented and is merely a placeholder. You will need to implement
// this function according to the backend that this data source control will be used for.

HRESULT CDscdaoCtrl::XVBDSC::CancelUnload(BOOL FAR * /*pfCancel*/)
{
	METHOD_PROLOGUE(CDscdaoCtrl, VBDSC)
	return S_OK;
}

// This function is not implemented and is merely a placeholder. You will need to implement
// this function according to the backend that this data source control will be used for.

HRESULT CDscdaoCtrl::XVBDSC::Error(DWORD /*dwErr*/, BOOL FAR * /*pfShowError*/)
{
	METHOD_PROLOGUE(CDscdaoCtrl, VBDSC)
	return S_OK;
}

// This function is called by VB4's Binding Manager, once the DSC has been created, to create
// RowCursor and get a pointer to an ICursor interface implemented by it.

HRESULT CDscdaoCtrl::XVBDSC::CreateCursor(ICursor FAR * FAR *ppCursor)
{
	METHOD_PROLOGUE(CDscdaoCtrl, VBDSC)
	
   if(!pThis->m_pRowCursor)
      pThis->m_pRowCursor = new CRowCursor(pThis);

	if (!pThis->m_pRowCursor)return E_OUTOFMEMORY;
	
   pThis->m_pRowCursor->m_xCURSOR.QueryInterface(IID_ICursor,(void **)ppCursor);
//   pThis->m_pRowCursor->ExternalRelease();

   return S_OK;
}

HRESULT CDscdaoCtrl::XVBDSC::QueryInterface(REFIID riid, LPVOID FAR *ppvObj)
{
	METHOD_PROLOGUE(CDscdaoCtrl, VBDSC)
	return (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
}

ULONG CDscdaoCtrl::XVBDSC::AddRef()
{
	METHOD_PROLOGUE(CDscdaoCtrl, VBDSC)
	return (ULONG)pThis->ExternalAddRef();
}

ULONG CDscdaoCtrl::XVBDSC::Release()
{
	METHOD_PROLOGUE(CDscdaoCtrl, VBDSC)
	return (ULONG)pThis->ExternalRelease();
}

void CDscdaoCtrl::OnDatabasePathChanged() 
{
	// TODO: Add notification handler code
try{
	try{
		if(m_pRowCursor)m_pRowCursor->Close();
	}catch(CDaoException *f){
		f->Delete();
	}
	try{
		if(m_pdatabase->IsOpen())m_pdatabase->Close();
      m_pdatabase->Open(m_databasePath,m_openExclusive,FALSE,CString(";PWD=")+m_password);
	}catch(CDaoException *e){
		e->Delete();
	}
	if (m_pColOffsetInRow){
		CoTaskMemFree((LPVOID)m_pColOffsetInRow);
		m_pColOffsetInRow=NULL;
	}
	if (m_pMetaData){
		for(unsigned long int i=0;i<m_nColumns + m_nBookmarks;i++){
			if(m_pMetaData[i].szColName){
				::SysFreeString((BSTR)m_pMetaData[i].szColName);		
				m_pMetaData[i].szColName=NULL;
			}
		}
		CoTaskMemFree ((LPVOID)m_pMetaData);
		m_nColumns=m_nBookmarks=0;
		m_pMetaData=NULL;
	}
}catch(CDaoException *e){
	e->Delete();
}
SetModifiedFlag();
}

void CDscdaoCtrl::OnTableNameChanged() 
{
	// TODO: Add notification handler code
if(m_pRowCursor){
	try{
		if(m_pRowCursor)m_pRowCursor->Close();
	}catch(CDaoException *e){
		e->Delete();
	}
	if(m_pColOffsetInRow){
		CoTaskMemFree((LPVOID)m_pColOffsetInRow);
		m_pColOffsetInRow=NULL;
	}

	if(m_pMetaData){
		for(unsigned long int i=0;i<m_nColumns + m_nBookmarks;i++){
			if(m_pMetaData[i].szColName){
				::SysFreeString((BSTR)m_pMetaData[i].szColName);		
				//m_pMetaData[i].szColName=NULL;
			}
		}
		CoTaskMemFree((LPVOID)m_pMetaData);
		m_nColumns=m_nBookmarks=0;
		m_pMetaData=NULL;
	}
	m_pRowCursor->m_xCURSOR.Requery();
	if(m_pRowCursor->m_pColCursor)m_pRowCursor->m_pColCursor->m_xCOLUMNSCURSOR.Requery();
}
SetModifiedFlag();
}


LPCONNECTIONPOINT CDscdaoCtrl::GetConnectionHook(REFIID iid)
{
	
   LPCONNECTIONPOINT lpc;
   lpc=COleControl::GetConnectionHook(iid);
   if(lpc)return lpc;
   
   if(!IsEqualIID((IID)iid, IID_INotifyDBEvents))return NULL;	

   if(!m_pRowCursor){
      m_pRowCursor=new CRowCursor(this);
      //InterlockedDecrement(&m_pRowCursor->m_dwRef);
   }

   return m_pRowCursor->GetConnectionHook(iid);
}

void CDscdaoCtrl::OnIndexNameChanged() 
{
	// TODO: Add notification handler code
	OnTableNameChanged();
}

void CDscdaoCtrl::OnReadOnlyChanged() 
{
	// TODO: Add notification handler code
	OnTableNameChanged();
}


void CDscdaoCtrl::OnOpenTypeChanged() 
{
	// TODO: Add notification handler code
	OnTableNameChanged();
}


void CDscdaoCtrl::OnSqlTypeChanged() 
{
	// TODO: Add notification handler code
	OnTableNameChanged();
}

void CDscdaoCtrl::OnSqlStringChanged() 
{
	// TODO: Add notification handler code
	OnTableNameChanged();
}

void CDscdaoCtrl::OnDrawColorChanged() 
{
	// TODO: Add notification handler code
	SetModifiedFlag();
   InvalidateControl();
}

void CDscdaoCtrl::OnOpenExclusiveChanged() 
{
	// TODO: Add notification handler code
OnDatabasePathChanged();
OnTableNameChanged();
}

void CDscdaoCtrl::OnPasswordChanged() 
{
	// TODO: Add notification handler code
OnDatabasePathChanged();
OnTableNameChanged();
}




void CDscdaoCtrl::BeginTransaction() 
{
	// TODO: Add your dispatch handler code here
try{
   if(m_pdatabase && m_pdatabase->m_pWorkspace)m_pdatabase->m_pWorkspace->BeginTrans();
}catch(CDaoException *e){
   e->Delete();
}
}

void CDscdaoCtrl::CommitTransaction() 
{
	// TODO: Add your dispatch handler code here
try{
   if(m_pdatabase && m_pdatabase->m_pWorkspace)m_pdatabase->m_pWorkspace->CommitTrans();
}catch(CDaoException *e){
   e->Delete();
}
}

void CDscdaoCtrl::Rollback() 
{
	// TODO: Add your dispatch handler code here
try{
   if(m_pdatabase && m_pdatabase->m_pWorkspace)m_pdatabase->m_pWorkspace->Rollback();
}catch(CDaoException *e){
   e->Delete();
}
}

void CDscdaoCtrl::OnLockingModeChanged() 
{
	// TODO: Add notification handler code
   if(m_pRowCursor && m_pRowCursor->m_pRecordSet && m_pRowCursor->m_pRecordSet->IsOpen())m_pRowCursor->m_pRecordSet->SetLockingMode(m_lockingMode);
	SetModifiedFlag();
}


//void CDscdaoCtrl::OnDestroy() 
//{
//	COleControl::OnDestroy();
	
	// TODO: Add your message handler code here
//   if(m_pRowCursor){
         //   m_pRowCursor->Close();
//         m_pRowCursor->ExternalRelease();
//         m_pRowCursor=NULL;
//   }
//   m_pdatabase->Close();
//   delete m_pdatabase;
//   m_pdatabase=NULL;
//   if (m_pColOffsetInRow)CoTaskMemFree((LPVOID)m_pColOffsetInRow);
//   m_pColOffsetInRow=NULL;
//   if (m_pMetaData){
//	   for(unsigned long int i=0;i<m_nColumns + m_nBookmarks;i++){
//		   if(m_pMetaData[i].szColName){
//			   ::SysFreeString((BSTR)m_pMetaData[i].szColName);		
//		   }
//      }
//	   CoTaskMemFree ((LPVOID)m_pMetaData);
//      m_pMetaData=NULL;
//   }
//   count--;
//   if(count<=0)AfxDaoTerm();
//}

