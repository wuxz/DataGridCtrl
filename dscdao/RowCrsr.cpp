// RowCrsr.cpp : implementation file
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


#include "stdafx.h"

#define GO_BACK

#define DBINITCONSTANTS

#include <comdef.h>


#include <ocdbid.h>
#include <ocdb.h>

#include "vbdsc.h"

#include "DscdaoCtl.h"
#include "ColCrsr.h"
#include "RowCrsr.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


/////////////////////////////////////////////////////////////////////////////
// CRowCursor Implementation

IMPLEMENT_DYNCREATE(CRowCursor, CCmdTarget)

/////////////////////////////////////////////////////////////////////////////
// Connection map

BEGIN_CONNECTION_MAP(CRowCursor, CCmdTarget)
    CONNECTION_PART(CRowCursor, IID_INotifyDBEvents, ConnectionPoint)
END_CONNECTION_MAP()


/////////////////////////////////////////////////////////////////////////////
// Interface map

// Support these interfaces
BEGIN_INTERFACE_MAP(CRowCursor, CCmdTarget)
	INTERFACE_PART(CRowCursor, IID_ICursor, CURSOR)
	INTERFACE_PART(CRowCursor, IID_ICursorMove, CURSORMOVE)
	INTERFACE_PART(CRowCursor, IID_ICursorUpdateARow, CursorUpdateARow)
	INTERFACE_PART(CRowCursor, IID_ICursorScroll, CURSORSCROLL)
	INTERFACE_PART(CRowCursor, IID_ICursorFind, CURSORFIND)
	INTERFACE_PART(CRowCursor, IID_IConnectionPointContainer, ConnPtContainer)
END_INTERFACE_MAP()

CRowCursor::CRowCursor(CDscdaoCtrl *pControl,BOOL create_cursor)
{
   TRACE("\nRow class constructor\n");
	// Enable connections for our connection point
	is_clone=FALSE;
	parent=clone=NULL;
	m_pControl = pControl;
//   pControl->ExternalAddRef();
	m_pColCursor = NULL;
	// Start on the First Data Row rather than the DBBMK_BEGINNING.
	// The internal Data Control does this.
	m_lRowDataLength = 0;
	m_pDataColumns = NULL;
	m_nExistingDataCols = 0;
	m_dwEditMode = DBEDITMODE_NONE;
	m_bUpdateInProgress = FALSE;
   try{
     if(!m_pControl->m_pdatabase->IsOpen())
         m_pControl->m_pdatabase->Open(m_pControl->m_databasePath,m_pControl->m_openExclusive,FALSE,CString(";PWD=")+m_pControl->m_password);
   }catch(CDaoException *f){
         f->Delete();
   }
   if(create_cursor){
   m_pRecordSet=new CMyDaoRecordset(m_pControl,m_pControl->m_pdatabase);
   if(m_pControl->m_sqlType==0){
	CString mquery(_T("select * from "));
	mquery=mquery+m_pControl->m_Table;
	try{
		m_pRecordSet->Open((m_pControl->m_openType==0 ? dbOpenTable : (m_pControl->m_openType==1 ? dbOpenDynaset : dbOpenSnapshot)),
			(m_pControl->m_openType ? (LPCSTR)mquery : NULL),
		   ((m_pControl->m_openType==2 || m_pControl->m_readOnly) ? dbReadOnly : 0));	
      if(!m_pControl->m_openType && !m_pControl->m_indexName.IsEmpty() && m_pControl->m_indexName!="None"){
         m_pRecordSet->SetCurrentIndex(m_pControl->m_indexName);
      }
	}catch(CDaoException *e){
		e->Delete();
	}
   }else if(m_pControl->m_sqlType==1){
      CDaoQueryDef tquery(m_pControl->m_pdatabase);
	try{
      tquery.Open(m_pControl->m_Table);
      m_pRecordSet->Open(&tquery,(m_pControl->m_openType==0 ? dbOpenDynaset : dbOpenSnapshot),((m_pControl->m_openType==2 || m_pControl->m_readOnly) ? dbReadOnly : 0));
   }catch(CDaoException *e){
		e->Delete();
	}
   }else{
      CDaoQueryDef tquery(m_pControl->m_pdatabase);
	try{
      tquery.Create(NULL,m_pControl->m_sqlString);
      m_pRecordSet->Open(&tquery,(m_pControl->m_openType==0 ? dbOpenDynaset : dbOpenSnapshot),((m_pControl->m_openType==2 || m_pControl->m_readOnly) ? dbReadOnly : 0));
   }catch(CDaoException *e){
		e->Delete();
	}
   }
   }else m_pRecordSet=NULL;
	EnableConnections();
}

CRowCursor::~CRowCursor()
{
TRACE("\nRow class destructor\n");
if (m_pDataColumns){
	CoTaskMemFree ((LPVOID)m_pDataColumns);
//	m_pDataColumns=NULL;
}
//if (m_pColCursor){
//   m_pColCursor->ExternalRelease();
//	m_pColCursor=0;
//}
if(m_pRecordSet){
	delete m_pRecordSet;
}
}

/////////////////////////////////////////////////////////////////////////////
// ICursor Implementation for the Row Cursor

HRESULT CRowCursor::XCURSOR::GetColumnsCursor(REFIID riid,IUnknown **ppvColumnsCursor,ULONG *pcRows)
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	if (pcRows)*pcRows=0;
	if (!ppvColumnsCursor)return E_INVALIDARG;
	*ppvColumnsCursor = NULL;
	// Create new instance of our columns cursor class for the meta data
   pThis->m_pColCursor = new CColCursor(pThis->m_pControl, pThis);
	// Various error checking
	if (!pThis->m_pColCursor)return E_OUTOFMEMORY;
	// Creates the columns cursor from static data. Change the code in GetMetaData
	// for your particular back end.
	if (FAILED(pThis->m_pControl->GetMetaData())){
		delete pThis->m_pColCursor;
      pThis->m_pColCursor=NULL;
		return E_OUTOFMEMORY;
	}
	// QI the new instance for the specified cursor interface and return
	if FAILED(pThis->m_pColCursor->m_xCOLUMNSCURSOR.QueryInterface(riid,(void **)ppvColumnsCursor))	{
		delete pThis->m_pColCursor;
      pThis->m_pColCursor=NULL;
		return E_NOINTERFACE;
	}
   
   pThis->m_pColCursor->ExternalRelease();

	if (pcRows)	*pcRows = pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks;
	return S_OK;
}

HRESULT CRowCursor::XCURSOR::SetBindings(ULONG cCol, DBCOLUMNBINDING rgBoundColumns[], ULONG cbRowLength, DWORD dwFlags)
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)

	// Validate Column Bindings. There is no functionality in this
	// function
	HRESULT hr = pThis->m_pControl->ValidateColumnBindings(cCol, rgBoundColumns);
   if(FAILED(hr))return hr;
	// Do the actual setting of the column bindings
	hr = pThis->m_pControl->BindColumns(&pThis->m_pDataColumns, &pThis->m_nExistingDataCols, cCol, rgBoundColumns, dwFlags);
   if(FAILED(hr))return hr;
	// If cbRowLength is 0, then we need to find the minimum number of bytes 
	// required by the new column bindings. Currently not supported. Add code
	// to FindDataRowLength to support this feature.
	if (!cbRowLength)	pThis->m_pControl->FindDataRowLength();
	else	pThis->m_lRowDataLength = cbRowLength;
	return hr;
}

HRESULT CRowCursor::XCURSOR::GetBindings(ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength)
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	if (!pcCol || !prgBoundColumns)return E_FAIL;
	*pcCol = pThis->m_nExistingDataCols;
	if (!(*pcCol)){
		// Allocate space to return current bindings
		*prgBoundColumns = (DBCOLUMNBINDING *)CoTaskMemAlloc(sizeof(DBCOLUMNBINDING) * (*pcCol));
		if (*prgBoundColumns){
			// Copy all bindings into the newly allocated memory.
			for (ULONG i = 0; i < *pcCol; i++)
				(*prgBoundColumns)[i] = pThis->m_pDataColumns[i];
		}else return E_OUTOFMEMORY;
	}else *prgBoundColumns = NULL;
	if(pcbRowLength)*pcbRowLength = pThis->m_lRowDataLength;
	return S_OK;
}

HRESULT CRowCursor::XCURSOR::GetNextRows(LARGE_INTEGER udlRowsToSkip, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	// Add the udlRowsToSkip parameter so the current row now 
	// reflects where we want to start fetching data
if(pThis->m_bUpdateInProgress)return DB_E_UPDATEINPROGRESS;
HRESULT hr;
DBNOTIFYREASON rgReasons[1];
SAFEARRAY *psa;

	if(pFetchParams)pFetchParams->cRowsReturned = 0;

	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);
	pThis->MakeSafearrayFromBuf(&psa, (void *)&DBBMK_CURRENT, DB_BMK_SIZE);
	// Bookmarks should be passed in like this to allow VariantChangeByte
	// to be used by bound controls
	rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
	rgReasons[0].arg1.parray = psa;
	
	rgReasons[0].arg2.vt = VT_I4;
	rgReasons[0].arg2.lVal = (long)udlRowsToSkip.LowPart; //+1;


	rgReasons[0].dwReason = DBREASON_MOVE;
	hr = pThis->DoNotifyBefore(DBEVENT_CURRENT_ROW_CHANGED, 1, rgReasons);
	if (hr != S_OK){
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}


long int of=udlRowsToSkip.LowPart+1;
try{
   if(pThis->m_pRecordSet->IsBOF() && of>0){
	   pThis->m_pRecordSet->MoveFirst();
	   of--;
   }else if(pThis->m_pRecordSet->IsEOF() && of<0){
	   pThis->m_pRecordSet->MoveLast();
	   of++;
   }
	pThis->m_pRecordSet->Move(of);
}catch(CDaoException *e){
	e->Delete();
}
	//Make sure not to go over EOF
hr = pThis->FetchRows(pFetchParams);

if (FAILED(hr)){
	pThis->CallFailedToDo(DBEVENT_CURRENT_ROW_CHANGED, 1, rgReasons);
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	return hr;
}

SafeArrayDestroy(psa);
pThis->MakeSafearrayFromBok(&psa);
rgReasons[0].arg1.parray = psa;
rgReasons[0].arg2.lVal = (long)0;

pThis->DoNotifyAfter(DBEVENT_CURRENT_ROW_CHANGED, 1, rgReasons);
VariantClear((VARIANT *)&rgReasons[0].arg1);
VariantClear((VARIANT *)&rgReasons[0].arg2);
return hr;
}

HRESULT CRowCursor::XCURSOR::Requery()
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)

	HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);
	pThis->MakeSafearrayFromBuf(&psa, (void *)&DBBMK_CURRENT, DB_BMK_SIZE); 
	rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
	rgReasons[0].arg1.parray = psa;
	rgReasons[0].arg2.vt = VT_I4;
	rgReasons[0].arg2.lVal = 0;
	cReasons = 1;

	if(!pThis->is_clone){ 
	rgReasons[0].dwReason = DBREASON_REFRESH;
	dwEventWhat = DBEVENT_CURRENT_ROW_CHANGED | DBEVENT_CURRENT_ROW_DATA_CHANGED | 
			DBEVENT_NONCURRENT_ROW_DATA_CHANGED | DBEVENT_SET_OF_ROWS_CHANGED | 
			DBEVENT_ORDER_OF_ROWS_CHANGED;
		//DBEVENT_CURRENT_ROW_DATA_CHANGED |  DBEVENT_NONCURRENT_ROW_DATA_CHANGED | DBEVENT_SET_OF_ROWS_CHANGED;
		 
	if (S_OK != (hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons)))return hr;
	 
	if (pThis->m_bUpdateInProgress){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		return DB_E_STATEERROR;
	} 
	}



	// Reset current row pointer to the BOF. Cursors would normally populate
	// cache with data from backend at this point as well.
	//pThis->m_CurrentRow = 0;
	if(pThis->m_pRecordSet && !pThis->is_clone){
		try{
			if(pThis->m_pRecordSet->IsOpen() && pThis->m_pRecordSet->CanRestart()){
				pThis->m_pRecordSet->Requery();
			}else{
				if(pThis->m_pRecordSet->IsOpen())pThis->m_pRecordSet->Close();
            if(pThis->m_pControl->m_sqlType==0){
               CString mquery(_T("select * from "));
				   mquery=mquery+pThis->m_pControl->m_Table;
				   pThis->m_pRecordSet->Open((pThis->m_pControl->m_openType==0 ? dbOpenTable : (pThis->m_pControl->m_openType==1 ? dbOpenDynaset : dbOpenSnapshot)),
					   (pThis->m_pControl->m_openType ? (LPCSTR)mquery: NULL),
					   ((pThis->m_pControl->m_openType==2 || pThis->m_pControl->m_readOnly) ? dbReadOnly : 0));	
               if(!pThis->m_pControl->m_openType && !pThis->m_pControl->m_indexName.IsEmpty() && pThis->m_pControl->m_indexName!="None"){
                  pThis->m_pRecordSet->SetCurrentIndex(pThis->m_pControl->m_indexName);
               }
            }else if(pThis->m_pControl->m_sqlType==1){
               CDaoQueryDef tquery(pThis->m_pControl->m_pdatabase);
   	         try{
                  tquery.Open(pThis->m_pControl->m_Table);
                  pThis->m_pRecordSet->Open(&tquery,(pThis->m_pControl->m_openType==0 ? dbOpenDynaset : dbOpenSnapshot),((pThis->m_pControl->m_openType==2 || pThis->m_pControl->m_readOnly) ? dbReadOnly : 0));
               }catch(CDaoException *e){
		            e->Delete();
		            return E_FAIL;
	            }
            }else{
               CDaoQueryDef tquery(pThis->m_pControl->m_pdatabase);
   	         try{
                  tquery.Create(NULL,pThis->m_pControl->m_sqlString);
                  pThis->m_pRecordSet->Open(&tquery,(pThis->m_pControl->m_openType==0 ? dbOpenDynaset : dbOpenSnapshot),((pThis->m_pControl->m_openType==2 || pThis->m_pControl->m_readOnly) ? dbReadOnly : 0));
               }catch(CDaoException *e){
		            e->Delete();
		            return E_FAIL;
	            }
            }
			}
		}catch(CDaoException *e){
			e->Delete();
		}
   }else if(pThis->m_pRecordSet){
      try{
         if(pThis->m_pRecordSet->IsOpen() && pThis->m_pRecordSet->CanRestart()){
				pThis->m_pRecordSet->Requery();
			}else{
				pThis->m_pRecordSet->Reclone(pThis->parent->m_pRecordSet);
         }
       }catch(CDaoException *e){
         e->Delete();
       }
   }
	
   if(pThis->clone)pThis->clone->m_xCURSOR.Requery();
	else pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);	
	return S_OK;
}

HRESULT CRowCursor::XCURSOR::QueryInterface(REFIID riid, LPVOID FAR *ppvObj)
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	return (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
}

ULONG CRowCursor::XCURSOR::AddRef()
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	return (ULONG)pThis->ExternalAddRef();
}

ULONG CRowCursor::XCURSOR::Release()
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	return (ULONG)pThis->ExternalRelease();
}

/////////////////////////////////////////////////////////////////////////////
// ICursorMove Implementation for the Row Cursor

STDMETHODIMP CRowCursor::XCURSORMOVE::GetBindings(ULONG *pcCol, DBCOLUMNBINDING *prgBoundColumns[], ULONG *pcbRowLength)
{
	METHOD_PROLOGUE(CRowCursor, CURSOR)
	// delegate to ICursor::GetBindings()
	return pThis->m_xCURSOR.GetBindings(pcCol, prgBoundColumns, pcbRowLength);
}	

STDMETHODIMP CRowCursor::XCURSORMOVE::GetColumnsCursor(REFIID riid, IUnknown *ppunkColumnsCursor[], ULONG *pcRows)
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	// delegate to ICursor::GetColumnsCursor()
   return pThis->m_xCURSOR.GetColumnsCursor(riid, ppunkColumnsCursor, pcRows);
}	

STDMETHODIMP CRowCursor::XCURSORMOVE::GetNextRows(LARGE_INTEGER dlRowsToSkip, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	// delegate to ICursor::GetNextRows()
	return pThis->m_xCURSOR.GetNextRows(dlRowsToSkip, pFetchParams);
}	

STDMETHODIMP CRowCursor::XCURSORMOVE::Requery()
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	// delegate to ICursor::Requery()
	return pThis->m_xCURSOR.Requery();
}	

STDMETHODIMP CRowCursor::XCURSORMOVE::SetBindings(ULONG cCol, DBCOLUMNBINDING rgBoundColumns[], ULONG cbRowLength, DWORD dwFlags)
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	// delegate to ICursor::SetBindings()
	return pThis->m_xCURSOR.SetBindings(cCol, rgBoundColumns, cbRowLength, dwFlags);
}	

// Move current row to a certain bookmark and fetch rows if requested
HRESULT CRowCursor::XCURSORMOVE::Move(ULONG cbBookmark, void *pBookmark, LARGE_INTEGER dlOffset, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)

   HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;
	
   if(pFetchParams)pFetchParams->cRowsReturned = 0;

   if(!pBookmark)return E_FAIL;

   if(!pThis->IsOpen())return E_FAIL;

	// Initialize arguments for notifications
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);

	pThis->MakeSafearrayFromBuf(&psa, pBookmark, cbBookmark);

   //pThis->MakeSafearrayFromBok(&psa);	

	// Bookmarks should be passed in like this to allow VariantChangeByte
	// to be used by bound controls
	rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
	rgReasons[0].arg1.parray = psa;
	
	rgReasons[0].arg2.vt = VT_I4;
	rgReasons[0].arg2.lVal = (long)dlOffset.LowPart;

	dwEventWhat = DBEVENT_CURRENT_ROW_CHANGED;
	rgReasons[0].dwReason = DBREASON_MOVE;
	cReasons = 1;
	// Don't notify if retrieving the current row
	if (*(BYTE *)pBookmark != DBBMK_CURRENT || dlOffset.LowPart ||
		(pFetchParams && pFetchParams->cRowsRequested>1)){		

      TRACE("\nNotify before move\n");
		
      hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons);
		if (hr != S_OK){
			VariantClear((VARIANT *)&rgReasons[0].arg1);
			VariantClear((VARIANT *)&rgReasons[0].arg2);
			return hr;
		}

      if(pThis->m_bUpdateInProgress){
			pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			return DB_E_STATEERROR;
		}		
	}

	
	
	
	long int offset;
	// Find which bookmark was passed in and adjust current row to it
	switch (*((BYTE*)pBookmark)){
		case DBBMK_BEGINNING:	
               
         TRACE("\nIn move from beginning with offset %d in %s\n",dlOffset.LowPart,(pThis->is_clone ? "Clone" : "Original"));

			      if(pThis->m_pRecordSet->IsBOF() && !dlOffset.LowPart)break;
					try{
						pThis->m_pRecordSet->MoveFirst();
						pThis->m_pRecordSet->Move(dlOffset.LowPart-1);
					}catch(CDaoException *e){
						e->Delete();
					}
					break;

		case DBBMK_CURRENT:
					
               TRACE("\nIn move from current with offset %d in %s\n",dlOffset.LowPart,(pThis->is_clone ? "Clone" : "Original"));

               if(!dlOffset.LowPart)break;
					offset=dlOffset.LowPart;
					try{
						if(pThis->m_pRecordSet->IsBOF()){
								pThis->m_pRecordSet->MoveFirst();
								offset--;
								if(offset<0){
									pThis->m_pRecordSet->MovePrev();
									break;
								}
						}else if(pThis->m_pRecordSet->IsEOF()){
								pThis->m_pRecordSet->MoveLast();
								offset++;
								if(offset>0){
									pThis->m_pRecordSet->MoveNext();
									break;
								}
						}
					}catch(CDaoException *e){
						e->Delete();
					}
					
					try{
						if(offset)pThis->m_pRecordSet->Move(offset);
					}catch(CDaoException *e){
						e->Delete();
					}
					break;

		case DBBMK_END:			
				
            TRACE("\nIn move from end with offset %d in %s\n",dlOffset.LowPart,(pThis->is_clone ? "Clone" : "Original"));
         
            if(pThis->m_pRecordSet->IsEOF() && !dlOffset.LowPart)break;
				try{
					pThis->m_pRecordSet->MoveLast();
					pThis->m_pRecordSet->Move(dlOffset.LowPart+1);
				}catch(CDaoException *e){
					e->Delete();
				}
				break;
		case DBBMK_INVALID:	
			pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			VariantClear((VARIANT *)&rgReasons[0].arg1);
			VariantClear((VARIANT *)&rgReasons[0].arg2);
			return DB_E_BADBOOKMARK;

		// If no standard bookmarks were passed in, then call function to 
		// get and move current position to bookmark position
		default:

         
            TRACE("\nIn move from bookmark with offset %d in %s\n",dlOffset.LowPart,(pThis->is_clone ? "Clone" : "Original"));

				{				
				VARIANT var;
            VariantInit(&var);
				var.vt=VT_ARRAY|VT_UI1;
				//BSTR b=::SysAllocStringLen((BSTR)((char*)pBookmark+1),cbBookmark-1);
				//::VectorFromBstr((BSTR)((char*)pBookmark+1),&var.parray);
				//::SysFreeString(b);	
            void *ppvData;
            SAFEARRAYBOUND rgsabound[1];	
            rgsabound[0].lLbound = 0;
	         rgsabound[0].cElements = cbBookmark-1;            
            var.parray=SafeArrayCreate(VT_UI1,1,rgsabound);
            SafeArrayAccessData(var.parray,&ppvData);
            memcpy(ppvData,((char*)pBookmark)+1,cbBookmark-1),
            SafeArrayUnaccessData(var.parray);
            COleVariant pas(var);
 				try{
				   pThis->m_pRecordSet->SetBookmark(pas);
				}catch(CDaoException* e){
					e->Delete();
               ::SafeArrayDestroy(var.parray);
				   pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
				   VariantClear((VARIANT *)&rgReasons[0].arg1);
				   VariantClear((VARIANT *)&rgReasons[0].arg2);
               
               TRACE("\nBad bookmark\n");

               if(!dlOffset.LowPart)return DB_E_ROWDELETED;
               else return DB_E_BADBOOKMARK;
				}
				::SafeArrayDestroy(var.parray);

            try{
					pThis->m_pRecordSet->Move(dlOffset.LowPart);
				}catch(CDaoException *e){
					e->Delete();
				}
			}
			break;
	}
  
	// Adjust new row position if before first row or after last row

	// If they also requested rows, then get the rows.
	if(pFetchParams && pFetchParams->cRowsRequested){
		hr=pThis->FetchRows(pFetchParams);
	// If we didn't fetch rows, then check for BOF or EOF and return DB_S_ENDOFCURSOR
	}else if(pThis->m_pRecordSet->IsBOF() || pThis->m_pRecordSet->IsEOF())hr = DB_S_ENDOFCURSOR;
	else hr = S_OK;


	// If the event failed in any way, then call FailedToDo
	if (FAILED(hr)){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}

	
	// Don't notify if only retrieved the current row
	if (*(BYTE *)pBookmark != DBBMK_CURRENT || dlOffset.LowPart ||
		(pFetchParams && pFetchParams->cRowsRequested>1)){

		SafeArrayDestroy(psa);
      pThis->MakeSafearrayFromBok(&psa);
		rgReasons[0].arg1.parray = psa;
		rgReasons[0].arg2.lVal = (long)0;
		TRACE("\nNotify after move!\n");
		pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	}


	
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	return hr;
}

// You will need to re-implement this function according to the Bookmark scheme that you use.
HRESULT CRowCursor::XCURSORMOVE::GetBookmark(DBCOLUMNID *pBookmarkType, ULONG cbMaxSize, ULONG *pcbBookmark, void *pBookmark)
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	
	if(!pBookmark || !pBookmarkType || !pcbBookmark)return DB_E_BUFFERTOOSMALL;
   if(!pThis->IsOpen())return E_FAIL;
	// Only bookmarks of type COLUMN_BMKTEMPORARY are supported in this sample. So check
	// to see if the requested bookmark is the same type.
	if(memcmp((void *)&COLUMN_BMKTEMPORARY, pBookmarkType, sizeof(DBCOLUMNID))==0){
			if(pThis->m_pRecordSet->IsEOF()){ //Return End
				*pcbBookmark = DB_BMK_SIZE;
				if(cbMaxSize<DB_BMK_SIZE)return DB_E_BUFFERTOOSMALL;
				memcpy(pBookmark, &DBBMK_END, DB_BMK_SIZE);
			}else if(pThis->m_pRecordSet->IsBOF()){ //Return Beginning
				*pcbBookmark = DB_BMK_SIZE;
				if(cbMaxSize<DB_BMK_SIZE)return DB_E_BUFFERTOOSMALL;
				memcpy(pBookmark, &DBBMK_BEGINNING, DB_BMK_SIZE);
			}else{
				// Calculate the new bookmark based on our handy formula
				COleVariant bok;
				try{
					bok=pThis->m_pRecordSet->GetBookmark();
				}catch(CDaoException *e){
					e->Delete();
					*pcbBookmark = DB_BMK_SIZE;
					if(cbMaxSize<DB_BMK_SIZE)return DB_E_BUFFERTOOSMALL;
//					if(pThis->m_pRecordSet->IsDeleted()){
//						memcpy(pBookmark, &DBBMK_CURRENT, DB_BMK_SIZE);						
//						return S_OK;
//					}else
					memcpy(pBookmark, &DBBMK_INVALID, DB_BMK_SIZE);
					return E_FAIL;
				}
				//BSTR bstr;
				//BstrFromVector(((LPCVARIANT)bok)->parray,(unsigned short**)&bstr);
            void *ppvData;
            SafeArrayAccessData(((LPCVARIANT)bok)->parray,&ppvData);
            long int ub,lb;
            SafeArrayGetUBound(((LPCVARIANT)bok)->parray,1,&ub);
            SafeArrayGetLBound(((LPCVARIANT)bok)->parray,1,&lb);
            int len=ub-lb+1;
            *pcbBookmark = 1 + len;

				//*pcbBookmark = 1+(lstrlenW(bstr)+1)*2;
				if(pBookmark && cbMaxSize<*pcbBookmark){
					//::SysFreeString(bstr);
               SafeArrayUnaccessData(((LPCVARIANT)bok)->parray);
					return DB_E_BUFFERTOOSMALL;
				}
				*(char*)pBookmark='a';
				memcpy(((char*)pBookmark)+1,(char*)ppvData,*pcbBookmark-1);
            SafeArrayUnaccessData(((LPCVARIANT)bok)->parray);
				//::SysFreeString(bstr);
			}
	}else return DB_E_BADCOLUMNID;	
	return S_OK;
}	

HRESULT CRowCursor::XCURSORMOVE::Clone(DWORD dwFlags, REFIID riid, IUnknown **ppvClonedCursor)
{

	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	if (dwFlags != DBCLONEOPTS_DEFAULT && dwFlags != DBCLONEOPTS_SAMEROW) return E_INVALIDARG;
	if (dwFlags == DBCLONEOPTS_SAMEROW){
		// Return a reference to the requested Interface using the 
		// original Row Cursor
		pThis->m_xCURSOR.QueryInterface(riid, (void **)ppvClonedCursor);   
   }else{
		// create a new row cursor object so that it may have an independent current row
	CRowCursor *pRowCursor = new CRowCursor(pThis->m_pControl,FALSE);
  	if (!pRowCursor)return E_OUTOFMEMORY;
   try{
         pRowCursor->m_pRecordSet=pThis->m_pRecordSet->Clone();
	}catch(CDaoException *e){
			e->Delete();
			delete pRowCursor->m_pRecordSet;
			pRowCursor->m_pRecordSet=NULL;
			delete pRowCursor;
			return E_OUTOFMEMORY;
	}
	if(!pRowCursor->m_pRecordSet->IsOpen()){
			delete pRowCursor->m_pRecordSet;
			pRowCursor->m_pRecordSet=NULL;
			delete pRowCursor;			
			return E_OUTOFMEMORY;
	}

	pRowCursor->m_xCURSOR.QueryInterface(riid, (void **)ppvClonedCursor);
   pRowCursor->ExternalRelease();
				
	// copy the original row cursor object's members into the clone's
	pRowCursor->m_lRowDataLength = pThis->m_lRowDataLength;
	pRowCursor->m_nExistingDataCols = pThis->m_nExistingDataCols;
	pRowCursor->is_clone=TRUE;
	pThis->clone=pRowCursor;
   pRowCursor->parent=pThis;
	ULONG bAllocSize = sizeof(DBCOLUMNBINDING) * pThis->m_nExistingDataCols;
		
	if (bAllocSize){
		pRowCursor->m_pDataColumns = (DBCOLUMNBINDING *)CoTaskMemAlloc(bAllocSize);
		
		if (!pRowCursor->m_pDataColumns)	return E_OUTOFMEMORY;

		// copy the individual column bindings also.
		// NOTE: Only COLUMNIDs of type DBCOLKIND_GUID_NUMBER are supported.in this sample
		for (ULONG i = 0; i < pThis->m_nExistingDataCols; i++)
			pRowCursor->m_pDataColumns[i] = pThis->m_pDataColumns[i];
		}else	pRowCursor->m_pDataColumns = NULL;
	}
	return S_OK;
}

HRESULT CRowCursor::XCURSORMOVE::QueryInterface(REFIID riid, LPVOID FAR *ppvObj)
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	return (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
}

ULONG CRowCursor::XCURSORMOVE::AddRef()
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	return (ULONG)pThis->ExternalAddRef();
}

ULONG CRowCursor::XCURSORMOVE::Release()
{
	METHOD_PROLOGUE(CRowCursor, CURSORMOVE)
	return (ULONG)pThis->ExternalRelease();
}

/////////////////////////////////////////////////////////////////////////
//
//	ICursorUpdateARow implementation
//
/////////////////////////////////////////////////////////////////////////
    
STDMETHODIMP_(ULONG) CRowCursor::XCursorUpdateARow::AddRef()
{
    METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
	 if(pThis->m_pControl->m_readOnly)return (unsigned long)E_NOTIMPL;
    return pThis->ExternalAddRef();
}
 
STDMETHODIMP_(ULONG) CRowCursor::XCursorUpdateARow::Release()
{
    METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
	 if(pThis->m_pControl->m_readOnly)return (unsigned long)E_NOTIMPL;
    return pThis->ExternalRelease();
}
    
STDMETHODIMP CRowCursor::XCursorUpdateARow::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
	if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	return pThis->ExternalQueryInterface(&iid, ppvObj);
}    
    
STDMETHODIMP CRowCursor::XCursorUpdateARow::BeginUpdate(DWORD dwFlags)
{
	METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
   if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	if(pThis->m_bUpdateInProgress)return DB_E_UPDATEINPROGRESS;


   
	HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;

	
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);
	memset(&rgReasons[0].arg1, 0, sizeof(DBVARIANT));
	memset(&rgReasons[0].arg2, 0, sizeof(DBVARIANT));
	
	// Select notification options
	if (dwFlags == DBROWACTION_ADD){
		cReasons = 1;
		pThis->m_dwEditMode = DBEDITMODE_ADD;
		rgReasons[0].dwReason = DBREASON_ADDNEW;
		dwEventWhat = DBEVENT_NONCURRENT_ROW_DATA_CHANGED | DBEVENT_CURRENT_ROW_CHANGED | DBEVENT_CURRENT_ROW_DATA_CHANGED | DBEVENT_SET_OF_ROWS_CHANGED;
      //DBEVENT_CURRENT_ROW_DATA_CHANGED | DBEVENT_SET_OF_ROWS_CHANGED;
//		pThis->MakeSafearrayFromBuf(&psa, (void*)&DBBMK_CURRENT, DB_BMK_SIZE);
      pThis->MakeSafearrayFromBok(&psa);
		rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
		rgReasons[0].arg1.parray = psa;
		rgReasons[0].arg2.vt = VT_I4;
		rgReasons[0].arg2.lVal = 0;
	}else if (dwFlags == DBROWACTION_UPDATE){
		cReasons = 1;
		pThis->m_dwEditMode = DBEDITMODE_UPDATE;
		rgReasons[0].dwReason = DBREASON_EDIT;
		dwEventWhat = DBEVENT_CURRENT_ROW_DATA_CHANGED;
//      pThis->MakeSafearrayFromBok(&psa);
//		rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
//		rgReasons[0].arg1.parray = psa;
//		rgReasons[0].arg2.vt = VT_I4;
//		rgReasons[0].arg2.lVal = 0;
	}
	

	// Call notifications
	if (S_OK!=(hr=pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons)))return hr;
	
	if (dwFlags!=DBROWACTION_ADD && dwFlags!=DBROWACTION_UPDATE){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return E_INVALIDARG;
	}

	if (pThis->m_bUpdateInProgress){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return DB_E_UPDATEINPROGRESS;
	}

	pThis->m_bUpdateInProgress = TRUE;
 	ULONG cbRowSize = pThis->m_pControl->m_cbRowSize;

	if (dwFlags == DBROWACTION_ADD){
//      try{
//			pThis->m_pRecordSet->MoveLast();
//			pThis->m_pRecordSet->MoveNext();
//      }catch(CDaoException *e){
//         e->Delete();
//      }
      try{
			pThis->m_pRecordSet->AddNew();
		}catch(CDaoException *e){
			e->Delete();
			pThis->m_bUpdateInProgress = FALSE;
			pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			VariantClear((VARIANT *)&rgReasons[0].arg1);
			VariantClear((VARIANT *)&rgReasons[0].arg2);
			return DB_S_CANCEL;
		}
	}else if (dwFlags == DBROWACTION_UPDATE){
		try{
			pThis->m_pRecordSet->Edit();		
		}catch(CDaoException *e){
			e->Delete();
			pThis->m_bUpdateInProgress = FALSE;
			pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			VariantClear((VARIANT *)&rgReasons[0].arg1);
			VariantClear((VARIANT *)&rgReasons[0].arg2);
			return DB_S_CANCEL;
		}
	}
	
	// Call after notifcations
	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
   return S_OK;
}

STDMETHODIMP CRowCursor::XCursorUpdateARow::SetColumn(DBCOLUMNID *pcid, DBBINDPARAMS *pBindParams)
{
	METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)

   
   if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);

	if(!pcid || !pBindParams)return DB_E_BUFFERTOOSMALL;

	cReasons = 1;
	rgReasons[0].dwReason = DBREASON_SETCOLUMN;
	dwEventWhat = DBEVENT_CURRENT_ROW_DATA_CHANGED;
	
	if (S_OK != (hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons)))return hr;
	if (!pThis->m_bUpdateInProgress){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		return DB_E_STATEERROR;
	}
		
	USES_CONVERSION;
	LONG cCol;
	if(pcid!=NULL)cCol = pcid->lNumber;
	else cCol = pThis->m_pDataColumns[0].columnID.lNumber;	//first column binding
	
	BSTR *pbstrClient = NULL;
	LPSTR *plpstrClient = NULL;
	LPWSTR *plpwstrClient = NULL;
	BLOB *pBlob = NULL;
	DATE dtVal=0;
	CURRENCY cyVal;
	cyVal.Lo=cyVal.Hi=0;
	double dVal=0;
	long lVal = 0;
	DWORD dwDataType = 0;

	if (pBindParams->dwDataType == DBTYPE_ANYVARIANT && pBindParams->dwBinding == DBBINDING_VARIANT)
		dwDataType = ((DBVARIANT *)(pBindParams->pData))->vt;
	else
		dwDataType = pBindParams->dwDataType;

	int lng=pThis->m_pControl->m_pColOffsetInRow[cCol+1]-pThis->m_pControl->m_pColOffsetInRow[cCol];
	BYTE *pDestColValue=(unsigned char*)CoTaskMemAlloc(lng);
	if(!pDestColValue)return E_OUTOFMEMORY;
	memset(pDestColValue,0,lng);
		
	if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_LPSTR){
		switch (dwDataType){
			case VT_EMPTY:
			case VT_NULL:
				*(LPSTR *)pDestColValue=(LPSTR)CoTaskMemAlloc(1);
				(*(LPSTR *)pDestColValue)[0]=0;
				break;
         case VT_PTR:
			case VT_BSTR:
				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)
                  pbstrClient = &(((DBVARIANT *)(pBindParams->pData))->bstrVal);
					else pbstrClient=NULL;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)pbstrClient = (BSTR *)(pBindParams->pData);
				else pbstrClient=NULL;
				
				if(pbstrClient && *pbstrClient){
					*(LPSTR *)pDestColValue	= (LPSTR)CoTaskMemAlloc(SysStringLen(*(BSTR*)pbstrClient) + 1);
					if (*(LPSTR *)pDestColValue == NULL){
						hr = E_OUTOFMEMORY;
						break;
					}
					lstrcpy(*(LPSTR *)pDestColValue, OLE2T(*(BSTR*)pbstrClient));
				}else
					*(LPSTR *)pDestColValue	= NULL;
				break;

			case VT_LPSTR:
				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)
                  plpstrClient = &(((DBVARIANT *)(pBindParams->pData))->pszVal);
					else plpstrClient=NULL;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)plpstrClient = (LPSTR *)(pBindParams->pData);
				else plpstrClient=NULL;

				if (plpstrClient && *plpstrClient){
					*(LPSTR *)pDestColValue	= (LPSTR)CoTaskMemAlloc(strlen(*plpstrClient) + 1);
					if (*(LPSTR *)pDestColValue == NULL){
						hr = E_OUTOFMEMORY;
						break;
					}
					lstrcpy(*(LPSTR *)pDestColValue, *plpstrClient);
				}else	*(LPSTR *)pDestColValue	= NULL;
				break;

			case VT_LPWSTR:

				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)plpwstrClient = &(((DBVARIANT *)(pBindParams->pData))->pwszVal);
					else plpwstrClient=NULL;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)plpwstrClient = (LPWSTR *)(pBindParams->pData);
				else plpwstrClient=NULL;

				if (plpwstrClient && *plpwstrClient){
					*(LPSTR *)pDestColValue	= (LPSTR)CoTaskMemAlloc(SysStringLen(*plpwstrClient) + 1);

					if (*(LPSTR *)pDestColValue == NULL){
						hr = E_OUTOFMEMORY;
						break;
					}
					lstrcpy(*(LPSTR *)pDestColValue, OLE2T(*(LPWSTR*)plpwstrClient));
				}else	*(LPSTR *)pDestColValue	= NULL;

				break;

			
			
			default :	
				
				hr = DB_E_SCHEMAVIOLATION;
				break;
		}
	}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I4 
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI4 
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_INT
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UINT
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I2 
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I1
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI1
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI2 
		|| pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R4 
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R8
		|| pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_DATE 
      || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_CY){
		switch (dwDataType){
         case VT_NULL:
			case VT_EMPTY:
				
				if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I4 || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I2 || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI1){
					*(long *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R4 || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R8){
					*(double *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_DATE){
					COleDateTime dt;
					dtVal=(DATE)dt;
					*(DATE *)pDestColValue = dtVal;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_CY){
					COleCurrency dt;
					cyVal=(CURRENCY)dt;
					*(CURRENCY *)pDestColValue = cyVal;
				}
				
				break;
         
         case VT_PTR:
         case VT_BSTR:

				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)pbstrClient = &(((DBVARIANT *)(pBindParams->pData))->bstrVal);
					else pbstrClient=NULL;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)pbstrClient = (BSTR *)(pBindParams->pData);
				else pbstrClient=NULL;
						
				
				if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I4 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I2 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI2 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI4
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I1
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_INT
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UINT 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI1){
					if (pbstrClient && *pbstrClient)*(long *)pDestColValue = atol(OLE2T(*pbstrClient));
					else *(long *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R4 || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R8){
					if (pbstrClient && *pbstrClient)*(double *)pDestColValue = atof(OLE2T(*pbstrClient));
					else *(double *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_DATE){
					COleDateTime dt;
					BOOL res=TRUE;
					if (pbstrClient && *pbstrClient){
						res=dt.ParseDateTime(OLE2T(*pbstrClient));
					}
					dtVal=(DATE)dt;
					if(res)*(DATE *)pDestColValue = dtVal;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_CY){
					COleCurrency dt;
					BOOL res=TRUE;
					if (pbstrClient && *pbstrClient){
						res=dt.ParseCurrency(OLE2T(*pbstrClient));
					}
					cyVal=(CURRENCY)dt;
					if(res)*(CURRENCY *)pDestColValue = cyVal;
				}
				break;

			case VT_LPSTR:

				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)plpstrClient = &(((DBVARIANT *)(pBindParams->pData))->pszVal);
					else plpstrClient=NULL;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)plpstrClient = (LPSTR *)(pBindParams->pData);
				else plpstrClient=NULL;
					
				if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I4 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I2 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI2 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI4
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I1                
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_INT
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UINT 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI1){
					if (plpstrClient && *plpstrClient)*(long *)pDestColValue = atol(*plpstrClient);
					else *(long *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R4 || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R8){
					if (plpstrClient && *plpstrClient)*(double *)pDestColValue = atof(*plpstrClient);
					else *(double *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_DATE){
					COleDateTime dt;
					BOOL res=TRUE;
					if (plpstrClient && *plpstrClient){
						res=dt.ParseDateTime(*plpstrClient);
					}
					dtVal=(DATE)dt;
					if(res)*(DATE *)pDestColValue = dtVal;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_CY){
					COleCurrency dt;
					BOOL res=TRUE;
					if (plpstrClient && *plpstrClient){
						res=dt.ParseCurrency(*plpstrClient);
					}
					cyVal=(CURRENCY)dt;
					if(res)*(CURRENCY *)pDestColValue = cyVal;
				}
			

				
				
				
				break;

			case VT_LPWSTR:

				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)plpwstrClient = &(((DBVARIANT *)(pBindParams->pData))->pwszVal);
					else plpwstrClient=NULL;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)plpwstrClient = (LPWSTR *)(pBindParams->pData);
				else plpwstrClient=NULL;

				if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I4 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI4 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_INT
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UINT 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I2 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI2 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_I1 
               || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_UI1){
					if (plpwstrClient && *plpwstrClient)*(long *)pDestColValue = atol(OLE2T(*plpwstrClient));					
					else *(long *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R4 || pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_R8){
					if (plpwstrClient && *plpwstrClient)*(double *)pDestColValue = atof(OLE2T(*plpwstrClient));
					else *(double *)pDestColValue = 0;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_DATE){
					COleDateTime dt;
					BOOL res=TRUE;
					if (plpwstrClient && *plpwstrClient){
						res=dt.ParseDateTime(OLE2T(*plpwstrClient));
					}
					dtVal=(DATE)dt;
					if(res)*(DATE *)pDestColValue = dtVal;
				}else if(pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_CY){
					COleCurrency dt;
					BOOL res=TRUE;
					if (plpwstrClient && *plpwstrClient){
						res=dt.ParseCurrency(OLE2T(*plpwstrClient));
					}
					cyVal=(CURRENCY)dt;
					if(res)*(CURRENCY *)pDestColValue = cyVal;
				}
			
				break;

         case VT_INT:
         case VT_UINT:
			case VT_UI4:
			case VT_I4:
			case VT_UI2:
			case VT_I2:
			case VT_UI1:
				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)lVal = ((DBVARIANT *)(pBindParams->pData))->lVal;
					else lVal=0;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)lVal = (long)(pBindParams->pData);
				else lVal=0;

				*(long *)pDestColValue = lVal;			
				break;
			case VT_R4:
			case VT_R8:
				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)dVal = ((DBVARIANT *)(pBindParams->pData))->dblVal;
					else dVal=0;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)memcpy(&dVal,pBindParams->pData,sizeof(double));
				else dVal=0;

				*(double *)pDestColValue = (double)dVal;
				break;

			case VT_DATE:
				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)dtVal = ((DBVARIANT *)(pBindParams->pData))->date;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)memcpy(&dtVal,(pBindParams->pData),sizeof(DATE));

					*(DATE *)pDestColValue = (DATE)dtVal;					
			break;
			
			case VT_CY:
				if (pBindParams->dwBinding == DBBINDING_VARIANT){
					if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)cyVal = ((DBVARIANT *)(pBindParams->pData))->cyVal;
				}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)memcpy(&cyVal,(pBindParams->pData),sizeof(CURRENCY));

				*(CURRENCY *)pDestColValue = cyVal;
				break;
			default :

				hr = DB_E_SCHEMAVIOLATION;
				break;
		}
	}else if (pThis->m_pControl->m_pMetaData[cCol].dwColType == VT_BLOB){

		if (dwDataType == VT_BLOB)	{
			if (pBindParams->dwBinding == DBBINDING_VARIANT){
				if(/*pBindParams->cbVarDataLen &&*/ pBindParams->pData)pBlob = &(((DBVARIANT *)(pBindParams->pData))->blob);
				else pBlob=NULL;
			}else if (pBindParams->dwBinding == DBBINDING_DEFAULT)
				pBlob = (BLOB *)(pBindParams->pData);
			else pBlob=NULL;

			((BLOB *)pDestColValue)->cbSize = (pBlob ? pBlob->cbSize : 0);
			((BLOB *)pDestColValue)->pBlobData = (BYTE *)CoTaskMemAlloc((pBlob ? pBlob->cbSize : 0));

			if (((BLOB *)pDestColValue)->pBlobData == NULL)
				hr = E_OUTOFMEMORY;
			else
				memcpy((LPVOID)(((BLOB *)pDestColValue)->pBlobData), (LPVOID)pBlob->pBlobData, pBlob->cbSize);
		}else
			hr = DB_E_SCHEMAVIOLATION;
	}else
		hr = DB_E_SCHEMAVIOLATION;
	
	VARIANT var;
   VariantInit(&var);
	var.vt=(unsigned short)pThis->m_pControl->m_pMetaData[cCol].dwColType;
	
	switch(pThis->m_pControl->m_pMetaData[cCol].dwColType){
		case VT_BOOL:
			var.boolVal=*(unsigned char *)pDestColValue;
			break;
		case VT_I1:
		case VT_UI1:
			var.bVal=*(unsigned char *)pDestColValue;
			break;
		case VT_UI2:
		case VT_I2:
			var.iVal=*(unsigned short *)pDestColValue;
			break;
		case VT_INT:
      case VT_UINT:
      case VT_UI4:
		case VT_I4:
			var.lVal=*(unsigned long *)pDestColValue;
			break;
		case VT_CY:
			var.cyVal=*(CURRENCY *)pDestColValue;
			break;
		case VT_R4:
			var.fltVal=*(float *)pDestColValue;
			break;
		case VT_R8:
			var.dblVal=*(double *)pDestColValue;
			break;
		case VT_DATE:
			var.date=*(DATE *)pDestColValue;
			break;
      case VT_PTR:
      case VT_BSTR:
		case VT_LPWSTR:
		case VT_LPSTR:
			break;
		case VT_BLOB:
			break;
		case DBTYPE_BYTES:
			break;
	}
	
	COleVariant value;
	if(pThis->m_pControl->m_pMetaData[cCol].dwColType==VT_LPSTR || pThis->m_pControl->m_pMetaData[cCol].dwColType==VT_PTR){
		value=COleVariant(*(LPSTR*)pDestColValue,VT_BSTRT);
		CoTaskMemFree(*(LPSTR*)pDestColValue);
	}else value=COleVariant(var);

	VariantClear(&var);
	
	CoTaskMemFree(pDestColValue);
	if (FAILED(hr)){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		return hr;
	}

	try{
		pThis->m_pRecordSet->SetFieldValue(cCol,value);
	}catch(CDaoException* e){
		e->Delete();
	}

	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	return S_OK;
}

STDMETHODIMP CRowCursor::XCursorUpdateARow::GetColumn(DBCOLUMNID *pcid, DBBINDPARAMS *pBindParams, DWORD *pdwFlags)
{
   METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
	 if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;

	LONG cCol = pcid->lNumber;
	COleVariant var;

	try{
		if(cCol<pThis->m_pRecordSet->GetFieldCount())
			pThis->m_pRecordSet->GetFieldValue(cCol,var);
		else{ 			
			return E_FAIL;
		}
	}catch(CDaoException *e){
		e->Delete();
		var=COleVariant();
	}

	BYTE *pSrcColVal = &((LPVARIANT)var)->bVal; 
	BYTE *pDest = (BYTE *)pBindParams->pData;

	VARIANT *vnt = (VARIANT *)CoTaskMemAlloc(sizeof(VARIANT));
	VariantInit(vnt);
   
	// Simulate Out of Line Memory;
	if(pBindParams->dwBinding == DBBINDING_VARIANT){

		*(ULONG *)pDest = (ULONG)(ULONG *)(pDest + 4);
		
		// Skip the size of a long address
		pDest += 4;
	}

	
	VariantInit((VARIANT *)pDest);
	((DBVARIANT *)pDest)->vt = VT_BYREF|VT_VARIANT;
	((DBVARIANT *)pDest)->pvarVal = vnt;

	USES_CONVERSION;

	vnt->vt = (VARTYPE)pThis->m_pControl->m_pMetaData[cCol].dwColType;

	switch (vnt->vt){
      case VT_INT:
      case VT_UINT:
		case VT_I4:
      case VT_UI4:   vnt->lVal = *(long *)pSrcColVal;
						break;

      case VT_I1:
		case VT_UI1:	vnt->bVal = *(unsigned char *)pSrcColVal;
						break;

      case VT_UI2:
		case VT_I2:		vnt->iVal = *(short *)pSrcColVal;
						break;

		case VT_R4:		vnt->fltVal = *(float *)pSrcColVal;
						break;

		case VT_R8:		vnt->dblVal = *(double *)pSrcColVal;
						break;

		case VT_BOOL:	vnt->boolVal = *(VARIANT_BOOL *)pSrcColVal;
						break;

		case VT_CY:		vnt->cyVal = *(CY *)pSrcColVal;
						break;

		case VT_DATE:	vnt->date = *(DATE *)pSrcColVal;
						break;

      case VT_PTR:
      case VT_BSTR:
      case VT_LPSTR:	vnt->bstrVal = SysAllocString(A2W(*(LPSTR *)pSrcColVal));
						vnt->vt = VT_BSTR;
						break;
			
		case VT_LPWSTR:	vnt->bstrVal = SysAllocString(*(LPWSTR *)pSrcColVal);
						vnt->vt = VT_BSTR;
						break;

		// can't coerce any other data type to a VARIANT!
		default:		
						((DBVARIANT *)pDest)->vt = DBTYPE_HRESULT;
						((DBVARIANT *)pDest)->scode = DB_E_CANTCOERCE;

						return E_FAIL;
	}

	if (pdwFlags){
		try{
		if(pThis->m_pRecordSet->IsFieldDirty(NULL))
			*pdwFlags = DBCOLUMNDATA_CHANGED;
		else
			*pdwFlags = DBCOLUMNDATA_UNCHANGED;
		}catch(CDaoException *e){
			e->Delete();
		}
		if(pThis->m_pRecordSet->IsEOF() || pThis->m_pRecordSet->IsBOF())*pdwFlags = DBCOLUMNDATA_UNCHANGED;
	}

	return S_OK;
}

STDMETHODIMP CRowCursor::XCursorUpdateARow::GetEditMode(DWORD *pdwState)
{
   METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
   if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	if (!pdwState)return E_INVALIDARG;
   if(!pThis->IsOpen()) *pdwState = DBEDITMODE_NONE;
	else if (!pThis->m_bUpdateInProgress) *pdwState = DBEDITMODE_NONE;
	else *pdwState = pThis->m_dwEditMode;
	return S_OK;
}

STDMETHODIMP CRowCursor::XCursorUpdateARow::Update(DBCOLUMNID *pBookmarkType, ULONG *pcbBookmark, void **ppBookmark)
{
   METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
   if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	ULONG OldCurrentRow = 0;
	if(pBookmarkType && memcmp((void *)&COLUMN_BMKTEMPORARY, pBookmarkType, sizeof(DBCOLUMNID)))return DB_E_BADCOLUMNID;
	
   HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);

   pThis->MakeSafearrayFromBok(&psa);
   //pThis->MakeSafearrayFromBuf(&psa,(void*)&DBBMK_CURRENT,DB_BMK_SIZE);

   rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
	rgReasons[0].arg1.parray = psa;
	rgReasons[0].arg2.vt = VT_I4;
	rgReasons[0].arg2.lVal = 0;
	
	// Use these notification flags for adding data
	if (pThis->m_dwEditMode == DBEDITMODE_ADD){
		cReasons = 1;
		rgReasons[0].dwReason = DBREASON_INSERTED;
		dwEventWhat = DBEVENT_CURRENT_ROW_CHANGED | DBEVENT_CURRENT_ROW_DATA_CHANGED | 
			DBEVENT_NONCURRENT_ROW_DATA_CHANGED | DBEVENT_SET_OF_ROWS_CHANGED | 
			DBEVENT_ORDER_OF_ROWS_CHANGED;
	}else if (pThis->m_dwEditMode == DBEDITMODE_UPDATE){
		cReasons = 1;
	   rgReasons[0].dwReason = DBREASON_MODIFIED;
		dwEventWhat = /*DBEVENT_CURRENT_ROW_CHANGED |*/ DBEVENT_CURRENT_ROW_DATA_CHANGED | 
			DBEVENT_ORDER_OF_ROWS_CHANGED | DBEVENT_SET_OF_ROWS_CHANGED;
	}
	
	TRACE("\n In update in %s\n",(pThis->is_clone ? "Clone" : "Original"));

   hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons);

	if (FAILED(hr))	{
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		return hr;
	}

	if (!pThis->m_bUpdateInProgress)	{
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		return DB_E_STATEERROR;
	}
	BOOL was_add=FALSE;
	
	try{
	   pThis->m_pRecordSet->Update();
		pThis->m_bUpdateInProgress=FALSE;
		if(pThis->m_dwEditMode == DBEDITMODE_ADD)was_add=TRUE;
		pThis->m_dwEditMode = DBEDITMODE_NONE;      
      if(was_add){ 
			COleVariant bok;
			bok=pThis->m_pRecordSet->GetLastModifiedBookmark();
			pThis->m_pRecordSet->SetBookmark(bok);
		}
	}catch(CDaoException*e){
		e->Delete();
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		return hr;
	}
	

	
	
	
	
	
	if (pBookmarkType != NULL && pcbBookmark != NULL && ppBookmark != NULL)	{
			if(pThis->m_pRecordSet->IsEOF()){ //Return End
				*pcbBookmark = DB_BMK_SIZE;
				*ppBookmark=CoTaskMemAlloc(DB_BMK_SIZE);
				memcpy(*ppBookmark,(void*)&DBBMK_END,DB_BMK_SIZE);
			}else if(pThis->m_pRecordSet->IsBOF()){ //Return Beginning
				*pcbBookmark = DB_BMK_SIZE;
				*ppBookmark=CoTaskMemAlloc(DB_BMK_SIZE);
				memcpy(*ppBookmark,(void*)&DBBMK_BEGINNING,DB_BMK_SIZE);
			}else{
				// Calculate the new bookmark based on our handy formula
				COleVariant bok;
				try{
					bok=pThis->m_pRecordSet->GetBookmark();
					//BSTR bstr;
					//BstrFromVector(((LPCVARIANT)bok)->parray,(unsigned short**)&bstr);
            void *ppvData;
            SafeArrayAccessData(((LPCVARIANT)bok)->parray,&ppvData);
            long int ub,lb;
            SafeArrayGetUBound(((LPCVARIANT)bok)->parray,1,&ub);
            SafeArrayGetLBound(((LPCVARIANT)bok)->parray,1,&lb);
            long int len=ub-lb+1;
            *pcbBookmark = 1 + len;

					*ppBookmark=CoTaskMemAlloc(*pcbBookmark);
					*(char*)*ppBookmark='a';
					//*pcbBookmark = 1+(lstrlenW(bstr)+1)*2;
					memcpy(((char*)*ppBookmark)+1,(char*)ppvData,*pcbBookmark-1);
               SafeArrayUnaccessData(((LPCVARIANT)bok)->parray);
					//::SysFreeString(bstr);
				}catch(CDaoException *e){
					e->Delete();
					*pcbBookmark = DB_BMK_SIZE;
					*ppBookmark=CoTaskMemAlloc(DB_BMK_SIZE);
					//if(pThis->m_pRecordSet->IsDeleted()){
					//	memcpy(*ppBookmark,(void*)&DBBMK_CURRENT,DB_BMK_SIZE);
					//}else
						memcpy(*ppBookmark,(void*)&DBBMK_INVALID,DB_BMK_SIZE);
				}
       }				
	}
	

	if(was_add){
		SafeArrayDestroy(psa);
      pThis->MakeSafearrayFromBok(&psa);
	   rgReasons[0].arg1.parray = psa;
	   rgReasons[0].arg2.lVal = 0;
	}
	
	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	return S_OK;
}
	   
STDMETHODIMP CRowCursor::XCursorUpdateARow::Cancel()
{
   METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
   if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	if(!pThis->m_bUpdateInProgress)return DB_E_STATEERROR;		
	HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);
   
	cReasons = 1;
	rgReasons[0].dwReason = DBREASON_CANCELUPDATE;
	dwEventWhat = DBEVENT_CURRENT_ROW_DATA_CHANGED;
	if(S_OK != (hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons)))return hr;
	
	if (pThis->m_bUpdateInProgress){
		try{
			pThis->m_pRecordSet->CancelUpdate();
         pThis->m_dwEditMode = DBEDITMODE_NONE;
			pThis->m_bUpdateInProgress = FALSE;
		}catch(CDaoException *e){
			e->Delete();
			pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			return DB_E_STATEERROR;
		}
	}else{
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		return DB_E_STATEERROR;
	}

	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	return S_OK;
}

STDMETHODIMP CRowCursor::XCursorUpdateARow::Delete()
{
   METHOD_PROLOGUE(CRowCursor, CursorUpdateARow)
   if(pThis->m_pControl->m_readOnly)return E_NOTIMPL;
	if (pThis->m_bUpdateInProgress)return DB_E_UPDATEINPROGRESS;

	HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;
	// Initialize arguments for notifications
	
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);

	
	// Bookmarks should be passed in like this to allow VariantChangeByte
	// to be used by bound controls
	rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
   pThis->MakeSafearrayFromBok(&psa);
	
	rgReasons[0].arg1.parray = psa;
	rgReasons[0].arg2.vt = VT_NULL;
	dwEventWhat = DBEVENT_CURRENT_ROW_CHANGED;
	cReasons = 1;
	rgReasons[0].dwReason = DBREASON_DELETED;
	
	hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons);
	if (hr != S_OK){
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}	
	
	try{
		pThis->m_pRecordSet->Delete();
	}catch(CDaoException *e){
		e->Delete();
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return E_FAIL;
	}			
	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);

	return S_OK;
} 

////////////////////////////////////////////////////////////
// Helper function Implementations
//
// DoNotifyBefore and DoNotifyAfter are functions designed to take
// care of multicasting
//



HRESULT CRowCursor::DoNotifyBefore(DWORD dwEventWhat, ULONG cReasons, DBNOTIFYREASON rgReasons[]) 
{
	int i,j;
	HRESULT hr;
		
	const CPtrArray* pConnections = m_xConnectionPoint.GetConnections();
   if(!pConnections)return S_OK;

   int cConnections = pConnections->GetSize();
	INotifyDBEvents* pDBEventsSink;
	
	
	for (i = 0; i < cConnections; i++){
			pDBEventsSink = (INotifyDBEvents *)(pConnections->GetAt(i));
         ASSERT(pDBEventsSink != NULL);
		if (pDBEventsSink->OKToDo(dwEventWhat, cReasons, rgReasons) == S_FALSE){
			for(j=0;j<=i;j++)	pDBEventsSink->Cancelled(dwEventWhat, cReasons, rgReasons);
			return DB_S_OPERATIONCANCELLED;
		}
	}

	for (i = 0; i < cConnections; i++){
	    pDBEventsSink = (INotifyDBEvents *)(pConnections->GetAt(i));
        ASSERT(pDBEventsSink != NULL);
		  if FAILED(hr=pDBEventsSink->SyncBefore(dwEventWhat, cReasons, rgReasons)){
			for(j=0;j<cConnections;j++)pDBEventsSink->FailedToDo(dwEventWhat, cReasons, rgReasons);
			return hr;
		}
	}

	
	for (i = 0; i < cConnections; i++){
	    pDBEventsSink = (INotifyDBEvents *)(pConnections->GetAt(i));
        ASSERT(pDBEventsSink != NULL);
		if (pDBEventsSink->AboutToDo(dwEventWhat, cReasons, rgReasons) != S_OK){
			for(j=0;j<cConnections;j++)pDBEventsSink->FailedToDo(dwEventWhat, cReasons, rgReasons);
			return E_FAIL;
		}
	}

	return S_OK;
}




HRESULT CRowCursor::DoNotifyAfter(DWORD dwEventWhat, ULONG cReasons, DBNOTIFYREASON rgReasons[]) 
{
	int i;
	HRESULT hr;

	const CPtrArray* pConnections = m_xConnectionPoint.GetConnections();
	if(!pConnections)return S_OK;

   int cConnections = pConnections->GetSize();
	INotifyDBEvents* pDBEventsSink;
	
	for (i = 0; i < cConnections; i++){
		pDBEventsSink = (INotifyDBEvents *)(pConnections->GetAt(i));
        ASSERT(pDBEventsSink != NULL);
		  if FAILED(hr = pDBEventsSink->SyncAfter(dwEventWhat, cReasons, rgReasons)){
				return hr;
			}
	}	

	
	
	for (i = 0; i < cConnections; i++){
		pDBEventsSink = (INotifyDBEvents *)(pConnections->GetAt(i));
        ASSERT(pDBEventsSink != NULL);
		  if FAILED(hr = pDBEventsSink->DidEvent(dwEventWhat, cReasons, rgReasons)){
			return hr;
		}
	}	

	return S_OK;
}

// If an event fails then call this function. It handles multicasting
HRESULT CRowCursor::CallFailedToDo(DWORD dwEventWhat, ULONG cReasons, DBNOTIFYREASON rgReasons[]) 
{
	const CPtrArray* pConnections = m_xConnectionPoint.GetConnections();
   if(!pConnections)return S_OK;
   int cConnections = pConnections->GetSize();
   INotifyDBEvents* pDBEventsSink;
	for (int i=0; i<cConnections; i++) {
	   pDBEventsSink = (INotifyDBEvents *)(pConnections->GetAt(i));
      ASSERT(pDBEventsSink != NULL);
		pDBEventsSink->FailedToDo(dwEventWhat, cReasons, rgReasons);
	}
	return S_OK;
}



//=---------------------------------------------------------------------------=
// MakeSafearrayFromBuf
//=---------------------------------------------------------------------------=
// given a buffer, make a SAFEARRAY out of it.

HRESULT CRowCursor::MakeSafearrayFromBuf(SAFEARRAY FAR * FAR *ppsa, LPVOID p, ULONG cb)
{
    SAFEARRAY FAR *psaTmp;
    SAFEARRAYBOUND sab[1];
    VOID HUGEP *pData;
    HRESULT hr;
    sab[0].cElements = cb;
    sab[0].lLbound = 0L;
    psaTmp = SafeArrayCreate(VT_UI1, 1, sab);
    if (!psaTmp)return E_OUTOFMEMORY;
    hr = SafeArrayAccessData(psaTmp, &pData);
    if (FAILED(hr)){
        SafeArrayDestroy(psaTmp);
        *ppsa = NULL;
		return hr;
    }
    memcpy(pData, p, (size_t)cb);
    SafeArrayUnaccessData(psaTmp);
    *ppsa = psaTmp;
	 return NOERROR;
}

// This function is called by ICursor::GetNextRows() as well as ICursorMove::Move().
// It fetches the requested number of data rows according to the current data
// column bindings.

HRESULT CRowCursor::FetchRows(DBFETCHROWS *pFetchParams)
{
	COleVariant value;
	CDaoFieldInfo info;
	// loop counter for rows and columns
	ULONG i=0, j=0;
	// variable to keep track of starting position in out-of-line memory for each column.
	ULONG cbCumulativeOutOfLineMemSize = 0;
	// the ordinal position of the current column of data
	LONG nOrdinalFieldPos = 0;
	// variable to store return code
	HRESULT hr;
	// ptr to the beginning of a column of data within the current row of source data
	// in the memory that contains the resultset rows
	BYTE *pSrcColValue = NULL;
	// ptr to the beginning of each row of in-line memory
	BYTE *pDestBytePtr = NULL;
	// ptr to the beginning of a column of data within the current row of in-line memory
	BYTE *pDestColValue = NULL;
	// COLUMNID guids
	GUID gnumonly = GUID_NUMBERONLY;
	GUID bmkguid = DBBMKGUID;
   

   if (!pFetchParams)return E_FAIL;
   if(!pFetchParams->cRowsRequested)return S_OK;
	

	// No rows can be returned, because already at end of cursor!
	//test EOF cursor	
	if(m_pRecordSet->IsEOF() || m_pRecordSet->IsBOF())return DB_S_ENDOFCURSOR;
	
	pFetchParams->cRowsReturned=pFetchParams->cRowsRequested;
	
   TRACE("\nJust before fetch\n");
	
	if (pFetchParams->dwFlags & DBROWFETCH_CALLEEALLOCATES){
		// callee or provider allocates the memory
		if (m_lRowDataLength)
			pFetchParams->pData = CoTaskMemAlloc(m_lRowDataLength * pFetchParams->cRowsReturned);
		else{ // invalid in-line memory row size!
			pFetchParams->cRowsReturned = 0;
			return E_FAIL;
		}
		if (!pFetchParams->pData){
			pFetchParams->cRowsReturned = 0;
			return E_OUTOFMEMORY;
		}

      COleVariant start;
		if(pFetchParams->cRowsReturned>1){
			try{
				start=m_pRecordSet->GetBookmark();
			}catch(CDaoException *e){
				e->Delete();
				pFetchParams->cRowsReturned = 0;
				return DB_S_ENDOFCURSOR;		
			}
		}
		
		// determine the size of out-of-line memory needed
		for (i=0;/*!m_pRecordSet->IsEOF()*/;i++){
			for (j = 0; j < m_nExistingDataCols; j++){
				// hold the size of each column of data requiring out-of-line memory
				// Note: only BLOBs and strings of type VT_LPSTR and VT_LPWSTR require
				// out-of-line memory.
				ULONG cbVarMemLength = 0;
				if (m_pDataColumns[j].columnID.guid == gnumonly){
					// For regular data, COLUMNID # is treated as the ordinal position
					nOrdinalFieldPos = m_pDataColumns[j].columnID.lNumber;
					try{
						m_pRecordSet->GetFieldValue(nOrdinalFieldPos,value);
					}catch(CDaoException *e){
						e->Delete();
						break;
					}			
					//m_pControl->m_pRecordSet->GetFieldInfo(nOrdinalFieldPos,info);
				}else if (m_pDataColumns[j].columnID.guid == bmkguid){
					// add the total # of columns also for bookmark columns.
					nOrdinalFieldPos = m_pControl->m_nColumns + m_pDataColumns[j].columnID.lNumber;
					try{
						value=m_pRecordSet->GetBookmark();
					}catch(CDaoException *e){
						e->Delete();
						break;
					}
				}

				// calculate column data offset by indexing into the appropriate element of the
				// m_pColOffsetInRow array for each current row.				
				
				pSrcColValue=&((LPVARIANT)value)->bVal;
				if(((LPVARIANT)value)->vt==VT_NULL || ((LPVARIANT)value)->vt==VT_EMPTY)pSrcColValue=NULL;
				
				if (m_pDataColumns[j].dwDataType == VT_BLOB){
					//BSTR bstr;
					//BstrFromVector(((LPVARIANT)value)->parray,(unsigned short **)&bstr);
					//cbVarMemLength=3+lstrlenW((BSTR)bstr)*2;
					//::SysFreeString(bstr);
               //void **ppvData;
               //SafeArrayAccessData(((LPCVARIANT)value)->parray,ppvData);
               long int ub,lb;
               SafeArrayGetUBound(((LPCVARIANT)value)->parray,1,&ub);
               SafeArrayGetLBound(((LPCVARIANT)value)->parray,1,&lb);
               long int len=ub-lb+1;
               cbVarMemLength = 1 + len;
               //SafeArrayUnaccessData(((LPCVARIANT)value)->parray);
				}else if (m_pDataColumns[j].dwDataType == VT_LPSTR || m_pDataColumns[j].dwDataType == VT_LPWSTR){
					if (m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_LPSTR){
						if (pSrcColValue && *(LPSTR *)pSrcColValue != NULL)
							cbVarMemLength = strlen(*(LPSTR *)pSrcColValue) + 1;
						else cbVarMemLength = 1;
					}else if (m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_LPWSTR){
						if (pSrcColValue && *(LPWSTR *)pSrcColValue != NULL)
							cbVarMemLength = wcslen(*(LPWSTR *)pSrcColValue) + 1;
						else cbVarMemLength = 1;
					}else if (m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_DATE ||
								m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_CY ||
								m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_R4 ||
								m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_R8 ||
								m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_I4 ||
							   m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_I2 ||
							   m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_BOOL ||
                        m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_UI2 ||
                        m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_UI1 ||
                        m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_UI4 ||
                        m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_INT ||
                        m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_UINT)
						cbVarMemLength = 33;
				}

				if (m_pDataColumns[j].dwDataType == VT_LPWSTR)cbVarMemLength *= 2;
				cbCumulativeOutOfLineMemSize += cbVarMemLength;
			}
			
			try{
				if(i+1<pFetchParams->cRowsReturned)m_pRecordSet->MoveNext();
				else break;
				if(m_pRecordSet->IsEOF())break;
			}catch(CDaoException *e){
				e->Delete();
				break;
			}
		}

		if(pFetchParams->cRowsReturned>1){
			try{
				m_pRecordSet->SetBookmark(start);
			}catch(CDaoException *e){
				e->Delete();
				pFetchParams->cRowsReturned = 0;
				return DB_S_ENDOFCURSOR;		
			}
		}
      
      
      // Allocate it!
		pFetchParams->pVarData = CoTaskMemAlloc(cbCumulativeOutOfLineMemSize);
		cbCumulativeOutOfLineMemSize = 0;
		if (!pFetchParams->pVarData){  
			pFetchParams->cRowsReturned = 0;
			return E_OUTOFMEMORY;
		}
	}else if (pFetchParams->dwFlags != DBROWFETCH_DEFAULT && !(pFetchParams->dwFlags & DBROWFETCH_FORCEREFRESH)){
		pFetchParams->cRowsReturned = 0;
		return DB_E_BADFETCHINFO;
	}else{ // caller allocated memory
	// fail if caller did not allocate any in-line memory 
		if (!pFetchParams->pData){
			pFetchParams->cRowsReturned = 0;
			return DB_E_BADFETCHINFO;
		}
	}


   TRACE("\nStarting to fetch data!\n");

   // iterate through each bound column of each requested row to fetch the data
	ULONG really_fetch=0;
	
	for (i=0;/*!m_pRecordSet->IsEOF()*/;i++){
		pDestBytePtr = (BYTE *)pFetchParams->pData;
		pDestBytePtr += (i * m_lRowDataLength);


		for (j = 0;j < m_nExistingDataCols; j++)	{
			// proceed only if inline memory contains any data for the column j...
			if (m_pDataColumns[j].obData != DB_NOVALUE){
				// fill the column length field in in-line memory if requested by client
				if (m_pDataColumns[j].obVarDataLen != DB_NOVALUE){
					pDestColValue = pDestBytePtr + m_pDataColumns[j].obVarDataLen;
					*(DWORD *)pDestColValue = 0;
				}

				// fill the column info field in in-line memory if requested by client
				if (m_pDataColumns[j].obInfo != DB_NOVALUE){
					pDestColValue = pDestBytePtr + m_pDataColumns[j].obInfo;
					*(DWORD *)pDestColValue = DB_NOINFO;
				}

				// fill the actual in-line memory data...
				pDestColValue = pDestBytePtr + m_pDataColumns[j].obData;
			
				// ...for regular data columns of type DBCOLKIND_GUID_NUMBER only
				if (m_pDataColumns[j].columnID.guid == gnumonly){
					nOrdinalFieldPos = m_pDataColumns[j].columnID.lNumber;
					value.Clear();
					try{
						m_pRecordSet->GetFieldValue(nOrdinalFieldPos,value);
					}catch(CDaoException *e){
						e->Delete();
						break;
					}
					const VARIANT * var=(LPCVARIANT)value;
					
					__int64 i64=0;

					pSrcColValue=(unsigned char *)var->bstrVal;
					
					if(((LPVARIANT)value)->vt==VT_NULL || ((LPVARIANT)value)->vt==VT_EMPTY)
							pSrcColValue=NULL;

					// fetch the actual data according to the data type requested by the client
					// in the column bindings.

               
               switch (m_pDataColumns[j].dwDataType){
						case VT_PTR:
						case VT_BSTR:
						case VT_LPSTR:	
						case VT_LPWSTR:							
							switch(var->vt){
								case VT_BOOL:							
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->boolVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_UI1:
								case VT_I1:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->bVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_UI2:
								case VT_I2:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->iVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
                        case VT_INT:
                        case VT_UINT:
                        case VT_UI4:
                        case VT_I4:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->lVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_R4:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->fltVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_R8:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->dblVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_DATE:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->date, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_CY:
								hr = CvtToString(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->cyVal, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
								case VT_BSTR:
								default:
                           hr = CvtToString((m_pDataColumns[j].dwDataType==VT_PTR ? VT_BSTR : m_pDataColumns[j].dwDataType), pDestBytePtr, (pSrcColValue ? (unsigned char*)&pSrcColValue : (unsigned char*)&i64), nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
								break;
							}
							
							
							if (FAILED(hr) || hr == DB_S_BUFFERTOOSMALL)	{
								pFetchParams->cRowsReturned = i;
								return hr;
							}
							break;

						case VT_VARIANT:

							switch(var->vt){
								case VT_BOOL:							
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->boolVal,nOrdinalFieldPos, j);
								break;
								case VT_UI1:
								case VT_I1:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->bVal,nOrdinalFieldPos, j);
								break;
								case VT_UI2:
								case VT_I2:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->iVal,nOrdinalFieldPos, j);
								break;
								case VT_INT:
								case VT_UINT:
								case VT_UI4:
								case VT_I4:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->lVal,nOrdinalFieldPos, j);
								break;
								case VT_R4:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->fltVal,nOrdinalFieldPos, j);
								break;
								case VT_R8:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->dblVal,nOrdinalFieldPos, j);
								break;
								case VT_DATE:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->date,nOrdinalFieldPos, j);
								break;
								case VT_CY:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(unsigned char*)&var->cyVal,nOrdinalFieldPos, j);
								break;
                        case DBTYPE_ANYVARIANT:
								case VT_BSTR:
								default:
							         CvtToVariant(VT_VARIANT,pDestBytePtr,(pSrcColValue ? (unsigned char*)&pSrcColValue : (unsigned char*)&i64),nOrdinalFieldPos, j);
								break;
							}
							break;

						case VT_BLOB:

							hr = GetBlobData(pDestBytePtr, pSrcColValue, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize);
							if (FAILED(hr) || hr == DB_S_BUFFERTOOSMALL)	{
								pFetchParams->cRowsReturned = i;
								return hr;
							}

							break;

						case DBTYPE_BYTES:
						case DBTYPE_CHARS:
						case DBTYPE_WCHARS:

                  case VT_UI1:
                  case VT_I1:
                  case VT_I2:
                  case VT_UI2:
						case VT_I4:
                  case VT_UI4:
                  case VT_INT:
                  case VT_UINT:
						case VT_BOOL:

						case VT_R4:
						case VT_R8:
						case VT_DATE:
						case VT_CY:

							switch(var->vt){
								case VT_BOOL:							
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->boolVal, nOrdinalFieldPos, j);
								break;
								case VT_I1:
								case VT_UI1:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->bVal, nOrdinalFieldPos, j);
								break;
								case VT_UI2:
								case VT_I2:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->iVal, nOrdinalFieldPos, j);
								break;
								case VT_INT:
								case VT_UINT:
								case VT_UI4:
								case VT_I4:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->lVal, nOrdinalFieldPos, j);
								break;
								case VT_R4:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->fltVal, nOrdinalFieldPos, j);
								break;
								case VT_R8:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->dblVal, nOrdinalFieldPos, j);
								break;
								case VT_DATE:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->date, nOrdinalFieldPos, j);
								break;
								case VT_CY:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (unsigned char*)&var->cyVal, nOrdinalFieldPos, j);
								break;
								case VT_BSTR:
								default:
								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, (pSrcColValue ? (unsigned char*)&pSrcColValue : (unsigned char*)&i64), nOrdinalFieldPos, j);
								break;
							}

							// Can't coerce the actual data to the requested type
							if (hr == DB_E_CANTCOERCE){
								if (m_pDataColumns[j].dwBinding == DBBINDING_VARIANT){
									((DBVARIANT *)pDestColValue)->vt = DBTYPE_HRESULT;
									((DBVARIANT *)pDestColValue)->scode = DB_E_CANTCOERCE;
								}

								if (m_pDataColumns[j].obInfo != DB_NOVALUE){
									pDestColValue = pDestBytePtr + m_pDataColumns[j].obInfo;
									*(DWORD *)pDestColValue = DB_CANTCOERCE;
								}
							}

							break;

						default:
							
							// Can't coerce the actual data to the requested type
							if (m_pDataColumns[j].dwBinding == DBBINDING_VARIANT){
								((DBVARIANT *)pDestColValue)->vt = DBTYPE_HRESULT;
								((DBVARIANT *)pDestColValue)->scode = DB_E_CANTCOERCE;
							}

							if (m_pDataColumns[j].obInfo != DB_NOVALUE){
								pDestColValue = pDestBytePtr + m_pDataColumns[j].obInfo;
								*(DWORD *)pDestColValue = DB_CANTCOERCE;
							}

							break;
					}
				// ...for bookmark columns
				}else if (m_pDataColumns[j].columnID.guid == bmkguid){
					nOrdinalFieldPos = m_pControl->m_nColumns + m_pDataColumns[j].columnID.lNumber;
					
					try{
						value=m_pRecordSet->GetBookmark();
					}catch(CDaoException *e){
						e->Delete();
						break;
					}
					
					const VARIANT * var=(LPCVARIANT)value;
					pSrcColValue=(unsigned char*)var;
					switch (m_pDataColumns[j].columnID.lNumber){
						case BMK_NUMBER_BMKTEMPORARY:
							if (m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == VT_BLOB){
									//BSTR bstr;
									//if((hr=BstrFromVector(var->parray,(unsigned short **)&bstr))!=S_OK){
									//		pFetchParams->cRowsReturned = i;
									//		return hr;
									//}
					            void *ppvData;
                           SafeArrayAccessData(((LPCVARIANT)var)->parray,&ppvData);
                           long int ub,lb;
                           SafeArrayGetUBound(((LPCVARIANT)var)->parray,1,&ub);
                           SafeArrayGetLBound(((LPCVARIANT)var)->parray,1,&lb);
                           long int len=ub-lb+1;

									BLOB blob;
									blob.cbSize=/*3+lstrlenW(bstr)*2*/1+len;
									blob.pBlobData=(unsigned char*)ppvData;
									hr = GetBlobData(pDestBytePtr,(unsigned char*) &blob, nOrdinalFieldPos, j, pFetchParams, &cbCumulativeOutOfLineMemSize,"a");
                           SafeArrayUnaccessData(((LPCVARIANT)var)->parray);
									//::SysFreeString(bstr);
									if (FAILED(hr) || hr == DB_S_BUFFERTOOSMALL)	{
										pFetchParams->cRowsReturned = i;
										return hr;
									}
							}else if	(m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == DBTYPE_BYTES ||
									 m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == DBTYPE_CHARS ||
									 m_pControl->m_pMetaData[nOrdinalFieldPos].dwColType == DBTYPE_WCHARS){

								hr = GetFixedLengthInLineMemTypes(m_pDataColumns[j].dwDataType, pDestBytePtr, pSrcColValue, nOrdinalFieldPos, j);
								if (hr == DB_E_CANTCOERCE){
									if (m_pDataColumns[j].dwBinding == DBBINDING_VARIANT){
										((DBVARIANT *)pDestColValue)->vt = DBTYPE_HRESULT;
										((DBVARIANT *)pDestColValue)->scode = DB_E_CANTCOERCE;
									}
									if (m_pDataColumns[j].obInfo != DB_NOVALUE){
										pDestColValue = pDestBytePtr + m_pDataColumns[j].obInfo;
										*(DWORD *)pDestColValue = DB_CANTCOERCE;
									}
								}
							}else{ // bookmarks can only be retreived as BLOBS or BYTE\CHAR\WCHAR arrays
									pFetchParams->cRowsReturned = i;
									return E_FAIL;
							}

							break;

						// other types of bookmarks are not implemented in this sample
						default:	
									pFetchParams->cRowsReturned = i;
									return E_NOTIMPL;
					}//switch
				}//bookmarks
			}//DB_NOVALUE
		}//for
	   really_fetch++;
		try{
	      if(i+1<pFetchParams->cRowsReturned)m_pRecordSet->MoveNext();
			else break;
			if(m_pRecordSet->IsEOF())break;
		}catch(CDaoException *e){
		   e->Delete();
			break;
		}
	}


// update the current row position
pFetchParams->cRowsReturned=really_fetch;

TRACE("\nfetched %d",really_fetch);

if (m_pRecordSet->IsBOF() || m_pRecordSet->IsEOF())return DB_S_ENDOFCURSOR;

return S_OK;
}

// This function is called by CRowCursor FetchRows(). It fetches all the data types whose values
// are directly stored in in-line memory.
HRESULT CRowCursor::GetFixedLengthInLineMemTypes(DWORD dwRequestedType, BYTE *pDestBytePtr, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol)
{
	ULONG i;		// loop variable

	// requested # of bytes of data
	ULONG length = m_pDataColumns[CurrentCol].cbMaxLen;

	// available # of bytes of data
	ULONG cbAvailableLength = length;

	// ptr to the beginning of a column of data within the current row of in-line memory
	BYTE *pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obData;
	
	// convert to the requested type in the outer switch...
	switch (dwRequestedType){
		case DBTYPE_BYTES:

			// ...from the actual type of the data in the resultset
			if (m_pControl->m_pMetaData[nOrdinalPos].dwColType == DBTYPE_BYTES)
				for (i = 0; i < length; i++)pDestColVal[i] = pSrcColVal[i];
			else if (m_pControl->m_pMetaData[nOrdinalPos].dwColType == VT_LPSTR){
				cbAvailableLength = strlen(*(LPSTR *)pSrcColVal) + 1;
				if (cbAvailableLength < length)length = cbAvailableLength;
				for (i = 0; i < length; i++)pDestColVal[i] = (BYTE)((*(LPSTR *)pSrcColVal)[i]);
			}else	return DB_E_CANTCOERCE;

			if (m_pDataColumns[CurrentCol].obVarDataLen != DB_NOVALUE){
				pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obVarDataLen;
				*(DWORD *)pDestColVal = cbAvailableLength;
			}
			break;

		case DBTYPE_CHARS:

			// ...from the actual type of the data in the resultset
			if (m_pControl->m_pMetaData[nOrdinalPos].dwColType == DBTYPE_CHARS)
				for (i = 0; i < length; i++)((CHAR *)pDestColVal)[i] = (CHAR)pSrcColVal[i];
			else if (m_pControl->m_pMetaData[nOrdinalPos].dwColType == VT_LPSTR){
				cbAvailableLength = strlen(*(LPSTR *)pSrcColVal) + 1;

				if (cbAvailableLength < length)length = cbAvailableLength;

				for (i = 0; i < length; i++)((CHAR *)pDestColVal)[i] = (*(LPSTR *)pSrcColVal)[i];
			}else	return DB_E_CANTCOERCE;

			if (m_pDataColumns[CurrentCol].obVarDataLen != DB_NOVALUE){
				pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obVarDataLen;
				*(DWORD *)pDestColVal = cbAvailableLength;
			}

			break;

		case DBTYPE_WCHARS:

			// ...from the actual type of the data in the resultset
			if (m_pControl->m_pMetaData[nOrdinalPos].dwColType == DBTYPE_WCHARS)
				for (i = 0; i < length; i++)pDestColVal[i] = pSrcColVal[i];
			else if (m_pControl->m_pMetaData[nOrdinalPos].dwColType == VT_LPWSTR){
				cbAvailableLength = 2 * (wcslen(*(LPWSTR *)pSrcColVal) + 1);

				if (cbAvailableLength < length)length = cbAvailableLength;

				for (i = 0; i < length; i++)pDestColVal[i] = (BYTE)((*(LPSTR *)pSrcColVal)[i]);
			}else	return DB_E_CANTCOERCE;

			if (m_pDataColumns[CurrentCol].obVarDataLen != DB_NOVALUE){
				pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obVarDataLen;
				*(DWORD *)pDestColVal = cbAvailableLength;
			}

			break;

      case VT_UI4:
      case VT_I4:
      case VT_INT:
      case VT_UINT:

			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, long, lVal)
								break;
            case VT_UI2:
            case VT_I2:		CVTFIXEDTYPE(short, long, lVal)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
            case VT_I4:		CVTFIXEDTYPE(long, long, lVal)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, long, lVal)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, long, lVal)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, long, lVal)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, long, lVal)
								break;

				default:		return DB_E_CANTCOERCE;
			}

			break;

		case VT_UI2:
		case VT_I2:

			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, short, iVal)
								break;
            case VT_UI2:
				case VT_I2:		CVTFIXEDTYPE(short, short, iVal)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
				case VT_I4:		CVTFIXEDTYPE(long, short, iVal)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, short, iVal)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, short, iVal)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, short, iVal)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, short, iVal)
								break;

				default:		return DB_E_CANTCOERCE;
			}

			break;

		case VT_BOOL:

			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, short, xbool)
								break;
            case VT_UI2:
				case VT_I2:		CVTFIXEDTYPE(short, short, xbool)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
				case VT_I4:		CVTFIXEDTYPE(long, short, xbool)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, short, xbool)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, short, xbool)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, short, xbool)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, short, xbool)
								break;

				default:		return DB_E_CANTCOERCE;
			}

			break;

		case VT_R4:
		
			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, float, fltVal)
								break;
				
            case VT_UI2:
            case VT_I2:		CVTFIXEDTYPE(short, float, fltVal)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
				case VT_I4:		CVTFIXEDTYPE(long, float, fltVal)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, float, fltVal)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, float, fltVal)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, float, fltVal)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, float, fltVal)
								break;

				default:		return DB_E_CANTCOERCE;
			}

			break;

		case VT_R8:
		
			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, double, dblVal)
								break;
				
            case VT_UI2:
				case VT_I2:		CVTFIXEDTYPE(short, double, dblVal)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
				case VT_I4:		CVTFIXEDTYPE(long, double, dblVal)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, double, dblVal)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, double, dblVal)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, double, dblVal)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, double, dblVal)
								break;

				default:		return DB_E_CANTCOERCE;
			}

			break;

		case VT_DATE:
		
			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, DATE, date)
								break;
            case VT_UI2:
				case VT_I2:		CVTFIXEDTYPE(short, DATE, date)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
				case VT_I4:		CVTFIXEDTYPE(long, DATE, date)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, DATE, date)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, DATE, date)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, DATE, date)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, DATE, date)
								break;

				default:		return DB_E_CANTCOERCE;
			}

			break;

		case VT_CY:

			// ...from the actual type of the data in the resultset
			switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
            case VT_I1:
            case VT_UI1: CVTFIXEDTYPE(char, long, lVal)
								break;
            case VT_UI2:
				case VT_I2:		CVTFIXEDTYPE(short, long, lVal)
								break;

            case VT_INT:
            case VT_UINT:
            case VT_UI4:
				case VT_I4:		CVTFIXEDTYPE(long, long, lVal)
								break;

				case VT_R4:		CVTFIXEDTYPE(float, long, lVal)
								break;

				case VT_R8:		CVTFIXEDTYPE(double, long, lVal)
								break;

				case VT_BOOL:	CVTFIXEDTYPE(BOOL, long, lVal)
								break;

				case VT_DATE:	CVTFIXEDTYPE(DATE, long, lVal)
								break;

				default:		return DB_E_CANTCOERCE;
			}
			break;
		default:	return DB_E_CANTCOERCE;
	}

	return S_OK;
}



// This function is called by CRowCursor FetchRows(). It fetches columns whose data type is VT_BLOB

HRESULT CRowCursor::GetBlobData(BYTE *pDestBytePtr, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol, DBFETCHROWS *pFetchParams, ULONG *pcbCumulativeOutOfLineMemSize,char *add)
{
	// ptr to the beginning of a column of data within the current row of in-line memory
	BYTE *pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obData;
	// can't coerce any other data type to VT_BLOB
	if (m_pControl->m_pMetaData[nOrdinalPos].dwColType != VT_BLOB){
		if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_VARIANT){
			((DBVARIANT *)pDestColVal)->vt = DBTYPE_HRESULT;
			((DBVARIANT *)pDestColVal)->scode = DB_E_CANTCOERCE;
		}
		if (m_pDataColumns[CurrentCol].obInfo != DB_NOVALUE){
			pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obInfo;
			*(DWORD *)pDestColVal = DB_CANTCOERCE;
		}
		return S_OK;
	}
	ULONG i=0;		// loop counter
	// size in bytes of requested data
	ULONG length = ((BLOB *)pSrcColVal)->cbSize;
	// size in bytes of available data
	ULONG cbAvailableLength = length;
	// flag to indicate that the fetched data is truncated
	BOOL bTruncated = FALSE;									
	// truncate data if requested length is less than available length.
	if (m_pDataColumns[CurrentCol].cbMaxLen != DB_NOMAXLENGTH && length > m_pDataColumns[CurrentCol].cbMaxLen){
		bTruncated = TRUE;
		length = m_pDataColumns[CurrentCol].cbMaxLen;
	}
	// check if there is enough caller allocated out-of-line memory
	if (!(pFetchParams->dwFlags & DBROWFETCH_CALLEEALLOCATES)){
		if (!pFetchParams->pVarData || !pFetchParams->cbVarData)	return DB_E_BADFETCHINFO;
		if (pFetchParams->cbVarData - *pcbCumulativeOutOfLineMemSize < length)return DB_S_BUFFERTOOSMALL;
	}
	if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_VARIANT){
		((DBVARIANT *)pDestColVal)->vt = VT_BLOB;
		(((DBVARIANT *)pDestColVal)->blob).cbSize = length;
		// assign a pointer to the actual blob data to a DBVARIANT in-line memory
		(((DBVARIANT *)pDestColVal)->blob).pBlobData = (BYTE *)pFetchParams->pVarData + *pcbCumulativeOutOfLineMemSize;
		// copy the actual blob data into out-of-line memory
		if(add){
			((((DBVARIANT *)pDestColVal)->blob).pBlobData)[0]=add[0];
			for (i = 1;i < length;i++)
				((((DBVARIANT *)pDestColVal)->blob).pBlobData)[i] = (((BLOB *)pSrcColVal)->pBlobData)[i-1];
		}else
		for (i = 0;i < length;i++)
			((((DBVARIANT *)pDestColVal)->blob).pBlobData)[i] = (((BLOB *)pSrcColVal)->pBlobData)[i];
	
	}else if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_DEFAULT){
		// assign a pointer to the actual blob data to in-line memory
		((BLOB *)pDestColVal)->cbSize=length;
		((BLOB *)pDestColVal)->pBlobData = (BYTE *)pFetchParams->pVarData + *pcbCumulativeOutOfLineMemSize;
		// copy the actual blob data into out-of-line memory
		if(add){
			(((BLOB *)pDestColVal)->pBlobData)[0] = add[0];
			for (i = 1;i < length;i++)
				(((BLOB *)pDestColVal)->pBlobData)[i] = (((BLOB *)pSrcColVal)->pBlobData)[i-1];
		}else
		for (i = 0;i < length;i++)
			(((BLOB *)pDestColVal)->pBlobData)[i] = (((BLOB *)pSrcColVal)->pBlobData)[i];
	}
	// fill column info field if requested by client and data is truncated.
	if (bTruncated){
		if (m_pDataColumns[CurrentCol].obInfo != DB_NOVALUE){
			pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obInfo;
			*(DWORD *)pDestColVal = DB_TRUNCATED;
		}
	}
	// fill column length field with available # of bytes if requested by client.
	if (m_pDataColumns[CurrentCol].obVarDataLen != DB_NOVALUE){
		pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obVarDataLen;
		*(DWORD *)pDestColVal = cbAvailableLength;
	}
	// update the current out-of-line memory size for next column that uses it and pass it out
	*pcbCumulativeOutOfLineMemSize += length;
	return S_OK;
}

// This function is called by CRowCursor FetchRows(). It fetches columns with whose data type
// is VT_I2, VT_I4, VT_BOOL, VT_LPSTR or VT_LPWSTR as a string i.e (VT_BSTR, VT_LPSTR or VT_LPWSTR)
// after appropriate coersion.

HRESULT CRowCursor::CvtToString(DWORD dwRequestedType, BYTE *pDestBytePtr, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol, DBFETCHROWS *pFetchParams, ULONG *pcbCumulativeOutOfLineMemSize)
{
	USES_CONVERSION;

	// ptr to the beginning of a column of data within the current row of in-line memory
	BYTE *pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obData;

	// handle NULL string data
	if (pSrcColVal == NULL || (m_pControl->m_pMetaData[nOrdinalPos].dwColType == VT_LPSTR && *(LPSTR *)pSrcColVal == NULL) ||
		(m_pControl->m_pMetaData[nOrdinalPos].dwColType == VT_LPWSTR && *(LPWSTR *)pSrcColVal == NULL)){
		if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_VARIANT){
			((DBVARIANT *)pDestColVal)->vt = DBTYPE_NULL;
			if (dwRequestedType == VT_LPSTR)((DBVARIANT *)pDestColVal)->pszVal = NULL;
			else if (dwRequestedType == VT_LPWSTR)((DBVARIANT *)pDestColVal)->pwszVal = NULL;
			else if (dwRequestedType == VT_BSTR)((DBVARIANT *)pDestColVal)->bstrVal = NULL;
		}else if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_DEFAULT)*(DWORD *)pDestColVal = NULL;
		
		if (m_pDataColumns[CurrentCol].obInfo != DB_NOVALUE){
			pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obInfo;
			*(DWORD *)pDestColVal = DB_NULL;
		}
		return S_OK;
	}

	// all numeric types converted to long first, so that ltoa can be used for converting them
	// into a string.
	long lngVal = 0;
	double dVal=0;
	__int64 iVal=0;
	// size in bytes of requested data
	ULONG length = 0;
	// size in bytes of available data
	ULONG cbAvailableLength = 0;
	CURRENCY curVal;
	curVal.Lo=curVal.Hi=0;
	DATE dateVal=0;
	// flag to indicate that the fetched data is truncated
	BOOL bTruncated = FALSE;
	// determine the size of out-of-Line memory for this row data type.
	switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
      case VT_UI1:
      case VT_I1: lngVal = (long)(*(unsigned char *)pSrcColVal);
						length = 33;
						break;
      case VT_UI2:
      case VT_I2:	lngVal = (long)(*(short *)pSrcColVal);
						length = 33;
						break;
		case VT_BOOL:	lngVal = (long)(*(BOOL *)pSrcColVal);
						length = 33;
						break;
      case VT_INT:
      case VT_UINT:
      case VT_UI4:
		case VT_I4:	lngVal = *(long *)pSrcColVal;
						length = 33;
						break;
		case VT_LPSTR:	length = strlen(*(LPSTR *)pSrcColVal) + 1;
						break;
		case VT_LPWSTR:	length = wcslen(*(LPWSTR *)pSrcColVal) + 1;
						break;
		case VT_R4: dVal = *(float *)pSrcColVal;
						length = 33;
						break;
		case VT_R8: dVal = *(double *)pSrcColVal;
						length = 33;
						break;
		case VT_CY: curVal = *(CURRENCY *)pSrcColVal;
						length = 33;
						break;
		case VT_DATE: dateVal = *(DATE *)pSrcColVal;
						length = 33;
						break;

		// can't coerce any other data type to a string!
		default:		if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_VARIANT){
							((DBVARIANT *)pDestColVal)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestColVal)->scode = DB_E_CANTCOERCE;
						}

						if (m_pDataColumns[CurrentCol].obInfo != DB_NOVALUE){
							pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obInfo;
							*(DWORD *)pDestColVal = DB_CANTCOERCE;
						}
			
						return S_OK;
	}

	// if data type requested by client is a Unicode string...
	if (dwRequestedType == VT_LPWSTR)length *= 2;
	cbAvailableLength = length;
	// truncate data if requested length is less than available length.
	if (m_pDataColumns[CurrentCol].cbMaxLen != DB_NOMAXLENGTH && length > m_pDataColumns[CurrentCol].cbMaxLen){
		bTruncated = TRUE;
		length = m_pDataColumns[CurrentCol].cbMaxLen;
	}

	// check if there is enough caller allocated out-of-line memory for VT_LPSTRs and VT_LPWSTRs
	// only! VT_BSTRs are neither in-line nor out-of-line memory!
	if (dwRequestedType != VT_BSTR && !(pFetchParams->dwFlags & DBROWFETCH_CALLEEALLOCATES)){
		if (!pFetchParams->pVarData || !pFetchParams->cbVarData)return DB_E_BADFETCHINFO;
		if (pFetchParams->cbVarData - *pcbCumulativeOutOfLineMemSize < length)return DB_S_BUFFERTOOSMALL;
	}

	// requested data type could be an ANSI or UNICODE string in a DBVARIANT or in direct
	// memory; if so, assign a pointer to the string to in-line memory and copy the actual
	// string data into out-of-line memory; if not, can't coerce....
	if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_VARIANT)	{
		((DBVARIANT *)pDestColVal)->vt = (VARTYPE)dwRequestedType;

		//
		if (dwRequestedType == VT_LPSTR)
			((DBVARIANT *)pDestColVal)->pszVal = (LPSTR)((BYTE *)pFetchParams->pVarData + *pcbCumulativeOutOfLineMemSize);
		else if (dwRequestedType == VT_LPWSTR)
			((DBVARIANT *)pDestColVal)->pwszVal = (LPWSTR)((BYTE *)pFetchParams->pVarData + *pcbCumulativeOutOfLineMemSize);

		switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
			case VT_I2:		
			case VT_BOOL:
			case VT_I4:	if (dwRequestedType == VT_LPSTR)
								_ltoa (lngVal, ((DBVARIANT *)pDestColVal)->pszVal, 10);
							else if (dwRequestedType == VT_LPWSTR)
								_ltow (lngVal, ((DBVARIANT *)pDestColVal)->pwszVal, 10);
							else if (dwRequestedType == VT_BSTR){
								CHAR bstrVal[33];
								_ltoa (lngVal, bstrVal, 10);
								((DBVARIANT *)pDestColVal)->bstrVal = SysAllocString(A2W(bstrVal));

								if (((DBVARIANT *)pDestColVal)->bstrVal == NULL)
									return E_OUTOFMEMORY;
							}
							break;
			case VT_R4:
			case VT_R8:
				{
				CString str;
				str.Format("%f",dVal);
				if (dwRequestedType == VT_LPSTR)
					strcpy(((DBVARIANT *)pDestColVal)->pszVal, str);
				else if (dwRequestedType == VT_LPWSTR)
					wcscpy(((DBVARIANT *)pDestColVal)->pwszVal, (const unsigned short*)(LPCTSTR)str);
				else if (dwRequestedType == VT_BSTR){
						((DBVARIANT *)pDestColVal)->bstrVal = str.AllocSysString();
						if (((DBVARIANT *)pDestColVal)->bstrVal == NULL)
						return E_OUTOFMEMORY;
				}
				}
				 break;
			case VT_CY:
				{
				COleCurrency cur(curVal);
				CString str;
				str=cur.Format();
				if (dwRequestedType == VT_LPSTR)
					strcpy(((DBVARIANT *)pDestColVal)->pszVal, str);
				else if (dwRequestedType == VT_LPWSTR)
					wcscpy(((DBVARIANT *)pDestColVal)->pwszVal, (const unsigned short*)(LPCTSTR)str);
				else if (dwRequestedType == VT_BSTR){
						((DBVARIANT *)pDestColVal)->bstrVal = str.AllocSysString();
						if (((DBVARIANT *)pDestColVal)->bstrVal == NULL)return E_OUTOFMEMORY;
				}
				}
				break;	
			case VT_DATE:
				{
				COleDateTime cur(dateVal);
				CString str;
				str=cur.Format();
				if (dwRequestedType == VT_LPSTR)
					strcpy(((DBVARIANT *)pDestColVal)->pszVal, str);
				else if (dwRequestedType == VT_LPWSTR)
					wcscpy(((DBVARIANT *)pDestColVal)->pwszVal, (const unsigned short*)(LPCTSTR)str);
				else if (dwRequestedType == VT_BSTR){
						((DBVARIANT *)pDestColVal)->bstrVal = str.AllocSysString();
						if (((DBVARIANT *)pDestColVal)->bstrVal == NULL)return E_OUTOFMEMORY;
				}
				}
				break;	
			case VT_LPSTR:	if (dwRequestedType == VT_LPSTR){
								strncpy (((DBVARIANT *)pDestColVal)->pszVal, *(LPSTR *)pSrcColVal, length);
			
							if (bTruncated)
									*((CHAR *)(((DBVARIANT *)pDestColVal)->pszVal) + length - 1) = (CHAR)0;
							}else if (dwRequestedType == VT_LPWSTR){
								wcsncpy (((DBVARIANT *)pDestColVal)->pwszVal, A2W(*(LPSTR *)pSrcColVal), length/2);
								
								if (bTruncated)
									*((WCHAR *)((BYTE *)(((DBVARIANT *)pDestColVal)->pwszVal) + length - 2)) = (WCHAR)0;
							}else if (dwRequestedType == VT_BSTR){
								((DBVARIANT *)pDestColVal)->bstrVal = SysAllocString(A2W(*(LPSTR *)pSrcColVal));
								if (((DBVARIANT *)pDestColVal)->bstrVal == NULL)return E_OUTOFMEMORY;
							}
							break;

			case VT_LPWSTR:	if (dwRequestedType == VT_LPSTR){
								strncpy (((DBVARIANT *)pDestColVal)->pszVal, W2A(*(LPWSTR *)pSrcColVal), length);
			
							if (bTruncated)
									*((CHAR *)(((DBVARIANT *)pDestColVal)->pszVal) + length - 1) = (CHAR)0;
							}else if (dwRequestedType == VT_LPWSTR){
								wcsncpy (((DBVARIANT *)pDestColVal)->pwszVal, *(LPWSTR *)pSrcColVal, length/2);
								
								if (bTruncated)
									*((WCHAR *)((BYTE *)(((DBVARIANT *)pDestColVal)->pwszVal) + length - 2)) = (WCHAR)0;
							}else if (dwRequestedType == VT_BSTR){
								((DBVARIANT *)pDestColVal)->bstrVal = SysAllocString(*(LPWSTR *)pSrcColVal);

								if (((DBVARIANT *)pDestColVal)->bstrVal == NULL)return E_OUTOFMEMORY;
							}

							break;
		}
	}else if (m_pDataColumns[CurrentCol].dwBinding == DBBINDING_DEFAULT){
		*(LPSTR *)pDestColVal = (LPSTR)((BYTE *)pFetchParams->pVarData + *pcbCumulativeOutOfLineMemSize);

		switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
			case VT_I4:
			case VT_I2:
			case VT_BOOL:	
							if (dwRequestedType == VT_LPSTR)
								_ltoa (lngVal, *(LPSTR *)pDestColVal, 10);
							else if (dwRequestedType == VT_LPWSTR)
								_ltow (lngVal, *(LPWSTR *)pDestColVal, 10);
							else if (dwRequestedType == VT_BSTR){
								CHAR bstrVal[33];
								_ltoa (lngVal, bstrVal, 10);
								*(BSTR *)pDestColVal = SysAllocString(A2W(bstrVal));

								if (*(BSTR *)pDestColVal == NULL)
									return E_OUTOFMEMORY;
							}
							break;
			case VT_R4:
			case VT_R8:
				{
				CString str;
				str.Format("%f",dVal);
				if (dwRequestedType == VT_LPSTR)
					strcpy(*(LPSTR *)pDestColVal, str);
				else if (dwRequestedType == VT_LPWSTR)
					wcscpy(*(LPWSTR *)pDestColVal, (const unsigned short*)(LPCTSTR)str);
				else if (dwRequestedType == VT_BSTR){
						*(BSTR *)pDestColVal = str.AllocSysString();
						if (*(BSTR *)pDestColVal == NULL)return E_OUTOFMEMORY;
				}
				}
				break;
			case VT_CY:
				{
				COleCurrency cur(curVal);
				CString str;
				str=cur.Format();
				if (dwRequestedType == VT_LPSTR)
					strcpy(*(LPSTR *)pDestColVal, str);
				else if (dwRequestedType == VT_LPWSTR)
					wcscpy(*(LPWSTR *)pDestColVal, (const unsigned short*)(LPCTSTR)str);
				else if (dwRequestedType == VT_BSTR){
						*(BSTR *)pDestColVal = str.AllocSysString();
						if (*(BSTR *)pDestColVal == NULL)return E_OUTOFMEMORY;
				}
				}
				break;
			case VT_DATE:
				{
				COleDateTime cur(dateVal);
				CString str;
				str=cur.Format();
				if (dwRequestedType == VT_LPSTR)
					strcpy(*(LPSTR *)pDestColVal, str);
				else if (dwRequestedType == VT_LPWSTR)
					wcscpy(*(LPWSTR *)pDestColVal, (const unsigned short*)(LPCTSTR)str);
				else if (dwRequestedType == VT_BSTR){
						*(BSTR *)pDestColVal = str.AllocSysString();
						if (*(BSTR *)pDestColVal == NULL)return E_OUTOFMEMORY;
				}
				}
				break;

			
			case VT_LPSTR:	if (dwRequestedType == VT_LPSTR){
								strncpy (*(LPSTR *)pDestColVal, *(LPSTR *)pSrcColVal, length);
	
								if (bTruncated)
									*((CHAR *)(pDestColVal + length - 1)) = (CHAR)0;
							}else if (dwRequestedType == VT_LPWSTR){
								wcsncpy (*(LPWSTR *)pDestColVal, A2W(*(LPSTR *)pSrcColVal), length/2);
							
								if (bTruncated)
									*((WCHAR *)(pDestColVal + length - 2)) = (WCHAR)0;
							}else if (dwRequestedType == VT_BSTR){
								*(BSTR *)pDestColVal = SysAllocString(A2W(*(LPSTR *)pSrcColVal));

								if (*(BSTR *)pDestColVal == NULL)
									return E_OUTOFMEMORY;
							}
							break;

			case VT_LPWSTR:	if (dwRequestedType == VT_LPSTR){
								strncpy (*(LPSTR *)pDestColVal, W2A(*(LPWSTR *)pSrcColVal), length);
	
							if (bTruncated)
									*((CHAR *)(pDestColVal + length - 1)) = (CHAR)0;
							}else if (dwRequestedType == VT_LPWSTR){
								wcsncpy (*(LPWSTR *)pDestColVal, *(LPWSTR *)pSrcColVal, length/2);
							
							if (bTruncated)
									*((WCHAR *)(pDestColVal + length - 2)) = (WCHAR)0;
							}else if (dwRequestedType == VT_BSTR){
								*(BSTR *)pDestColVal = SysAllocString(*(LPWSTR *)pSrcColVal);

								if (*(BSTR *)pDestColVal == NULL)
									return E_OUTOFMEMORY;
							}
							break;
		}
	}

	if (dwRequestedType == VT_LPSTR || dwRequestedType == VT_LPWSTR){
		// fill column info field if requested by client and data is truncated.
		if (bTruncated){
			if (m_pDataColumns[CurrentCol].obInfo != DB_NOVALUE){
				pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obInfo;
				*(DWORD *)pDestColVal = DB_TRUNCATED;
			}
		}										
		// fill column length field with available # of bytes if requested by client.
		if (m_pDataColumns[CurrentCol].obVarDataLen != DB_NOVALUE){
			pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obVarDataLen;
			*(DWORD *)pDestColVal = cbAvailableLength;
		}
	}
	// update the current out-of-line memory size for next column that uses it and pass it out
	*pcbCumulativeOutOfLineMemSize += length;
	return S_OK;
}

// This function is called by CRowCursor::FetchRows(). It fetches columns with any data type that
// can be held inside a VARIANT. In other words, it coerces those data types into the type VT_VARIANT
// and puts it into a DBVARIANT in in-line memory. (This is the way standard data-aware controls like
// the TextBox and the ListBox request for data)

HRESULT CRowCursor::CvtToVariant(DWORD /*dwRequestedType*/, BYTE *pDestBytePtr, BYTE *pSrcColVal, LONG nOrdinalPos, ULONG CurrentCol)
{
	USES_CONVERSION;
	BYTE *pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obData;
	VARIANT *vnt = (VARIANT *)CoTaskMemAlloc(sizeof(VARIANT));
	VariantInit(vnt);
	((DBVARIANT *)pDestColVal)->vt = VT_BYREF|VT_VARIANT;
	((DBVARIANT *)pDestColVal)->pvarVal = vnt;	
	vnt->vt = (VARTYPE)m_pControl->m_pMetaData[nOrdinalPos].dwColType;
	switch (m_pControl->m_pMetaData[nOrdinalPos].dwColType){
      case VT_INT:
      case VT_UINT:
      case VT_UI4:
      case VT_I4:		vnt->lVal = *(long *)pSrcColVal;
						break;
      case VT_I1:
      case VT_UI1:	vnt->bVal = *(unsigned char *)pSrcColVal;
						break;
      case VT_UI2:
		case VT_I2:		vnt->iVal = *(short *)pSrcColVal;
						break;
		case VT_R4:		vnt->fltVal = *(float *)pSrcColVal;
						break;
		case VT_R8:		vnt->dblVal = *(double *)pSrcColVal;
						break;
		case VT_BOOL:	vnt->boolVal = *(VARIANT_BOOL *)pSrcColVal;
						break;
		case VT_CY:		vnt->cyVal = *(CY *)pSrcColVal;
						break;
		case VT_DATE:	vnt->date = *(DATE *)pSrcColVal;
						break;
		case VT_LPSTR:
      case VT_PTR:            
                  vnt->bstrVal = SysAllocString(A2W(*(LPSTR *)pSrcColVal));
						vnt->vt = VT_BSTR;
						break;
		case VT_LPWSTR:	vnt->bstrVal = SysAllocString(*(LPWSTR *)pSrcColVal);
						vnt->vt = VT_BSTR;
						break;
		// can't coerce any other data type to a VARIANT!
		default:		((DBVARIANT *)pDestColVal)->vt = DBTYPE_HRESULT;
						((DBVARIANT *)pDestColVal)->scode = DB_E_CANTCOERCE;
						if (m_pDataColumns[CurrentCol].obInfo != DB_NOVALUE){
							pDestColVal = pDestBytePtr + m_pDataColumns[CurrentCol].obInfo;
							*(DWORD *)pDestColVal = DB_CANTCOERCE;
						}
						break;
	}

   return S_OK;
}

// This function is called by ICursorMove::Move. It returns the requested row position to move
// to from the given non-standard bookmark. You will need to re-implement this function according
// to the bookmark scheme that you will be using.

/////////////////////////////////////////////////////////////////////////
//
//	ICursorScroll implementation
//
/////////////////////////////////////////////////////////////////////////
 
ULONG CRowCursor::XCURSORSCROLL::AddRef()
{
    METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
    return pThis->ExternalAddRef();
}
    
ULONG CRowCursor::XCURSORSCROLL::Release()
{
    METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
    return pThis->ExternalRelease();
}
    
HRESULT CRowCursor::XCURSORSCROLL::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	return pThis->ExternalQueryInterface(&iid, ppvObj);
}

HRESULT CRowCursor::XCURSORSCROLL::GetColumnsCursor(REFIID riid, IUnknown *ppunkColumnsCursor[], ULONG *pcRows)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
   // delegate to ICursor::GetColumnsCursor()
	return pThis->m_xCURSOR.GetColumnsCursor(riid, ppunkColumnsCursor, pcRows);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::SetBindings(ULONG cCol, 
      DBCOLUMNBINDING rgBoundColumns[], ULONG cbRowLength, DWORD dwFlags)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	return pThis->m_xCURSOR.SetBindings(cCol, rgBoundColumns, cbRowLength, dwFlags);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::GetBindings(
      ULONG *pcCol, DBCOLUMNBINDING *prgBoundColumns[], ULONG *pcbRowLength)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	return pThis->m_xCURSOR.GetBindings(pcCol, prgBoundColumns, pcbRowLength);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::GetNextRows(LARGE_INTEGER dlRowsToSkip, DBFETCHROWS *pFetchParams)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	return pThis->m_xCURSOR.GetNextRows(dlRowsToSkip, pFetchParams);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::Requery()
{
	METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
   return pThis->m_xCURSOR.Requery();
}	
    
HRESULT CRowCursor::XCURSORSCROLL::Move(ULONG cbBookmark, void *pBookmark, LARGE_INTEGER dlOffset, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	return pThis->m_xCURSORMOVE.Move(cbBookmark, pBookmark, dlOffset, pFetchParams);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::GetBookmark(DBCOLUMNID *pBookmarkType, ULONG cbMaxSize, ULONG *pcbBookmark, void *pBookmark)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
   return pThis->m_xCURSORMOVE.GetBookmark(pBookmarkType, cbMaxSize, pcbBookmark, pBookmark);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::Clone(DWORD dwFlags, REFIID riid, LPUNKNOWN *ppunkClonedCursor)
{
   METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
   return pThis->m_xCURSORMOVE.Clone(dwFlags, riid, ppunkClonedCursor);
}	
    
HRESULT CRowCursor::XCURSORSCROLL::Scroll(ULONG ulNumerator, ULONG ulDenominator, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;
	if(ulDenominator<ulNumerator)return DB_E_BADFETCHINFO;
	if(pThis->m_bUpdateInProgress)return DB_E_UPDATEINPROGRESS;
   
   if(pFetchParams)pFetchParams->cRowsReturned=0;
	
   // Initialize arguments for notifications
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);
	
	// Bookmarks should be passed in like this to allow VariantChangeByte
	// to be used by bound controls
	rgReasons[0].arg1.vt = VT_I4;
	rgReasons[0].arg1.lVal = ulNumerator;
	
	rgReasons[0].arg2.vt = VT_I4;
	rgReasons[0].arg2.lVal = ulDenominator;
		

	dwEventWhat = DBEVENT_CURRENT_ROW_CHANGED;
	cReasons = 1;
	rgReasons[0].dwReason = DBREASON_MOVEPERCENT;
	
	hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons);
	if (hr != S_OK){
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}

	try{
		if(!ulNumerator && !ulDenominator){
			pThis->m_pRecordSet->MoveFirst();
			pThis->m_pRecordSet->MovePrev();
		}else if(ulNumerator==ulDenominator){
			pThis->m_pRecordSet->MoveLast();
			pThis->m_pRecordSet->MoveNext();
		}else{
			pThis->m_pRecordSet->SetPercentPosition((float)ulNumerator * (100/ulDenominator));
		}
	}catch(CDaoException *e){
	   e->Delete();
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return E_FAIL;  
	}
	
	if(pFetchParams && pFetchParams->cRowsRequested){
		hr = pThis->FetchRows(pFetchParams);
	// If we didn't fetch rows, then check for BOF or EOF and return DB_S_ENDOFCURSOR
	}else if(pThis->m_pRecordSet->IsBOF() || pThis->m_pRecordSet->IsEOF())hr=DB_S_ENDOFCURSOR;
	else hr = S_OK;

	
	
	// If the event failed in any way, then call FailedToDo
	if (FAILED(hr)){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}



	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);
	rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
   pThis->MakeSafearrayFromBok(&psa);

	rgReasons[0].arg1.parray = psa;
	
	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	return hr;
}
    
HRESULT CRowCursor::XCURSORSCROLL::GetApproximatePosition(ULONG cbBookmark,	void *pBookmark, ULONG *pulNumerator, ULONG *pulDenominator)
{
	METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)

	if(!pBookmark)return DB_E_BADBOOKMARK;
	if(!pulDenominator || !pulNumerator)return E_OUTOFMEMORY;

	switch (*((BYTE *)pBookmark)){
		case DBBMK_BEGINNING:	
							*pulNumerator = 0;
							*pulDenominator = 0;
							break;
		case DBBMK_CURRENT:		
							try{
								*pulNumerator = (unsigned long)pThis->m_pRecordSet->GetPercentPosition();
							}catch(CDaoException *e){
							   e->Delete();
								*pulNumerator=50;
							}
							*pulDenominator = 100;
							if(*pulNumerator>*pulDenominator)*pulNumerator=*pulDenominator;
							break;
		case DBBMK_END: *pulNumerator = 1;
							 *pulDenominator=1;
							 break;
		case DBBMK_INVALID:		
							return DB_E_BADBOOKMARK;
		// If no standard bookmarks were passed in, then call function to 
		// get and move current position
		default:{
					COleVariant bok;
					BOOL is_eof,is_bof;
					is_eof=is_bof=FALSE;
					if(pThis->m_pRecordSet->IsBOF())is_bof=TRUE;
					else if(pThis->m_pRecordSet->IsEOF())is_eof=TRUE;
					else{
					try{
						bok=pThis->m_pRecordSet->GetBookmark();
					}catch(CDaoException* e){
						e->Delete();
					}
					}

					VARIANT var;
               VariantInit(&var);
					var.vt=VT_ARRAY|VT_UI1;
					//BSTR b=::SysAllocStringLen((BSTR)((char*)pBookmark+1),cbBookmark-1);
				   //VectorFromBstr((BSTR)((char*)pBookmark+1),&var.parray);
					//::SysFreeString(b);
               void *ppvData;
               SAFEARRAYBOUND rgsabound[1];	
               rgsabound[0].lLbound = 0;
	            rgsabound[0].cElements = cbBookmark-1;            
               var.parray=SafeArrayCreate(VT_UI1,1,rgsabound);
               SafeArrayAccessData(var.parray,&ppvData);
               memcpy(ppvData,((char*)pBookmark)+1,cbBookmark-1),
               SafeArrayUnaccessData(var.parray);

               COleVariant pas(var);
					try{
						pThis->m_pRecordSet->SetBookmark(pas);
					}catch(CDaoException* e){
						e->Delete();
						SafeArrayDestroy(var.parray);
						if(is_bof){
							try{
								pThis->m_pRecordSet->MoveFirst();
								pThis->m_pRecordSet->MovePrev();
							}catch(CDaoException* f){
								f->Delete();
							}																		
							}else if(is_eof){
								try{
									pThis->m_pRecordSet->MoveLast();
									pThis->m_pRecordSet->MoveNext();
								}catch(CDaoException* f){
									f->Delete();
								}																		
							}else{
								try{
									pThis->m_pRecordSet->SetBookmark(bok);
								}catch(CDaoException* f){
									f->Delete();
								}									
							}
						return DB_E_BADBOOKMARK;
					}
					::SafeArrayDestroy(var.parray);
								

					if(pThis->m_pRecordSet->IsBOF())*pulNumerator=*pulDenominator=0;
					else if(pThis->m_pRecordSet->IsEOF())*pulNumerator=*pulDenominator=1;
					else{
						try{
							*pulNumerator = (unsigned long)pThis->m_pRecordSet->GetPercentPosition();
						}catch(CDaoException *e){
						   e->Delete();
							*pulNumerator=50;
						}
						*pulDenominator = 100;
						if(*pulNumerator>*pulDenominator)*pulNumerator=*pulDenominator;					
					}

					if(is_bof){
							try{
								pThis->m_pRecordSet->MoveFirst();
								pThis->m_pRecordSet->MovePrev();
							}catch(CDaoException* f){
								f->Delete();
							}																		
					}else if(is_eof){
							try{
								pThis->m_pRecordSet->MoveLast();
								pThis->m_pRecordSet->MoveNext();
							}catch(CDaoException* f){
								f->Delete();
							}																		
					}else{
							try{
								pThis->m_pRecordSet->SetBookmark(bok);
							}catch(CDaoException* f){
								f->Delete();
							}									
					}
					}
					break;
	}
	return S_OK;
}    
    
HRESULT CRowCursor::XCURSORSCROLL::GetApproximateCount(LARGE_INTEGER *pdlApproxCount, DWORD *pdwFullyPopulated)
{
	METHOD_PROLOGUE(CRowCursor, CURSORSCROLL)
	*pdwFullyPopulated = DBCURSORPOPULATED_PARTIALLY;
	pdlApproxCount->HighPart = 0;
	pdlApproxCount->LowPart = pThis->m_pRecordSet->GetRecordCount();
	return S_OK;
}

////////////////////////////////
//
// ICursorFind Implementation
//

HRESULT CRowCursor::XCURSORFIND::QueryInterface(REFIID riid, LPVOID FAR *ppvObj)
{
	METHOD_PROLOGUE(CRowCursor, CURSORFIND)
	return (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
}

ULONG CRowCursor::XCURSORFIND::AddRef()
{
	METHOD_PROLOGUE(CRowCursor, CURSORFIND)
	return (ULONG)pThis->ExternalAddRef();
}

ULONG CRowCursor::XCURSORFIND::Release()
{
	METHOD_PROLOGUE(CRowCursor, CURSORFIND)
	return (ULONG)pThis->ExternalRelease();
}

HRESULT CRowCursor::XCURSORFIND::FindByValues(ULONG cbBookmark, LPVOID pBookmark,DWORD dwFindFlags, ULONG cValues, 
DBCOLUMNID rgColumns[], DBVARIANT rgValues[], DWORD rgdwSeekFlags[], DBFETCHROWS * pFetchParams)
{
	METHOD_PROLOGUE(CRowCursor, CURSORFIND)
   
   //sorry, I don't know at this moment how to implement this for table type recordsets
   if(pThis->m_pControl->m_openType==0)return E_NOTIMPL;
      
#ifdef GO_BACK
COleVariant back;
try{
   back=pThis->m_pRecordSet->GetBookmark();
}catch(CDaoException *e){
	e->Delete();
}
#endif
      
      
   HRESULT hr;
	DWORD dwEventWhat;
	ULONG cReasons;
	DBNOTIFYREASON rgReasons[1];
	SAFEARRAY *psa;

   if(pFetchParams)pFetchParams->cRowsReturned = 0;
   if(!pBookmark){
      pBookmark=(void*)&DBBMK_CURRENT;
      cbBookmark=DB_BMK_SIZE;
   }
	// Initialize arguments for notifications
	VariantInit((VARIANT *)&rgReasons[0].arg1);
	VariantInit((VARIANT *)&rgReasons[0].arg2);

	pThis->MakeSafearrayFromBuf(&psa, pBookmark, cbBookmark);
	
	// Bookmarks should be passed in like this to allow VariantChangeByte
	// to be used by bound controls
	rgReasons[0].arg1.vt = VT_ARRAY|VT_UI1;
	rgReasons[0].arg1.parray = psa;
	
	rgReasons[0].arg2.vt = VT_I4;
	rgReasons[0].arg2.lVal = (long)0;

	dwEventWhat = DBEVENT_CURRENT_ROW_CHANGED;
	cReasons = 1;
	rgReasons[0].dwReason = DBREASON_MOVE/*DBREASON_FIND*/;
	rgReasons[0].arg1.parray = psa;			
	hr = pThis->DoNotifyBefore(dwEventWhat, cReasons, rgReasons);
	if (hr != S_OK){
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}
	if(pThis->m_bUpdateInProgress){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		return DB_E_STATEERROR;
	}

   if(!(dwFindFlags & DBFINDFLAGS_FINDNEXT) && !(dwFindFlags & DBFINDFLAGS_FINDPRIOR)){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		return E_INVALIDARG;
   }
	
	// Find which bookmark was passed in and adjust current row to it
	switch (*((BYTE*)pBookmark)){
		case DBBMK_BEGINNING:	//newRowPos = 0;
			      if(pThis->m_pRecordSet->IsBOF())break;
					try{
						pThis->m_pRecordSet->MoveFirst();
						pThis->m_pRecordSet->MovePrev();
					}catch(CDaoException *e){
						e->Delete();
					}
					break;

		case DBBMK_CURRENT:
               break;

		case DBBMK_END:			
				if(pThis->m_pRecordSet->IsEOF())break;
				try{
					pThis->m_pRecordSet->MoveLast();
					pThis->m_pRecordSet->MoveNext();
				}catch(CDaoException *e){
					e->Delete();
				}
				break;
		case DBBMK_INVALID:	
			pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			VariantClear((VARIANT *)&rgReasons[0].arg1);
			VariantClear((VARIANT *)&rgReasons[0].arg2);
			return DB_E_BADBOOKMARK;

		// If no standard bookmarks were passed in, then call function to 
		// get and move current position to bookmark position
		default:
				{				
				VARIANT var;
            VariantInit(&var);
				var.vt=VT_ARRAY|VT_UI1;
				//BSTR b=::SysAllocStringLen((BSTR)((char*)pBookmark+1),cbBookmark-1);
				//::VectorFromBstr((BSTR)((char*)pBookmark+1),&var.parray);
				//::SysFreeString(b);	
            void *ppvData;
            SAFEARRAYBOUND rgsabound[1];	
            rgsabound[0].lLbound = 0;
	         rgsabound[0].cElements = cbBookmark-1;            
            var.parray=SafeArrayCreate(VT_UI1,1,rgsabound);
            SafeArrayAccessData(var.parray,&ppvData);
            memcpy(ppvData,((char*)pBookmark)+1,cbBookmark-1),
            SafeArrayUnaccessData(var.parray);

            COleVariant pas(var);
 				try{
				   pThis->m_pRecordSet->SetBookmark(pas);
				}catch(CDaoException* e){
					e->Delete();
               ::SafeArrayDestroy(var.parray);
				   pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
				   VariantClear((VARIANT *)&rgReasons[0].arg1);
				   VariantClear((VARIANT *)&rgReasons[0].arg2);
 					return DB_E_BADBOOKMARK;
				}
				::SafeArrayDestroy(var.parray);
			}
			break;
	}
  
	// Adjust new row position if before first row or after last row

   if(dwFindFlags & DBFINDFLAGS_INCLUDECURRENT){
      try{
         if(dwFindFlags & DBFINDFLAGS_FINDNEXT){
            if(!pThis->m_pRecordSet->IsBOF())pThis->m_pRecordSet->MovePrev();
         }else{
            if(!pThis->m_pRecordSet->IsEOF())pThis->m_pRecordSet->MoveNext();
         }
      }catch(CDaoException *e){
		   e->Delete();
	   }
   }

   CString search="";
   CString comp;
   CString value;
   for(unsigned int i=0;i<cValues;i++){
      switch(rgdwSeekFlags[i]){
         case DBSEEK_LT:
            comp="<";
            break;
         case DBSEEK_LE:
            comp="<=";
            break;
         case DBSEEK_GE:
            comp=">=";
            break;
         case DBSEEK_GT:
            comp=">";
            break;
         case DBSEEK_EQ:
            comp="=";
            break;
         case DBSEEK_PARTIALEQ:
            comp=" LIKE ";
            break;
         default:
			   pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
			   VariantClear((VARIANT *)&rgReasons[0].arg1);
			   VariantClear((VARIANT *)&rgReasons[0].arg2);
            return E_INVALIDARG;
      }
      value=pThis->ConvertToString(rgValues[i]);
      search+=pThis->m_pControl->m_pMetaData[rgColumns[i].lNumber].szColName;
      if(rgdwSeekFlags[i]==DBSEEK_PARTIALEQ){
         int len=value.GetLength();
         if(value[len-1]=='\'')value=value.Left(len-1);
         value+="*";
         if(value[0]=='\'')value+="'";
      }
      search+=comp+value;
      if(i<cValues-1)search+=" AND ";
  }

   try{
      BOOL res=FALSE;
      if(pThis->m_pRecordSet->IsEOF() && !(dwFindFlags & DBFINDFLAGS_FINDNEXT))
         res=pThis->m_pRecordSet->FindLast(search);
      else if(pThis->m_pRecordSet->IsBOF() && (dwFindFlags & DBFINDFLAGS_FINDNEXT))
         res=pThis->m_pRecordSet->FindFirst(search);
      else
         res=pThis->m_pRecordSet->Find(((dwFindFlags & DBFINDFLAGS_FINDNEXT) ? AFX_DAO_NEXT : AFX_DAO_PREV),search);

      if(!res){
   		//pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
         SafeArrayDestroy(psa);
         pThis->MakeSafearrayFromBok(&psa);
	      rgReasons[0].arg1.parray = psa;
	      rgReasons[0].arg2.lVal = (long)0;		
	      
         pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);	   	
         VariantClear((VARIANT *)&rgReasons[0].arg1);
		   VariantClear((VARIANT *)&rgReasons[0].arg2);
                  #ifdef GO_BACK
                  try{
                     pThis->m_pRecordSet->SetBookmark(back);
                  }catch(CDaoException *e){
	                  e->Delete();
                  }
                  #endif
         return DB_S_ENDOFRESULTSET;
      }
   }catch(CDaoException *e){
         e->Delete();
         SafeArrayDestroy(psa);
         pThis->MakeSafearrayFromBok(&psa);
	      rgReasons[0].arg1.parray = psa;
	      rgReasons[0].arg2.lVal = (long)0;		

         pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);	   	
         VariantClear((VARIANT *)&rgReasons[0].arg1);
		   VariantClear((VARIANT *)&rgReasons[0].arg2);
                  #ifdef GO_BACK
                  try{
                     pThis->m_pRecordSet->SetBookmark(back);
                  }catch(CDaoException *f){
	                  f->Delete();
                  }
                  #endif
         return DB_S_ENDOFRESULTSET;
   }

   // If they also requested rows, then get the rows.
	if(pFetchParams && pFetchParams->cRowsRequested){
		hr=pThis->FetchRows(pFetchParams);
	// If we didn't fetch rows, then check for BOF or EOF and return DB_S_ENDOFCURSOR
	}else if(pThis->m_pRecordSet->IsBOF() || pThis->m_pRecordSet->IsEOF())hr = DB_S_ENDOFCURSOR;
	else hr = S_OK;


	// If the event failed in any way, then call FailedToDo
	if (FAILED(hr)){
		pThis->CallFailedToDo(dwEventWhat, cReasons, rgReasons);
		VariantClear((VARIANT *)&rgReasons[0].arg1);
		VariantClear((VARIANT *)&rgReasons[0].arg2);
		return hr;
	}

	
	SafeArrayDestroy(psa);
   pThis->MakeSafearrayFromBok(&psa);
	rgReasons[0].arg1.parray = psa;
	rgReasons[0].arg2.lVal = (long)0;		
	pThis->DoNotifyAfter(dwEventWhat, cReasons, rgReasons);
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	return hr;
}

void CRowCursor::Close()
{
if(!m_pRecordSet || !m_pRecordSet->IsOpen())return;
if(m_bUpdateInProgress)m_xCursorUpdateARow.Cancel();
HRESULT hr;
DBNOTIFYREASON rgReasons[1];
VariantInit((VARIANT *)&rgReasons[0].arg1);
VariantInit((VARIANT *)&rgReasons[0].arg2);

if(!is_clone){
rgReasons[0].dwReason = DBREASON_CLOSE;
hr = DoNotifyBefore(DBEVENT_SET_OF_COLUMNS_CHANGED, 1, rgReasons);
if (hr != S_OK){
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	return;
}



if (FAILED(hr)){
	CallFailedToDo(DBEVENT_SET_OF_COLUMNS_CHANGED, 1, rgReasons);
	VariantClear((VARIANT *)&rgReasons[0].arg1);
	VariantClear((VARIANT *)&rgReasons[0].arg2);
	return;
}
}

if(clone)clone->Close();
if (m_pDataColumns){
	CoTaskMemFree ((LPVOID)m_pDataColumns);
	m_pDataColumns=NULL;
	m_nExistingDataCols=0;
	m_lRowDataLength=0;
}
if(m_pColCursor)m_pColCursor->Close();

m_pRecordSet->Close();
if(!is_clone){
   DoNotifyAfter(DBEVENT_SET_OF_COLUMNS_CHANGED, 1, rgReasons);
   VariantClear((VARIANT *)&rgReasons[0].arg1);
   VariantClear((VARIANT *)&rgReasons[0].arg2);
}
}

void CRowCursor::MakeSafearrayFromBok(SAFEARRAY FAR * FAR *ppbBookmark)
{
unsigned long int cBookmark;
void *pbBookmark=NULL;
if(m_pRecordSet->IsEOF()){ //Return End
	cBookmark = DB_BMK_SIZE;
	pbBookmark=(void*)&DBBMK_END;
	MakeSafearrayFromBuf(ppbBookmark, pbBookmark, cBookmark);
}else if(m_pRecordSet->IsBOF()){ //Return Beginning
	cBookmark = DB_BMK_SIZE;
	pbBookmark=(void*)&DBBMK_BEGINNING;
	MakeSafearrayFromBuf(ppbBookmark, pbBookmark, cBookmark);
}else{
	// Calculate the new bookmark based on our handy formula
	COleVariant bok;
	try{
		bok=m_pRecordSet->GetBookmark();
		//BSTR bstr;
		//BstrFromVector(((LPCVARIANT)bok)->parray,(unsigned short**)&bstr);
      void *ppvData;
      SafeArrayAccessData(((LPCVARIANT)bok)->parray,&ppvData);
      long int ub,lb;
      SafeArrayGetUBound(((LPCVARIANT)bok)->parray,1,&ub);
      SafeArrayGetLBound(((LPCVARIANT)bok)->parray,1,&lb);
      cBookmark=ub-lb+1;		
      pbBookmark=CoTaskMemAlloc(cBookmark);
      //pbBookmark=CoTaskMemAlloc(1+(lstrlenW(bstr)+1)*2);
		*(char*)pbBookmark='a';
		//cBookmark = /*1+(lstrlenW(bstr)+1)*2*/1+len;
		memcpy(((char*)pbBookmark)+1,(char*)ppvData,cBookmark-1);
		//::SysFreeString(bstr);
      SafeArrayUnaccessData(((LPCVARIANT)bok)->parray);
		MakeSafearrayFromBuf(ppbBookmark, pbBookmark, cBookmark);
		CoTaskMemFree(pbBookmark);
   }catch(CDaoException *e){
		e->Delete();
		cBookmark = DB_BMK_SIZE;
		if(m_pRecordSet->IsDeleted()){
				pbBookmark=(void*)&DBBMK_CURRENT;						
		}else
				pbBookmark=(void*)&DBBMK_INVALID;					
		MakeSafearrayFromBuf(ppbBookmark, pbBookmark, cBookmark);
	}
}		
}

BOOL CRowCursor::IsOpen()
{
if(!m_pRecordSet)return FALSE;
return m_pRecordSet->IsOpen(); 
}

CString CRowCursor::ConvertToString(DBVARIANT value)
{
CString result;
USES_CONVERSION;
switch(value.vt){
   case VT_I1:
   case VT_UI1:
   case VT_BOOL:
      result.Format("%d",value.bVal);
      break;
   case VT_UI2:
   case VT_I2:
      result.Format("%d",value.iVal);
      break;
   case VT_INT:
   case VT_UINT:
   case VT_UI4:
   case VT_I4:
   case VT_I8: 
   case VT_UI8:
   case VT_DECIMAL:
      result.Format("%ld",value.lVal);
      break;
   case VT_R4:
      result.Format("%f",value.fltVal);
      break;	
   case VT_R8:
      result.Format("%f",value.dblVal);
      break;	
   case VT_CY:{
         COleCurrency cur(value.cyVal);
         result=CString("'")+cur.Format()+"'";
      }
      break;
   case VT_DATE:{
         COleDateTime cur(value.date);
         result=CString("'")+cur.Format()+"'";
      }
      break;
   case VT_BSTR:
   case VT_PTR:
   case VT_LPSTR:
      result=CString("'")+CString(value.bstrVal)+"'";
      break;
   case VT_LPWSTR:
      result=CString("'")+CString(W2T(value.bstrVal))+"'";
      break;
}
return result;
}





LPCONNECTIONPOINT CRowCursor::GetConnectionHook(REFIID iid)
{
   if(!IsEqualIID((IID)iid, IID_INotifyDBEvents))return NULL;	
	return (LPCONNECTIONPOINT)((char*)&m_xConnectionPoint +
				offsetof(CConnectionPoint, m_xConnPt));
}
