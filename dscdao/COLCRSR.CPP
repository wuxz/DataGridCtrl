// ColCrsr.cpp : implementation file
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
// CColCursor Implementation

IMPLEMENT_DYNCREATE(CColCursor, CCmdTarget)

/////////////////////////////////////////////////////////////////////////////
// Interface map

BEGIN_INTERFACE_MAP(CColCursor, CCmdTarget)
	INTERFACE_PART(CColCursor, IID_ICursor, COLUMNSCURSOR)
	INTERFACE_PART(CColCursor, IID_ICursorMove, COLUMNSCURSORMOVE)
END_INTERFACE_MAP()

CColCursor::CColCursor(CDscdaoCtrl *pControl, CRowCursor *pRowCursor)
{
	m_pControl = pControl;
	m_pRowCursor = pRowCursor;
	m_pMetaDataColumns = NULL;
	m_nExistingMetaDataCols = 0;
	m_lMetaDataRowLength = 0;
	m_CurrentMetaDataRow = 0;
}

CColCursor::~CColCursor()
{
if (m_pMetaDataColumns){
	CoTaskMemFree ((LPVOID)m_pMetaDataColumns);
//	m_pMetaDataColumns=NULL;
}
}

/////////////////////////////////////////////////////////////////////////////
// ICursor Implementation for the Columns Cursor

HRESULT CColCursor::XCOLUMNSCURSOR::GetColumnsCursor(REFIID /*riid*/,IUnknown ** /*ppvColumnsCursor*/,ULONG * /*pcRows*/)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
   return E_NOTIMPL;
}

HRESULT CColCursor::XCOLUMNSCURSOR::SetBindings(ULONG cCol, DBCOLUMNBINDING rgBoundColumns[],ULONG cbRowLength,DWORD dwFlags)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	// Validate Column Bindings. There is no functionality in this
	// function
	HRESULT hr = pThis->m_pControl->ValidateColumnBindings(cCol, rgBoundColumns);
   if(FAILED(hr))return hr;
	// Do the actual setting of the column bindings
	hr = pThis->m_pControl->BindColumns(&pThis->m_pMetaDataColumns, &pThis->m_nExistingMetaDataCols, cCol, rgBoundColumns, dwFlags);
   if(FAILED(hr))return hr;
	// If cbRowLength is 0, then we need to find the minimum number of bytes 
	// required by the new column bindings. Currently not supported. Add code
	// to FindDataRowLength to support this feature.
	if (!cbRowLength)pThis->m_pControl->FindDataRowLength();
	else pThis->m_lMetaDataRowLength = cbRowLength;
	return hr;
}

HRESULT CColCursor::XCOLUMNSCURSOR::GetBindings(ULONG *pcCol,DBCOLUMNBINDING *prgBoundColumns[],ULONG *pcbRowLength)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	if (!pcCol || !prgBoundColumns)return E_FAIL;
	*pcCol = pThis->m_nExistingMetaDataCols;
	if (!(*pcCol)){
		// Allocate space to store return current bindings
		*prgBoundColumns = (DBCOLUMNBINDING *)CoTaskMemAlloc(sizeof(DBCOLUMNBINDING) * (*pcCol)); //new DBCOLUMNBINDING[*pcCol];
			if (*prgBoundColumns){
			// Copy all bindings into the newly allocated memory.
			for (ULONG i = 0; i < *pcCol; i++)
				(*prgBoundColumns)[i] = pThis->m_pMetaDataColumns[i];
		}else	return E_OUTOFMEMORY;
	}else	*prgBoundColumns = NULL;
	if(pcbRowLength)*pcbRowLength = pThis->m_lMetaDataRowLength;
	return S_OK;
}

HRESULT CColCursor::XCOLUMNSCURSOR::GetNextRows(LARGE_INTEGER udlRowsToSkip, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	pThis->m_CurrentMetaDataRow += (udlRowsToSkip.LowPart + 1);
	if (pThis->m_CurrentMetaDataRow > pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks + 1)
		pThis->m_CurrentMetaDataRow = pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks + 1;
	return pThis->FetchRows(pFetchParams);
}

HRESULT CColCursor::XCOLUMNSCURSOR::Requery()
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	// Reset current row pointer to the BOF. Cursors would normally populate
	// cache with data from backend at this point as well.
	pThis->m_CurrentMetaDataRow=0;
	if (pThis->m_pMetaDataColumns){
		CoTaskMemFree((LPVOID)pThis->m_pMetaDataColumns);
		pThis->m_pMetaDataColumns=NULL;
	}
	pThis->m_pControl->GetMetaData();
	return S_OK;
}

HRESULT CColCursor::XCOLUMNSCURSOR::QueryInterface(REFIID riid, LPVOID FAR *ppvObj)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	return (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
}

ULONG CColCursor::XCOLUMNSCURSOR::AddRef()
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	return (ULONG)pThis->ExternalAddRef();
}

ULONG CColCursor::XCOLUMNSCURSOR::Release()
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSOR)
	return (ULONG)pThis->ExternalRelease();
}

/////////////////////////////////////////////////////////////////////////////
// ICursorMove Implementation for the Columns Cursor

STDMETHODIMP CColCursor::XCOLUMNSCURSORMOVE::GetBindings(ULONG *pcCol, DBCOLUMNBINDING *prgBoundColumns[], ULONG *pcbRowLength)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	// delegate to ICursor::GetBindings()
	return pThis->m_xCOLUMNSCURSOR.GetBindings(pcCol, prgBoundColumns, pcbRowLength);
}	

STDMETHODIMP CColCursor::XCOLUMNSCURSORMOVE::GetColumnsCursor(REFIID riid, IUnknown *ppunkColumnsCursor[], ULONG *pcRows)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	// delegate to ICursor::GetColumnsCursor()
    return pThis->m_xCOLUMNSCURSOR.GetColumnsCursor(riid, ppunkColumnsCursor, pcRows);
}	

STDMETHODIMP CColCursor::XCOLUMNSCURSORMOVE::GetNextRows(LARGE_INTEGER dlRowsToSkip, DBFETCHROWS *pFetchParams)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	// delegate to ICursor::GetNextRows()
	return pThis->m_xCOLUMNSCURSOR.GetNextRows(dlRowsToSkip, pFetchParams);
}	

STDMETHODIMP CColCursor::XCOLUMNSCURSORMOVE::Requery()
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	// delegate to ICursor::Requery()
	return pThis->m_xCOLUMNSCURSOR.Requery();
}	

STDMETHODIMP CColCursor::XCOLUMNSCURSORMOVE::SetBindings(ULONG cCol, DBCOLUMNBINDING rgBoundColumns[], ULONG cbRowLength, DWORD dwFlags)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	// delegate to ICursor::SetBindings()
	return pThis->m_xCOLUMNSCURSOR.SetBindings(cCol, rgBoundColumns, cbRowLength, dwFlags);
}	

HRESULT CColCursor::XCOLUMNSCURSORMOVE::Move(ULONG /*cbBookmark*/, void *pBookmark, LARGE_INTEGER dlOffset, DBFETCHROWS *pFetchParams)
{
	ULONG newRowPos = 0;
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	if (*(BYTE *)pBookmark == DBBMK_BEGINNING){
		newRowPos = 0;
		newRowPos += dlOffset.LowPart;
	}else if (*(BYTE *)pBookmark == DBBMK_CURRENT){
		newRowPos = pThis->m_CurrentMetaDataRow;
		newRowPos += dlOffset.LowPart;
	}else if (*(BYTE *)pBookmark == DBBMK_END){
		newRowPos = pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks + 1;
		newRowPos += dlOffset.LowPart;
	}else if (*(BYTE *)pBookmark == DBBMK_INVALID)return DB_E_BADBOOKMARK;
	else return DB_E_BADBOOKMARK;
	// adjust new row position if before first row or after last row
	if (newRowPos & 0x80000000) newRowPos = 0;
	if (newRowPos > pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks)
		newRowPos = pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks + 1;
	pThis->m_CurrentMetaDataRow = newRowPos;
	if (pFetchParams && pFetchParams->cRowsRequested){
		if (!newRowPos)pThis->m_CurrentMetaDataRow++;
		return pThis->FetchRows(pFetchParams);
	}else if (!newRowPos || (newRowPos == (pThis->m_pControl->m_nColumns + pThis->m_pControl->m_nBookmarks + 1)))
		return DB_S_ENDOFCURSOR;
	return S_OK;
}

HRESULT CColCursor::XCOLUMNSCURSORMOVE::GetBookmark(DBCOLUMNID * /*pBookmarkType*/, ULONG /*cbMaxSize*/, ULONG * /*pcbBookmark*/, void * /*pBookmark*/)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
   return E_NOTIMPL;
}

HRESULT CColCursor::XCOLUMNSCURSORMOVE::Clone(DWORD /*dwFlags*/, REFIID /*riid*/, IUnknown ** /*ppvClonedCOLUMNSCURSOR*/)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
   return E_NOTIMPL;
}

HRESULT CColCursor::XCOLUMNSCURSORMOVE::QueryInterface(REFIID riid, LPVOID FAR *ppvObj)
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	return (HRESULT)pThis->ExternalQueryInterface(&riid, ppvObj);
}

ULONG CColCursor::XCOLUMNSCURSORMOVE::AddRef()
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	return (ULONG)pThis->ExternalAddRef();
}

ULONG CColCursor::XCOLUMNSCURSORMOVE::Release()
{
	METHOD_PROLOGUE(CColCursor, COLUMNSCURSORMOVE)
	return (ULONG)pThis->ExternalRelease();
}


////////////////////////////////////////////////////////////
// Helper function Implementations

// This function is called by ICursor::GetNextRows() as well as ICursorMove::Move().
// It fetches the requested number of meta data rows according to the current meta data
// column bindings.

HRESULT CColCursor::FetchRows(DBFETCHROWS *pFetchParams)
{
	// loop counter for rows and columns
	ULONG i=0, j=0;
	// variable to keep track of starting position in out-of-line memory for each column.
	ULONG cbCumulativeOutOfLineMemSize = 0;
	// ptr to the beginning of each row of in-line memory
	BYTE *pDestBytePtr = NULL;
   // ptr to the beginning of a column of data within the current row of in-line memory
	BYTE *pDestMetaColValue = NULL;
	if (!pFetchParams || !pFetchParams->cRowsRequested)return E_FAIL;
	// adjust the # of rows returned if < the # of rows requested
	if (m_CurrentMetaDataRow + pFetchParams->cRowsRequested > m_pControl->m_nColumns + m_pControl->m_nBookmarks + 1)
		pFetchParams->cRowsReturned = m_pControl->m_nColumns + m_pControl->m_nBookmarks - m_CurrentMetaDataRow + 1;
	else
		pFetchParams->cRowsReturned = pFetchParams->cRowsRequested;
	// fail if no rows can be returned!
	if (!pFetchParams->cRowsReturned)return E_FAIL;
	if (pFetchParams->dwFlags & DBROWFETCH_CALLEEALLOCATES){
		// callee or provider allocates the memory
		if (m_lMetaDataRowLength)
			pFetchParams->pData = CoTaskMemAlloc(m_lMetaDataRowLength * pFetchParams->cRowsReturned);
		else // invalid in-line memory row size!
			return E_FAIL;
		if (!pFetchParams->pData)return E_OUTOFMEMORY;
		// determine the size of out-of-line memory needed
		for (i = 0; i < pFetchParams->cRowsReturned; i++)
			for (j = 0; j < m_nExistingMetaDataCols; j++){
				// hold the size of each column of data requiring out-of-line memory
				// For meta data, only COLUMNIDs inside DBVARIANTs and strings require
				// out-of-line memory.
				ULONG VarMemLength = 0;
				if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
					if (m_pMetaDataColumns[j].columnID.lNumber == CID_NUMBER_COLUMNID){
						VarMemLength = sizeof(DBCOLUMNID);
						cbCumulativeOutOfLineMemSize += VarMemLength;
					}
				}
				if (m_pMetaDataColumns[j].columnID.lNumber == CID_NUMBER_NAME){
					if (m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName)
						VarMemLength = wcslen(m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName) + 1;
					else continue;
					if (m_pMetaDataColumns[j].dwDataType == VT_LPWSTR)VarMemLength *= 2;
					cbCumulativeOutOfLineMemSize += VarMemLength;
				}
			}
		// Allocate it!
		pFetchParams->pVarData = CoTaskMemAlloc(cbCumulativeOutOfLineMemSize);
		cbCumulativeOutOfLineMemSize = 0;
		if (!pFetchParams->pVarData)return E_OUTOFMEMORY;
	}else if (pFetchParams->dwFlags != DBROWFETCH_DEFAULT && !(pFetchParams->dwFlags & DBROWFETCH_FORCEREFRESH))
		return DB_E_BADFETCHINFO;
	else{ // caller allocated memory
		// fail if caller did not allocate any in-line memory 
		if (!pFetchParams->pData)return DB_E_BADFETCHINFO;
	}
	// iterate through each bound column of each requested row to fetch the data
	for (i = 0; i < pFetchParams->cRowsReturned; i++){
		pDestBytePtr = (BYTE *)pFetchParams->pData;
		pDestBytePtr += (i * m_lMetaDataRowLength);
		for (j = 0; j < m_nExistingMetaDataCols; j++){
			// proceed only if inline memory contains any data for the column j...
			if (m_pMetaDataColumns[j].obData != DB_NOVALUE){
				// fill the column length field in in-line memory if requested by client
				if (m_pMetaDataColumns[j].obVarDataLen != DB_NOVALUE){
					pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obVarDataLen;
					*(DWORD *)pDestMetaColValue = 0;
				}
				// fill the column info field in in-line memory if requested by client
				if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
					pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
					*(DWORD *)pDestMetaColValue = DB_NOINFO;
				}
				// fill the actual in-line memory data.
				pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obData;
				// variable to hold length of string data that needs out-of-line memory
				ULONG length = 0;
				// available # of bytes of string data, in case it is truncated
				ULONG cbAvailableLength = 0;
				// flag to indicate that the fetched data is truncated
				BOOL bTruncated = FALSE;
				// fetch the meta data according to the COLUMNID requested.
				// Note: only COLUMNIDs of type DBCOLKIND_GUID_NUMBER are
				// supported in this sample.
				switch (m_pMetaDataColumns[j].columnID.lNumber){
				  case CID_NUMBER_COLUMNID:   
					if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						length = sizeof(DBCOLUMNID);
						// check if there is enough caller allocated out-of-line memory
						if (!(pFetchParams->dwFlags & DBROWFETCH_CALLEEALLOCATES)){
							if (!pFetchParams->pVarData || !pFetchParams->cbVarData){
								m_CurrentMetaDataRow += i;
								return DB_E_BADFETCHINFO;
							}
							if (pFetchParams->cbVarData - cbCumulativeOutOfLineMemSize < length){
								m_CurrentMetaDataRow += i;
								return DB_S_BUFFERTOOSMALL;
							}
						}
						// can't coerce COLUMNID to any other data type
						if (m_pMetaDataColumns[j].dwDataType != DBTYPE_COLUMNID){
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;
						}else{
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_COLUMNID;
							((DBVARIANT *)pDestMetaColValue)->pColumnid = (DBCOLUMNID *)((BYTE *)pFetchParams->pVarData + cbCumulativeOutOfLineMemSize);
							*(((DBVARIANT *)pDestMetaColValue)->pColumnid) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].ColumnID;
							cbCumulativeOutOfLineMemSize += length;
						}
					}else if(m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == DBTYPE_COLUMNID)
							(*((DBCOLUMNID *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].ColumnID;
					}
					if(m_pMetaDataColumns[j].dwDataType != DBTYPE_COLUMNID){
						if(m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
						}
					}
					break;
			   case CID_NUMBER_NAME:
					if(m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName == NULL){
						if(m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
							if(m_pMetaDataColumns[j].dwDataType == VT_LPSTR){
								((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_NULL;
								((DBVARIANT *)pDestMetaColValue)->pszVal = NULL;
							}else if(m_pMetaDataColumns[j].dwDataType == VT_LPWSTR){
								((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_NULL;
								((DBVARIANT *)pDestMetaColValue)->pwszVal = NULL;
							}else if(m_pMetaDataColumns[j].dwDataType == VT_BSTR){
								((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_NULL;
								((DBVARIANT *)pDestMetaColValue)->bstrVal = NULL;
							}
						}else if(m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
							*(DWORD *)pDestMetaColValue = NULL;
						
						if(m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_NULL;
						}
						continue;
					}

					if(m_pMetaDataColumns[j].dwDataType == VT_BSTR){
						BSTR bstrSrc = (BSTR)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName;
						if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
							((DBVARIANT *)pDestMetaColValue)->vt = VT_BSTR;
							((DBVARIANT *)pDestMetaColValue)->bstrVal = SysAllocString(bstrSrc);
							if(((DBVARIANT *)pDestMetaColValue)->bstrVal == NULL){
								m_CurrentMetaDataRow += i;
								return E_OUTOFMEMORY;
							}
						}else if(m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
							*((BSTR *)pDestMetaColValue) = SysAllocString(bstrSrc);
							if(!(*(BSTR *)pDestMetaColValue)){
								m_CurrentMetaDataRow += i;
								return E_OUTOFMEMORY;
							}
						}
						break;
					}
					// the Column name is *ALWAYS*  an LPWSTR
					length = wcslen(m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName) + 1;
					// if data type requested by client is a Unicode string...
					if (m_pMetaDataColumns[j].dwDataType == VT_LPWSTR)length *= 2;
					cbAvailableLength = length;
					// truncate data if requested length is less than available length.
					if (m_pMetaDataColumns[j].cbMaxLen != DB_NOMAXLENGTH && length > m_pMetaDataColumns[j].cbMaxLen){
						bTruncated = TRUE;
						length = m_pMetaDataColumns[j].cbMaxLen;
					}

					// check if there is enough caller allocated out-of-line memory
					if (!(pFetchParams->dwFlags & DBROWFETCH_CALLEEALLOCATES)){
						if (!pFetchParams->pVarData || !pFetchParams->cbVarData){
							m_CurrentMetaDataRow += i;
							return DB_E_BADFETCHINFO;
						}
						if (pFetchParams->cbVarData - cbCumulativeOutOfLineMemSize < length){
							m_CurrentMetaDataRow += i;
							return DB_S_BUFFERTOOSMALL;
						}
					}

					// requested data type could be an ANSI or UNICODE string in a
					// DBVARIANT or in direct memory; if so, assign a pointer to the
					// string to in-line memory; if not, can't coerce.
					if(m_pMetaDataColumns[j].dwDataType == VT_LPWSTR){
						if(m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
							((DBVARIANT *)pDestMetaColValue)->vt = VT_LPWSTR;
							((DBVARIANT *)pDestMetaColValue)->pwszVal = (LPWSTR)((BYTE *)pFetchParams->pVarData + cbCumulativeOutOfLineMemSize);
						}else if(m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
							*((LPWSTR *)pDestMetaColValue) = (LPWSTR)((BYTE *)pFetchParams->pVarData + cbCumulativeOutOfLineMemSize);
					}else if(m_pMetaDataColumns[j].dwDataType == VT_LPSTR){
						if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
							((DBVARIANT *)pDestMetaColValue)->vt = VT_LPSTR;
							((DBVARIANT *)pDestMetaColValue)->pszVal = (LPSTR)((BYTE *)pFetchParams->pVarData + cbCumulativeOutOfLineMemSize);
						}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
							*((LPSTR *)pDestMetaColValue) = (LPSTR)((BYTE *)pFetchParams->pVarData + cbCumulativeOutOfLineMemSize);
					}else{
						if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;
						}

						if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
						}
						break;
					}

					// update the current out-of-line memory size for next column that uses it.
					cbCumulativeOutOfLineMemSize += length;
					// Copy the actual string to out-of-line memory.
					if (m_pMetaDataColumns[j].dwDataType == VT_LPWSTR){
						if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT)
							wcsncpy (((DBVARIANT *)pDestMetaColValue)->pwszVal, m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName, length/2);
						else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
							wcsncpy (*((LPWSTR *)pDestMetaColValue), m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName, length/2);
					}else if (m_pMetaDataColumns[j].dwDataType == VT_LPSTR){
						USES_CONVERSION;

						if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT)
							strncpy (((DBVARIANT *)pDestMetaColValue)->pszVal, W2A(m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName), length);
						else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
							strncpy (*((LPSTR *)pDestMetaColValue), W2A(m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].szColName), length);
					}

					// append terminating NULL byte(s) if truncated data.
					if (bTruncated){
						if (m_pMetaDataColumns[j].dwDataType == VT_LPWSTR){
							if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT)
								*((WCHAR *)((BYTE *)(((DBVARIANT *)pDestMetaColValue)->pwszVal) + length - 2)) = (WCHAR)0;
							else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
								*((WCHAR *)(pDestMetaColValue + length - 2)) = (WCHAR)0;
						}else if (m_pMetaDataColumns[j].dwDataType == VT_LPSTR){
							if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT)
								*((CHAR *)(((DBVARIANT *)pDestMetaColValue)->pszVal) + length - 1) = (CHAR)0;
							else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT)
								*((CHAR *)(pDestMetaColValue + length - 1)) = (CHAR)0;
						}

						if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_TRUNCATED;
						}
					}
											
					if (m_pMetaDataColumns[j].obVarDataLen != DB_NOVALUE){
						pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obVarDataLen;
						*(DWORD *)pDestMetaColValue = cbAvailableLength;
					}

					break;

				  case CID_NUMBER_TYPE:		   
						
					if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						((DBVARIANT *)pDestMetaColValue)->vt = (VARTYPE)m_pMetaDataColumns[j].dwDataType;

						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							((DBVARIANT *)pDestMetaColValue)->lVal = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwColType;
						else {
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;
						}
					}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							(*((DWORD *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwColType;
					}

					if (m_pMetaDataColumns[j].dwDataType != VT_I4 || m_pMetaDataColumns[j].dwDataType != VT_I2 || m_pMetaDataColumns[j].dwDataType != VT_R4 || m_pMetaDataColumns[j].dwDataType != VT_R8){
						if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
						}
					}
					break;
				  case CID_NUMBER_BINDTYPE:
					if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						((DBVARIANT *)pDestMetaColValue)->vt = (VARTYPE)m_pMetaDataColumns[j].dwDataType;

						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							((DBVARIANT *)pDestMetaColValue)->lVal = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							((DBVARIANT *)pDestMetaColValue)->iVal = (short int)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							((DBVARIANT *)pDestMetaColValue)->fltVal = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							((DBVARIANT *)pDestMetaColValue)->dblVal = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else {
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;
						}
						}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							(*((DWORD *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							(*((int *)pDestMetaColValue)) = (int)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							(*((float *)pDestMetaColValue)) = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							(*((double *)pDestMetaColValue)) = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwBindType;
					}
							
					if (m_pMetaDataColumns[j].dwDataType != VT_I4 || m_pMetaDataColumns[j].dwDataType != VT_I2 || m_pMetaDataColumns[j].dwDataType != VT_R4 || m_pMetaDataColumns[j].dwDataType != VT_R8){
						if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
						}
					}
					break;

				  case CID_NUMBER_DATACOLUMN:
						
					if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						((DBVARIANT *)pDestMetaColValue)->vt = (VARTYPE)m_pMetaDataColumns[j].dwDataType;

						if (m_pMetaDataColumns[j].dwDataType == VT_BOOL)
							((DBVARIANT *)pDestMetaColValue)->xbool = (VARIANT_BOOL)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_I2)
							((DBVARIANT *)pDestMetaColValue)->iVal = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_I4)
							((DBVARIANT *)pDestMetaColValue)->lVal = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_R4)
							((DBVARIANT *)pDestMetaColValue)->fltVal = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_R8)
							((DBVARIANT *)pDestMetaColValue)->dblVal = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else {
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;

							if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
								pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
								*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
							}
						}
					}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == VT_BOOL)
							(*((BOOL *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							(*((short *)pDestMetaColValue)) = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							(*((long *)pDestMetaColValue)) = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							(*((float *)pDestMetaColValue)) = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							(*((double *)pDestMetaColValue)) = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].bDataCol;
						else{
							if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
								pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
								*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
							}
						}
					}

					break;

				  case CID_NUMBER_MAXLENGTH:
						
					if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						((DBVARIANT *)pDestMetaColValue)->vt = (VARTYPE)m_pMetaDataColumns[j].dwDataType;

						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							((DBVARIANT *)pDestMetaColValue)->lVal = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_I2)
							((DBVARIANT *)pDestMetaColValue)->iVal = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_R4)
							((DBVARIANT *)pDestMetaColValue)->fltVal = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_R8)
							((DBVARIANT *)pDestMetaColValue)->dblVal = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else {
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;

							if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
								pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
								*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
							}
						}
					}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							(*((ULONG *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							(*((short *)pDestMetaColValue)) = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							(*((float *)pDestMetaColValue)) = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							(*((double *)pDestMetaColValue)) = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nColMaxLength;
						else{
							if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE)	{
								pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
								*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
							}
						}
					}
					break;

				  case CID_NUMBER_ENTRYIDMAXLENGTH:
					  if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						((DBVARIANT *)pDestMetaColValue)->vt = (VARTYPE)m_pMetaDataColumns[j].dwDataType;

						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							((DBVARIANT *)pDestMetaColValue)->lVal = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_I2)
							((DBVARIANT *)pDestMetaColValue)->iVal = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_R4)
							((DBVARIANT *)pDestMetaColValue)->fltVal = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else if	(m_pMetaDataColumns[j].dwDataType == VT_R8)
							((DBVARIANT *)pDestMetaColValue)->dblVal = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else {
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;

							if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
								pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
								*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
							}
						}
					}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							(*((ULONG *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							(*((short *)pDestMetaColValue)) = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							(*((float *)pDestMetaColValue)) = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							(*((double *)pDestMetaColValue)) = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].nEntryIDMaxLen;
						else{
							if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
								pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
								*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
							}
						}
					}
	
					break;

				  case CID_NUMBER_UPDATABLE:
						
					if (m_pMetaDataColumns[j].dwBinding == DBBINDING_VARIANT){
						((DBVARIANT *)pDestMetaColValue)->vt = (VARTYPE)m_pMetaDataColumns[j].dwDataType;

						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							((DBVARIANT *)pDestMetaColValue)->lVal = (long)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							((DBVARIANT *)pDestMetaColValue)->iVal = (short)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							((DBVARIANT *)pDestMetaColValue)->fltVal = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							((DBVARIANT *)pDestMetaColValue)->dblVal = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else {
							((DBVARIANT *)pDestMetaColValue)->vt = DBTYPE_HRESULT;
							((DBVARIANT *)pDestMetaColValue)->scode = DB_E_CANTCOERCE;
						}
					}else if (m_pMetaDataColumns[j].dwBinding == DBBINDING_DEFAULT){
						if (m_pMetaDataColumns[j].dwDataType == VT_I4)
							(*((DWORD *)pDestMetaColValue)) = m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else if (m_pMetaDataColumns[j].dwDataType == VT_I2)
							(*((int *)pDestMetaColValue)) = (int)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R4)
							(*((float *)pDestMetaColValue)) = (float)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
						else if (m_pMetaDataColumns[j].dwDataType == VT_R8)
							(*((double *)pDestMetaColValue)) = (double)m_pControl->m_pMetaData[i + m_CurrentMetaDataRow - 1].dwUpdatable;
					}
							
					if (m_pMetaDataColumns[j].dwDataType != VT_I4 || m_pMetaDataColumns[j].dwDataType != VT_I2 || m_pMetaDataColumns[j].dwDataType != VT_R4 || m_pMetaDataColumns[j].dwDataType != VT_R8){
						if (m_pMetaDataColumns[j].obInfo != DB_NOVALUE){
							pDestMetaColValue = pDestBytePtr + m_pMetaDataColumns[j].obInfo;
							*(DWORD *)pDestMetaColValue = DB_CANTCOERCE;
						}
					}
					break;

				  default: 	m_CurrentMetaDataRow += i;
							return E_NOTIMPL;
				}
			}
		}
	}
			
	// update the current row position
	m_CurrentMetaDataRow += (pFetchParams->cRowsReturned - 1);

	// check to see if before the first row or after the last
	if (!m_CurrentMetaDataRow || (m_CurrentMetaDataRow == (m_pControl->m_nColumns + m_pControl->m_nBookmarks + 1)))
		return DB_S_ENDOFCURSOR;
	
	return S_OK;
}

void CColCursor::Close()
{
if (m_pMetaDataColumns){
	CoTaskMemFree ((LPVOID)m_pMetaDataColumns);
}
m_pMetaDataColumns = NULL;
m_nExistingMetaDataCols = 0;
m_lMetaDataRowLength = 0;
m_CurrentMetaDataRow = 0;
}
