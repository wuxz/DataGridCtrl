#if !defined(AFX_DSCDAOPPG_H__64D3F3E6_494D_11D2_BC7D_0000216A06C9__INCLUDED_)
#define AFX_DSCDAOPPG_H__64D3F3E6_494D_11D2_BC7D_0000216A06C9__INCLUDED_

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

// DscdaoPpg.h : Declaration of the CDscdaoPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CDscdaoPropPage : See DscdaoPpg.cpp.cpp for implementation.

class CDscdaoPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CDscdaoPropPage)
	DECLARE_OLECREATE_EX(CDscdaoPropPage)

// Constructor
public:
	CDscdaoPropPage();

// Dialog Data
	//{{AFX_DATA(CDscdaoPropPage)
	enum { IDD = IDD_PROPPAGE_DSCDAO };
	CEdit	m_Edit2;
	CComboBox	m_Combo4;
	CComboBox	m_Combo3;
	CButton	m_Check2;
	CButton	m_Check1;
	CComboBox	m_Index;
	CComboBox	m_Combo;
	CString	m_databasePath;
	CString	m_Table;
	CString	m_indexName;
	BOOL	m_ReadOnly;
	int		m_openType;
	int		m_sqlType;
	CString	m_SqlString;
	BOOL	m_lockingMode;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CDscdaoPropPage)
	afx_msg void OnButton1();
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeCombo3();
	afx_msg void OnSelchangeCombo4();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DSCDAOPPG_H__64D3F3E6_494D_11D2_BC7D_0000216A06C9__INCLUDED)
