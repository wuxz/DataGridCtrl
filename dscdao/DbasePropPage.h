#if !defined(AFX_DBASEPROPPAGE_H__E2A25D84_6369_11D2_BD24_0000216A06C9__INCLUDED_)
#define AFX_DBASEPROPPAGE_H__E2A25D84_6369_11D2_BD24_0000216A06C9__INCLUDED_

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
// DbasePropPage.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDbasePropPage : Property page dialog

class CDbasePropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CDbasePropPage)
	DECLARE_OLECREATE_EX(CDbasePropPage)

// Constructors
public:
	CDbasePropPage();

// Dialog Data
	//{{AFX_DATA(CDbasePropPage)
	enum { IDD = IDD_PROPPAGE_DBASE };
	BOOL	m_Exclusive;
	CString	m_password;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CDbasePropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DBASEPROPPAGE_H__E2A25D84_6369_11D2_BD24_0000216A06C9__INCLUDED_)
