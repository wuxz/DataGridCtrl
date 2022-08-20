#if !defined(AFX_GENPROPPAGE_H__E6581864_8083_11D2_BE57_0000216A06C9__INCLUDED_)
#define AFX_GENPROPPAGE_H__E6581864_8083_11D2_BE57_0000216A06C9__INCLUDED_


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
// GenPropPage.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGenPropPage : Property page dialog

class CGenPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CGenPropPage)
	DECLARE_OLECREATE_EX(CGenPropPage)

// Constructors
public:
	CGenPropPage();

// Dialog Data
	//{{AFX_DATA(CGenPropPage)
	enum { IDD = IDD_PROPPAGE_GEN };
	BOOL	m_ShowAddnewButton;
	BOOL	m_ShowCancelButton;
	BOOL	m_ShowUpdateButton;
	int		m_Appearance;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CGenPropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GENPROPPAGE_H__E6581864_8083_11D2_BE57_0000216A06C9__INCLUDED_)
