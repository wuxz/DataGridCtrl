// GenPropPage.cpp : implementation file
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
#include "dscdao.h"
#include "GenPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGenPropPage dialog

IMPLEMENT_DYNCREATE(CGenPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CGenPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CGenPropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {E6581863-8083-11D2-BE57-0000216A06C9}
IMPLEMENT_OLECREATE_EX(CGenPropPage, "dscdao.CGenPropPage",
	0xe6581863, 0x8083, 0x11d2, 0xbe, 0x57, 0x0, 0x0, 0x21, 0x6a, 0x6, 0xc9)


/////////////////////////////////////////////////////////////////////////////
// CGenPropPage::CGenPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CGenPropPage

BOOL CGenPropPage::CGenPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_GEN_TYPENAME);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CGenPropPage::CGenPropPage - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CGenPropPage::CGenPropPage() :
	COlePropertyPage(IDD, IDS_GEN_CAPTION)
{
	//{{AFX_DATA_INIT(CGenPropPage)
	m_ShowAddnewButton = FALSE;
	m_ShowCancelButton = FALSE;
	m_ShowUpdateButton = FALSE;
	m_Appearance = 0;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CGenPropPage::DoDataExchange - Moves data between page and properties

void CGenPropPage::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CGenPropPage)
	DDP_Check(pDX, IDC_CHECK2, m_ShowAddnewButton, _T("ShowAddnewButton") );
	DDX_Check(pDX, IDC_CHECK2, m_ShowAddnewButton);
	DDP_Check(pDX, IDC_CHECK3, m_ShowCancelButton, _T("ShowCancelButton") );
	DDX_Check(pDX, IDC_CHECK3, m_ShowCancelButton);
	DDP_Check(pDX, IDC_CHECK4, m_ShowUpdateButton, _T("ShowUpdateButton") );
	DDX_Check(pDX, IDC_CHECK4, m_ShowUpdateButton);
	DDP_Text(pDX, IDC_EDIT1, m_Appearance, _T("Appearance") );
	DDX_Text(pDX, IDC_EDIT1, m_Appearance);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CGenPropPage message handlers
