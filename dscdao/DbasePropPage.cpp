// DbasePropPage.cpp : implementation file
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
#include "DbasePropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDbasePropPage dialog

IMPLEMENT_DYNCREATE(CDbasePropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CDbasePropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CDbasePropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {E2A25D83-6369-11D2-BD24-0000216A06C9}
IMPLEMENT_OLECREATE_EX(CDbasePropPage, "dscdao.CDbasePropPage",
	0xe2a25d83, 0x6369, 0x11d2, 0xbd, 0x24, 0x0, 0x0, 0x21, 0x6a, 0x6, 0xc9)


/////////////////////////////////////////////////////////////////////////////
// CDbasePropPage::CDbasePropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CDbasePropPage

BOOL CDbasePropPage::CDbasePropPageFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_DBASE_TYPENAME);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CDbasePropPage::CDbasePropPage - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CDbasePropPage::CDbasePropPage() :
	COlePropertyPage(IDD, IDS_DBASE_CAPTION)
{
	//{{AFX_DATA_INIT(CDbasePropPage)
	m_Exclusive = FALSE;
	m_password = _T("");
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CDbasePropPage::DoDataExchange - Moves data between page and properties

void CDbasePropPage::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CDbasePropPage)
	DDP_Check(pDX, IDC_CHECK1, m_Exclusive, _T("OpenExclusive") );
	DDX_Check(pDX, IDC_CHECK1, m_Exclusive);
	DDP_Text(pDX, IDC_EDIT1, m_password, _T("Password") );
	DDX_Text(pDX, IDC_EDIT1, m_password);
	DDV_MaxChars(pDX, m_password, 50);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CDbasePropPage message handlers
