#if !defined(AFX_DSCDAO_H__64D3F3DC_494D_11D2_BC7D_0000216A06C9__INCLUDED_)
#define AFX_DSCDAO_H__64D3F3DC_494D_11D2_BC7D_0000216A06C9__INCLUDED_

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

// dscdao.h : main header file for DSCDAO.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CDscdaoApp : See dscdao.cpp for implementation.

class CDscdaoApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DSCDAO_H__64D3F3DC_494D_11D2_BC7D_0000216A06C9__INCLUDED)
