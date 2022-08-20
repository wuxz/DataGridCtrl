// dscdao.cpp : Implementation of CDscdaoApp and DLL registration.

#include "stdafx.h"
#include "dscdao.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


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


CDscdaoApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0x64d3f3d3, 0x494d, 0x11d2, { 0xbc, 0x7d, 0, 0, 0x21, 0x6a, 0x6, 0xc9 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;


#ifdef _DEBUG    
CMemoryState oldMemState, newMemState, diffMemState;
#endif


////////////////////////////////////////////////////////////////////////////
// CDscdaoApp::InitInstance - DLL initialization

BOOL CDscdaoApp::InitInstance()
{
#ifdef _DEBUG
  oldMemState.Checkpoint();
#endif
  return COleControlModule::InitInstance();
}


////////////////////////////////////////////////////////////////////////////
// CDscdaoApp::ExitInstance - DLL termination

int CDscdaoApp::ExitInstance()
{
	// TODO: Add your own module termination code here.
   AfxDaoTerm();
#ifdef _DEBUG
   int res=COleControlModule::ExitInstance();
   newMemState.Checkpoint();
   if( diffMemState.Difference( oldMemState, newMemState ) ) {
      TRACE( "\nObjects still allocated:\n" );    
      diffMemState.DumpAllObjectsSince();
   }
   return res;
#else
   return COleControlModule::ExitInstance();
#endif
}


/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}


/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
