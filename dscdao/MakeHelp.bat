@echo off
REM -- First, make map file from resource.h
echo // MAKEHELP.BAT generated Help Map file.  Used by DSCDAO.HPJ. > "hlp\dscdao.hm"
echo. >>hlp\dscdao.hm
echo // Commands (ID_* and IDM_*) >> "hlp\dscdao.hm"
makehm ID_,HID_,0x10000 IDM_,HIDM_,0x10000 resource.h >> "hlp\dscdao.hm"
echo. >> "hlp\dscdao.hm"
echo // Prompts (IDP_*) >> "hlp\dscdao.hm"
makehm IDP_,HIDP_,0x30000 resource.h >> "hlp\dscdao.hm"
echo. >> "hlp\dscdao.hm"
echo // Resources (IDR_*) >> "hlp\dscdao.hm"
makehm IDR_,HIDR_,0x20000 resource.h >> "hlp\dscdao.hm"
echo. >> "hlp\dscdao.hm"
echo // Dialogs (IDD_*) >> "hlp\dscdao.hm"
makehm IDD_,HIDD_,0x20000 resource.h >> "hlp\dscdao.hm"
echo. >> "hlp\dscdao.hm"
echo // Frame Controls (IDW_*) >> "hlp\dscdao.hm"
makehm IDW_,HIDW_,0x50000 resource.h >> "hlp\dscdao.hm"
REM -- Make help for Project DSCDAO
start /wait hcw /C /E /M "dscdao.hpj"
echo.
