; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CGenPropPage
LastTemplate=COlePropertyPage
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "dscdao.h"
CDK=Y

ClassCount=4
Class1=CDscdaoCtrl
Class2=CDscdaoPropPage

ResourceCount=6
Resource1=IDD_PROPPAGE_GEN
LastPage=0
Resource2=IDD_PROPPAGE_DBASE
Class3=CDbasePropPage
Class4=CGenPropPage
Resource3=IDD_PROPPAGE_DSCDAO
Resource4=IDD_PROPPAGE_DBASE (English (U.S.))
Resource5=IDD_PROPPAGE_DSCDAO (English (U.S.))
Resource6=IDD_PROPPAGE_GEN (English (U.S.))

[CLS:CDscdaoCtrl]
Type=0
HeaderFile=DscdaoCtl.h
ImplementationFile=DscdaoCtl.cpp
Filter=W
BaseClass=COleControl
VirtualFilter=wWC
LastObject=CDscdaoCtrl

[CLS:CDscdaoPropPage]
Type=0
HeaderFile=DscdaoPpg.h
ImplementationFile=DscdaoPpg.cpp
Filter=D
LastObject=CDscdaoPropPage
BaseClass=COlePropertyPage
VirtualFilter=idWC

[DLG:IDD_PROPPAGE_DSCDAO]
Type=1
Class=CDscdaoPropPage
ControlCount=14
Control1=IDC_EDIT1,edit,1350631552
Control2=IDC_BUTTON1,button,1342242816
Control3=IDC_COMBO4,combobox,1344339971
Control4=IDC_COMBO3,combobox,1344339971
Control5=IDC_COMBO1,combobox,1344339971
Control6=IDC_COMBO2,combobox,1344339971
Control7=IDC_CHECK1,button,1342242819
Control8=IDC_EDIT2,edit,1350631552
Control9=IDC_STATIC,static,1342308354
Control10=IDC_STATIC,static,1342308354
Control11=IDC_STATIC,static,1342308352
Control12=IDC_STATIC,static,1342308354
Control13=IDC_STATIC,static,1342308352
Control14=IDC_CHECK2,button,1342242819

[DLG:IDD_PROPPAGE_DBASE]
Type=1
Class=CDbasePropPage
ControlCount=3
Control1=IDC_STATIC,static,1342308866
Control2=IDC_EDIT1,edit,1350631584
Control3=IDC_CHECK1,button,1342242819

[CLS:CDbasePropPage]
Type=0
HeaderFile=DbasePropPage.h
ImplementationFile=DbasePropPage.cpp
BaseClass=COlePropertyPage
Filter=D
LastObject=CDbasePropPage
VirtualFilter=idWC

[DLG:IDD_PROPPAGE_GEN]
Type=1
Class=CGenPropPage
ControlCount=5
Control1=IDC_CHECK2,button,1342242819
Control2=IDC_CHECK3,button,1342242819
Control3=IDC_CHECK4,button,1342242819
Control4=IDC_EDIT1,edit,1350639744
Control5=IDC_STATIC,static,1342308354

[CLS:CGenPropPage]
Type=0
HeaderFile=GenPropPage.h
ImplementationFile=GenPropPage.cpp
BaseClass=COlePropertyPage
Filter=D
LastObject=CGenPropPage
VirtualFilter=idWC

[DLG:IDD_PROPPAGE_DSCDAO (English (U.S.))]
Type=1
Class=?
ControlCount=14
Control1=IDC_EDIT1,edit,1350631552
Control2=IDC_BUTTON1,button,1342242816
Control3=IDC_COMBO4,combobox,1344339971
Control4=IDC_COMBO3,combobox,1344339971
Control5=IDC_COMBO1,combobox,1344339971
Control6=IDC_COMBO2,combobox,1344339971
Control7=IDC_CHECK1,button,1342242819
Control8=IDC_EDIT2,edit,1350631552
Control9=IDC_STATIC,static,1342308354
Control10=IDC_STATIC,static,1342308354
Control11=IDC_STATIC,static,1342308352
Control12=IDC_STATIC,static,1342308354
Control13=IDC_STATIC,static,1342308352
Control14=IDC_CHECK2,button,1342242819

[DLG:IDD_PROPPAGE_DBASE (English (U.S.))]
Type=1
Class=?
ControlCount=3
Control1=IDC_STATIC,static,1342308866
Control2=IDC_EDIT1,edit,1350631584
Control3=IDC_CHECK1,button,1342242819

[DLG:IDD_PROPPAGE_GEN (English (U.S.))]
Type=1
Class=?
ControlCount=5
Control1=IDC_CHECK2,button,1342242819
Control2=IDC_CHECK3,button,1342242819
Control3=IDC_CHECK4,button,1342242819
Control4=IDC_EDIT1,edit,1350639744
Control5=IDC_STATIC,static,1342308354

