@ECHO  ON
set systemdir=  c:\windows\system32
@ECHO  copy to OPC Common ........

copy opcproxy.dll %systemdir%
copy opccomn_ps.dll %systemdir%
copy opc_aeps.dll %systemdir%
copy opchda_ps.dll %systemdir%
copy opcdaauto.dll %systemdir%
copy aprxdist.exe %systemdir%
copy opcenum.exe %systemdir%
cd %systemdir%
@ECHO  Register OPC Common ........

REGSVR32 /s opcproxy.dll
REGSVR32 /s opccomn_ps.dll
REGSVR32 /s opc_aeps.dll
REGSVR32 /s opchda_ps.dll
regsvr32 /s opcdaauto.dll
aprxdist.exe
opcenum /regserver
