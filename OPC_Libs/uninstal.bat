@ECHO  ON
set systemdir=  c:\windows\system32
@ECHO  UnRegister OPC Common .......

regsvr32 /s /u %systemdir%\opcproxy.dll
regsvr32 /s /u %systemdir%\opccomn_ps.dll
regsvr32 /s /u %systemdir%\opc_aeps.dll
regsvr32 /s /u %systemdir%\opchda_ps.dll
regsvr32 /s /u %systemdir%\opcdaauto.dll
opcenum /unregserver

@ECHO  Delete OPC Common .......

del %systemdir%\opcproxy.dll
del %systemdir%\opccomn_ps.dll
del %systemdir%\opc_aeps.dll
del %systemdir%\opchda_ps.dll
del %systemdir%\opcdaauto.dll
del %systemdir%\aprxdist.exe
del %systemdir%\opcenum.exe

echo .pause