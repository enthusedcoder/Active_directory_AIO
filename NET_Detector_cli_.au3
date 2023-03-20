#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Icon=Untitled - 11.ico
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=Tool to check installed .NET Framework versions
#AutoIt3Wrapper_Res_Fileversion=1.0.1.0
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region
#EndRegion
$incver = 0
If $cmdline[0] > 0 AND $cmdline[1] = "/v" Then
	$incver = 1
ElseIf $cmdline[0] > 0 AND $cmdline[1] = "/?" Then
	ConsoleWrite(@CRLF & ".NET Detector CLI 2017" & @CRLF & @CRLF & "Use /v to show full version numbers" & @CRLF)
	Exit
EndIf
ConsoleWrite(@CRLF & ".NET Detector CLI" & @CRLF & @CRLF)
If $incver = 0 Then
	ConsoleWrite("Installed:" & @CRLF)
Else
	ConsoleWrite("Installed:" & @TAB & @TAB & @TAB & @TAB & "Version:" & @CRLF)
EndIf
$net1 = RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Active Setup\Installed Components\{78705f0d-e8db-4b2d-8193-982bdda15ecd}", "version")
If $net1 Then
	If $net1 = "1.0.3705.0" Then
		ConsoleWrite(".NET Framework 1.0")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & "1.0.3705.0" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
	If $net1 = "1.0.3705.1" Then
		ConsoleWrite(".NET Framework 1.0 Service Pack 1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & "1.0.3705.1" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
	If $net1 = "1.0.3705.2" Then
		ConsoleWrite(".NET Framework 1.0 Service Pack 2")
		If $incver = 1 Then
			ConsoleWrite(@TAB & "1.0.3705.2" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
	If $net1 = "1.0.3705.3" Then
		ConsoleWrite(".NET Framework 1.0 Service Pack 3")
		If $incver = 1 Then
			ConsoleWrite(@TAB & "1.0.3705.3" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
Else
EndIf
$net11 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v1.1.4322", "install")
$net1164 = RegRead("HKLM\SOFTWARE\Wow6432Node\Microsoft\NET Framework Setup\NDP\v1.1.4322", "install")
If $net11 = "1" OR $net1164 = "1" Then
	$net11sp = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v1.1.4322", "sp")
	$net1164sp = RegRead("HKLM\SOFTWARE\Wow6432Node\Microsoft\NET Framework Setup\NDP\v1.1.4322", "sp")
	If $net11sp = "1" OR $net1164sp = "1" Then
		ConsoleWrite(".NET Framework 1.1 Service Pack 1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & "1.1.4322" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 1.1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & "1.1.4322" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
$net2 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727", "version")
If $net2 Then
	$net2sp = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727", "sp")
	If $net2sp = "1" Then
		ConsoleWrite(".NET Framework 2.0 Service Pack 1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & $net2 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	ElseIf $net2sp = "2" Then
		ConsoleWrite(".NET Framework 2.0 Service Pack 2")
		If $incver = 1 Then
			ConsoleWrite(@TAB & $net2 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 2.0")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net2 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
$net3 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.0", "version")
If $net3 Then
	$net3sp = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.0", "sp")
	If $net3sp = "1" Then
		ConsoleWrite(".NET Framework 3.0 Service Pack 1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & $net3 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	ElseIf $net3sp = "2" Then
		ConsoleWrite(".NET Framework 3.0 Service Pack 2")
		If $incver = 1 Then
			ConsoleWrite(@TAB & $net3 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 3.0")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net3 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
$net35 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5", "version")
If $net35 Then
	$net35sp = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5", "sp")
	If $net35sp = "1" Then
		ConsoleWrite(".NET Framework 3.5 Service Pack 1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & $net35 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 3.5")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net35 & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
$net4cl = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client", "Version")
$net45check = StringLeft($net4cl, 3)
$net45xcheck = StringRight($net4cl, 5)
If $net45check = "4.0" Then
	ConsoleWrite(".NET Framework 4.0 Client Profile")
	If $incver = 1 Then
		ConsoleWrite(@TAB & $net4cl & @CRLF)
	Else
		ConsoleWrite(@CRLF)
	EndIf
EndIf
$net4full = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\full", "Version")
$net45check = StringLeft($net4full, 3)
$net45xcheck = StringRight($net4full, 5)
If $net45check = "4.0" Then
	ConsoleWrite(".NET Framework 4.0 Full Package")
	If $incver = 1 Then
		ConsoleWrite(@TAB & @TAB & $net4full & @CRLF)
	Else
		ConsoleWrite(@CRLF)
	EndIf
EndIf
$net45cl = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\full", "release")
If $net45cl = "378389" Then
	ConsoleWrite(".NET Framework 4.5")
	If $incver = 1 Then
		ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
	Else
		ConsoleWrite(@CRLF)
	EndIf
EndIf
If $net45cl = "378758" OR $net45cl = "378675" Then
	If $net45cl = "378675" Then
		ConsoleWrite(".NET Framework 4.5.1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & " (Windows 8.1)" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 4.5.1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
	$label3 = GUICtrlCreateLabel(".NET 4.5.1", 30, 245, 210, 22)
EndIf
If $net45cl = "379893" Then
	ConsoleWrite(".NET Framework 4.5.2")
	If $incver = 1 Then
		ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
	Else
		ConsoleWrite(@CRLF)
	EndIf
EndIf
If $net45cl = "393297" OR $net45cl = "393295" Then
	If $net45cl = "393295" Then
		ConsoleWrite(".NET Framework 4.6")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & " (Windows 10)" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 4.6")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
If $net45cl = "394271" OR $net45cl = "394254" Then
	If $net45cl = "394254" Then
		ConsoleWrite(".NET Framework 4.6.1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & " (Windows 10 November Update)" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 4.6.1")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
If $net45cl = "394802" OR $net45cl = "394806" Then
	If $net45cl = "394802" Then
		ConsoleWrite(".NET Framework 4.6.2")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & " (Windows 10 Anniversary Update)" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 4.6.2")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
If $net45cl = "460798" OR $net45cl = "460805" Then
	If $net45cl = "460798" Then
		ConsoleWrite(".NET Framework 4.7")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & " (Windows 10 Creators Update)" & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	Else
		ConsoleWrite(".NET Framework 4.7")
		If $incver = 1 Then
			ConsoleWrite(@TAB & @TAB & @TAB & $net4full & @CRLF)
		Else
			ConsoleWrite(@CRLF)
		EndIf
	EndIf
EndIf
