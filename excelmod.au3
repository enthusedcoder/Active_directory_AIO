; *** Start added by AutoIt3Wrapper ***
#include <ExcelConstants.au3>
; *** End added by AutoIt3Wrapper ***
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Add_Constants=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.15.0 (Beta)
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Excel.au3>
#include <Array.au3>

If _GetFilenameExt($CmdLine[1]) = "xlsx" Then
	$oExcel = _Excel_Open ()
	$oBook = _Excel_BookOpen ( $oExcel, $CmdLine[1] )
	$find = _Excel_RangeFind ( $oBook, "Employee Id" )
	$end = Int ( StringRight ( $find[0][2], 1 ) ) - 1
	_Excel_RangeDelete ( $oBook.ActiveSheet, "1:" & $end )
	_Excel_BookSaveAs ( $oBook, @AppDataDir & "\store.csv", $xlCSVWindows, True )
	_Excel_BookClose ( $oBook )
	_Excel_Close ( $oExcel )
Else
	FileCopy ( $CmdLine[1], @AppDataDir & "\store.csv", 1 )
EndIf




Func _GetFilenameDrive($sFilePath)
    Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
    Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
    If IsObj($oColFiles) Then
        For $oObjectFile In $oColFiles
            Return StringUpper($oObjectFile.Drive)
        Next
    EndIf
    Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenameDrive

Func _GetFilenamePath($sFilePath)
    Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
    Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
    If IsObj($oColFiles) Then
        For $oObjectFile In $oColFiles
            Return $oObjectFile.Path
        Next
    EndIf
    Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenamePath
Func _GetFilename($sFilePath)
    Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
    Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
    If IsObj($oColFiles) Then
        For $oObjectFile In $oColFiles
            Return $oObjectFile.FileName
        Next
    EndIf
    Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilename

Func _GetFilenameExt($sFilePath)
    Local $oWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & "." & "\root\cimv2")
    Local $oColFiles = $oWMIService.ExecQuery("Select * From CIM_Datafile Where Name = '" & StringReplace($sFilePath, "\", "\\") & "'")
    If IsObj($oColFiles) Then
        For $oObjectFile In $oColFiles
            Return $oObjectFile.Extension
        Next
    EndIf
    Return SetError(1, 1, 0)
EndFunc   ;==>_GetFilenameExt