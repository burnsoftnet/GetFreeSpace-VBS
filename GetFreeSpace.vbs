'==========================================================================
'
' VBScript Source File -- Created with BurnSoft BurnPad
'
' NAME: GetFreeSpace.vbs
'
' AUTHOR:  BurnSoft , www.burnsoft.net, http://opensource.burnsoft.net
' DATE  : 6/14/2010
'
' COMMENT:  This script will list the free Diskspace of the selected machine.
'
'==========================================================================
Dim sMachine, sMsg
Dim NL
Const sTitle = "Free Diskspace Report"
'==========================================================================
Function FileSizeTrans(intID)
	'This is just a simple translator for the size of the file
	Dim s
	Select Case intID
		Case 0
			s = "Bytes"
		Case 1
			s = "KB"
		Case 2
			s = "MB"
		Case 3
			s = "GB"
		Case 4
			s = "TB"
	End Select
	FileSizeTrans = s
End function
'==========================================================================
Function ConvertType(strSize, ByRef MYSizeType)
	Dim s
	'Convert to KiloBytes
	If strSize > 1024 Then 
		s = strSize /1024 : MYSizeType=1
	Else
		s = strSize : MYSizeType=0
	End If
	'Convert to MegaBytes
	If s > 1024 Then s = s /1024 : MYSizeType=2
	'Convert to Gigabytes
	If s > 1024 Then s = s /1024 : MYSizeType=3
	'Convert to TerraBytes
	If s > 1024 Then s = s /1024 : MYSizeType=4
	ConvertType = round(cdbl(s))
End Function
'==========================================================================
Sub AddMsg(sValue)
	If Len(sMsg) = 0 Then
		sMsg = sTitle & NL & NL
		sMsg = sMsg & sValue & NL
	Else
		sMsg = sMsg & sValue & NL
	End if
End Sub
'==========================================================================
Sub GetFreeSpace(Machine)
	Dim sQuery
	Dim ObjEnum,ObjService,ObjInstance
	Dim lSpace, SizeType
	
	sQuery = "SELECT DriveType, FreeSpace, DeviceID from Win32_LogicalDisk"
	Set ObjService = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & Machine & "\root\cimv2")
	Set ObjEnum = ObjService.ExecQuery(sQuery, , 0)
	For Each ObjInstance In ObjEnum
		If Not (ObjInstance Is Nothing) And ObjInstance.drivetype = 3 Then
			lSpace = ConvertType(ObjInstance.FreeSpace,SizeType)
			Call AddMsg(ObjInstance.DeviceId & vbtab & lSpace & " " & FileSizeTrans(SizeType))
		End if
	next
End Sub
'==========================================================================
'Sub Main
	Dim sAns
	Do Until sAns = vbno
		sMsg = ""
		NL = Chr(10)
		sMachine = InputBox("Please Type in Machine Name","Free Disk Space Checker")
		If Len(sMachine) > 0 Then
			Call GetFreeSpace(sMachine)
			If Len(sMsg) > 0 Then MsgBox(sMsg)
		Else
			MsgBox("Please type in a machine to scan!")
		End If
		sAns = MsgBox("Do you wish to look at another machine",vbyesno)
	loop
'End Sub