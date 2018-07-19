Option Explicit

Dim MenuName, oFolder, oFile, s, oTS, oShortCut, fname

Dim oShell: Set oShell = CreateObject("Wscript.Shell")
Dim oFSO: set oFSO = CreateObject("Scripting.FileSystemObject")

'���� � ����� �� ��������
Dim cur: cur = oFSO.GetParentFolderName(WScript.ScriptFullName)

'���� � ����� "���������"
Dim SendTo: SendTo = oShell.SpecialFolders("SendTo")

Select case msgbox("�� - ����������" & vblf & "��� - �������",vbYesNoCancel,"ALF: ����������� ���� ""���������""")
case vbYes
	MenuName = inputBox("�������� ��� ��� ������������ ������","ALF","ALF - ��������� ��� AVZ")
	if len(MenuName) = 0 then msgbox "��������� ��������.",,"ALF": WScript.Quit
	MenuName = NormalizeFileName(MenuName)
	Dim MenuPath: MenuPath = SendTo & "\" & MenuName & ".lnk"
	if oFSO.FileExists(MenuPath) then
		select case (msgbox ("����� " & MenuName & ".vbs" & " ��� ����������!" & vblf & "������� ���?",vbInformation + vbYesNo,"ALF"))
		case vbYes
			oFSO.DeleteFile(MenuPath)
		case vbNo
			msgbox "��������� ��������.",vbCritical,"ALF"
			WScript.Quit 1
		End Select
	end if
	set oShortCut = oShell.CreateShortCut(MenuPath)
	oShortCut.TargetPath = cur & "\AVZ Logs Fixer.vbs"
	oShortCut.WorkingDirectory = cur
	oShortCut.Save
	msgbox "��������� ���������." & vblf & "������� �� ����� ���� AVZ ������ ������ ���� -> ���������" & vblf & "-> " & MenuName,,"ALF"
case vbNo
	'on error resume next
	set oFolder = oFSO.GetFolder(SendTo)
	For Each oFile in oFolder.Files
		if lcase(right(oFile.Path, 4)) = ".lnk" then
			' ������������ � ������ � ��������� ������� ���� (���������� ����� �����)
			set oShortCut = nothing
			set oShortCut = oShell.CreateShortCut(oFile.Path)
			fname = lcase(oFSO.GetFileName(oShortCut.TargetPath))
			if fname = lcase("AVZ Logs Fixer.vbs") or fname = lcase("AVZ Logs Is64 Fixer.vbs") then
				s = oFile.Name
				oFile.Delete(true)
				msgbox s & " ��� ������",,"ALF"
				WScript.Quit
			end if
		end if
	Next
	msgbox "�� ������� ������������ ������ ����. ������� ��� ��������������.",,"ALF"
	oShell.Run("explorer.exe " & """" & SendTo & """")
case vbCancel
	WScript.Quit 1
End Select
WScript.Quit

Function NormalizeFileName(FileName) '������� ����������� ��� �������� ������� �����
	Dim Char
	For each Char in array("\","/",":","*","?",chr(34),"<",">","|")
		FileName = Replace(FileName, Char, "")
	Next
	NormalizeFileName = FileName
End Function