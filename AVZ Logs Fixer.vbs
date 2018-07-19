' AVZ Logs Fixer by Dragokas & regist
' ������ ���������� ��� ���������� VirusNet [ SafeZone.cc ]
' ------------------------------------------------------------- ver. 1.14
Option Explicit: ver = "1.14"

' <<< ------ ������� �������� �����, ��� �� ���������� ���� AVZ !!!
LogFolder = "c:\����� � ������ ������"

' <<< ------ true - ���������� ���� � ������������� ���� � ����� ������
CopytoBufer = false

' <<< ------ Silent-�����. ������� �������� true, ���� �� ������ �������� ������ ����: "������", "����� ��� ����������..." � �.�.
Silent = false

' <<< ------ ���� true, �� �� ��������� ����� �����. ������� ����� � �������� �����.
NoBackup = false

' <<< ------ ����� �������� ������� ����� �� ���������� ������ ����� (������� ������ ���� � EXE)
AVZParser = ""
' <<< ------ ����� � ��������� ��� ������ �������. ������ <LOG> ����� ������� �� ������ ���� � ������������� ����.
ParserSwitches = "<LOG>"

'--------------------------------------------------------
' ������� ALF ������������� ��� ����������� ������ � ���� XML �� AVZ,
' ������� �� ��������� ���������� ��� � �������� ��� ����������� � ��������� ������ XML-������.
' ����� � ������� ���� ��������� � ��������� ������ ��� HTM (��������, ����� ������ ������ "��������" ����������� ��� � ���� ������).
'
' ���������:
' ������ ����� ������� �����, ��� �� ������ ������� ���� - ���������� LogFolder (��. ����)
'
' ������� �������������:
'  - ���������� ����� ��� XML-���� (��� ����� � ����) �� ������ ��� ��� �����
'  - ��������� ������ � ������� ����� � ������� ��� XML-������ (��� ZIP-�������)
'  - ����� ����������� ���� -> "���������". ��� ����� ���������� ����� ����� ������ "��������� � ����������� ���� "���������".vbs"
'
' -------------------------------------------------------
' ��������� �����: Dragokas, regist
'
' ��������� ������:
' ����� �������� ZIP-����� ���������� Windows by ALX_2002
' (������������� ���� + ��������� ������� ������������ ������ ZIP by Rob van der Woude (����� ��������������)
' -------------------------------------------------------
'
' ******* ������ ��������� *******
'
' ������ ����� - v.1.0
' Fix # 1. ���������� ������������ ��������� Is64.
'
' ��������� - v.1.1
' ------------------
' �������� ���� �������� ��������� ����� � ������
' ����������� � ����� ������ ���� � ������
' ������ ������ ����� - ��������
'
' ��������� - v.1.2
' ����� SetClipBoard.hta � ����� %temp%
' ������ ���� ��������. ������ ����� � ������ ����������� � ����� ��������� ����.
' "�����" ����� - �������� ����������� � ����� ������� (��-���������, ��������)
'
' ��������� - v.1.3
' ������ SetClipBoard.hta � ������� ������� (���� �������� � �����������)
' ����������� � ����� ���� � ���� ����������, ���� ������� 'CopytoBufer = true' � ����� ���� (��-���������, ���������)
'
' ��������� - v.1.4
' ����� ��� �� ���������, ���� � ��� ������ ���������� (���� ������� Silent-����� ��� ������ ���������)
' ������ ������ �������� ������, �������� �� .zip � .xml
'
' ��������� - v.1.5
' ������ ����������� ����� �������� � ���������� ���� ����� ������� ZIP.
' ���� � ��������� ����� �� ���� ������� ������ virusinfo_syscheck.zip, virusinfo_syscure.zip, avz_sysinfo.xml,
' ������������ ������ ��������� � ���� ����� ����� ZIP, ��� ���� XML.
' ��������� ��� � ���������� ����� �� ����� �����.
' ���� �� ���������� ���� ��� �������� �����, ������������� ��������� ����������.
'
' ��������� - v.1.6
' �������� ���� # 2 (���������������� ���������)
' �������� ���� # 3 (������� NUL)
' �������� ���� # 4 (�������� CheckResult ������� ������).
' ����������� ������� �������� ������� ����� � ����������� �� ���������� ������ ����� (�������� � ����������).
' ������ ������� �� ���������� ������ (������ ������� �� ������� ��� ������/������ �������� ASCII � ����� < 32 � ����).
' ������ ����� ����������� ���� ����� �������� ����� � ������, � �� ������ ����� / XML.
' ����� "������� ����� �������� ��� ��� ������" (��-���������, ���������)
' ��� �������� ���������������� � "�������".
' ������� ������������� � ALF (���������� ������������ ���� ����������).
'
' ��������� - v.1.7
' � v.1.6 �� �������� ����������� � ����� ������ � ����� ��������� ����������. ��� ���������� � ��������.
' ������ "SetClipBoard.hta" ����� � �������.
' ��������� ������ � ����������� ������������ ����.
'
' ��������� - v.1.8
' �������� ���� # 5. ���������������� �������� <local> � ��������� VAL ������ Key.
'
' ��������� - v.1.8h2
' ��������� HTML-����� (������ ���� # 3 (������� NUL))
' �������� ��������� ����������� � ����� (����������������. ������� ComSpec)
'
' ��������� - v.1.9
' �������� ���� # 6. ���������������� ������� � ��������� "CmdLine" ������ <PROCESS>.
'
' ��������� - v.1.10
' ��������� ����������� ������������ �� ������ (��� ��������) �����, ��������� Autologger-�� (��� ������ ������ ��������� ����� "CollectionLog").
'
' ��������� - v.1.11-1.12
' �������� ���� #3 - ������ ������� ��� ��������� ������� (�� ����������� TAB, CR, LF).
'
' ��������� - v.1.13
' �������� ���� # 7 (HTM). ���������������� ���� '<' � ���������� ��������� ������ ������ "������ ���������"
' ��������� ����������� ���������� ������� ������������� ���� � ��������
' ��������� ���, ����� �� �������� HTM � ������ �������� �� ������ ZIP-������
' ������� ��� (����� ������� ������ �������).

Dim ver, AppName, LogFolder, CopytoBufer, Silent, AVZParser, ParserSwitches, NoBackup, syscure, syscheck, syscheckH, syscureH, oFSO, Zip
Dim wasAutoLoggerArc, ALArcFile, ALArcFolder, NoBackup_AL, FolderDestTMP

AppName = "ALF ver." & ver
call main()

Sub main()
	Dim arg, Folder, s, isChanges, isChanges1, isChanges2, isChanges3, isChanges4, ALArcFileNew

	set oFSO = CreateObject("Scripting.FileSystemObject")
	Set Zip = New ZipClass

	'*******************************
	'  �������� ���������� �������
	'*******************************

	if WScript.Arguments.Count > 0 then
		'���� ������ ����������(���) - ������ ������ �� ������, ��������� � ���
		for each arg in WScript.Arguments
			syscheck = "": syscure = "": syscheckH = "": syscureH = ""
			if oFSO.FolderExists(arg) then '���� ������ ����� - ���������� ����� ������� ��� XML � ���
				FindLogs arg, syscheck, syscure
			else
				syscheck = arg '����� �������� � ���������� ������
				if isAutoLoggerArchive(arg) then
					call ALUnpack(arg)	' ���������� CollectionLog
					FindLogs FolderDestTMP, syscheck, syscure
				end if
			end if
			isChanges1 = FixIt(syscheck)
			isChanges2 = FixIt(syscure)
			isChanges = (isChanges1 or isChanges2) or isChanges ' �������� �� ���� ��� ���� �� ����� ������ �������
		next
	else
		'���� �������� �� �����

		'��������� ������� ����� � ������ AVZ
		if not oFSO.FolderExists(LogFolder) then 
			msgbox "����� � ������ ������� �������." & vblf & "����������, ������� ������ � 7-� ������ ����� �������.",vbInformation,AppName
			CreateObject("WScript.Shell").Run "notepad.exe " & """" & WScript.ScriptFullName & """"
			WScript.Quit 1
		end if
		'��������� ������ ������ �����/�����
		Folder = ""
		Folder = OpenFileDialogue(LogFolder)
		if len(Folder) = 0 then WScript.Quit 1 '������ �� �������

		if oFSO.FolderExists(Folder) then
			FindLogs Folder, syscheck, syscure ' ���������� ����� ������� ��� XML � ��������� �����
		else '������ �����
			if oFSO.FileExists(Folder) then	
				syscure = Folder
				if isAutoLoggerArchive(Folder) then
					call ALUnpack(Folder)	' ���������� CollectionLog
					FindLogs FolderDestTMP, syscheck, syscure
				end if
			else
				WScript.Quit 1
			end if
		end if

		' ��������� ����
		isChanges1 = FixIt(syscheck)
		isChanges2 = FixIt(syscure)
		isChanges3 = FixIt(syscheckH)
		isChanges4 = FixIt(syscureH)
		isChanges = (isChanges1 or isChanges2 or isChanges3 or isChanges4)
	end if

	' ���� �������� � ������� �����������
	if wasAutoLoggerArc then
		' ���� ���� ���������
		if isChanges then
			if NoBackup_AL then ' ���� ����� ����������� ����� ������������� �� ������ �����
				ALArcFileNew = ALArcFile
				' ������� ������ �����
				oFSO.DeleteFile ALArcFile, true
			else '����� - � ���������� _fixed
				ALArcFileNew = oFSO.BuildPath(ALArcFolder, oFSO.GetBaseName(ALArcFile) & "_fixed.zip")
			end if
			if Zip.CreateArchive (ALArcFileNew) then
				if oFSO.FolderExists(FolderDestTMP) then
			    	Zip.CopyFolderToArchive FolderDestTMP
					oFSO.DeleteFolder FolderDestTMP, true
				end if
			else
				if not Silent then msgbox "������ ��� ������� ������� ������������ ����� �� ������ �����!",,AppName
			end if
		else ' ��������� ��� - ������� ��������� ����� ���������� ������ �����������
			if oFSO.FolderExists(FolderDestTMP) then oFSO.DeleteFolder FolderDestTMP, true
		end if
	end if

	Set Zip = Nothing

	' ��������� �� ������ ������ �����.
	if isChanges then s = "������." else s = "� ���� ��� ������." & vblf & "���� �� ��� ������."
	if CopytoBufer then s = s & vblf & "���� ���������� � ����� ������."
	if not Silent then msgbox s,,AppName

	' ������ �������� ������� ����� AVZ
	if (Silent or isChanges) then ' ���� ������� ����� Silent ��� ���� ��������� � �������� �����
		if len(AVZParser) <> 0 then ' ���� ������ ������
			if len(syscheck) <> 0 then s = GetFixedName(syscheck) else s = GetFixedName(syscure)
			CreateObject("WScript.Shell").Run chr(34) & AVZParser & chr(34) & " " & chr(34) & replace(ParserSwitches, "<LOG>", s, vbtextcompare) & chr(34)
		end if
	end if

	set oFSO = Nothing
End Sub

Function isAutoLoggerArchive(file)
	isAutoLoggerArchive = (instr(1, oFSO.GetBaseName(file), "CollectionLog", 1) <> 0)
End Function

' ���������� ������ Collectionlog �� ��������� ����� -> ���� ��� + �������� _TMP
' ���������� ��������� ���������� ���������� (wasAutoLoggerArc, NoBackup_AL, ALArcFile, ALArcFolder, FolderDestTMP)
' ����������� ���������� ���������� NoBackup (��� �������������)
Sub ALUnpack(file)
	wasAutoLoggerArc = true ' ����� ������ ����������� �����
	NoBackup_AL = NoBackup ' ���� - ����� �����, ����� �� ������� � ��� �� ������ CollectionLog, ���� ��������� ����� _fixed
	NoBackup = true ' ����� � ������������� ����� �� �������� ����� ������� AVZ � ���������� _fixed
	ALArcFile = file
	ALArcFolder = oFSO.GetParentFolderName(file)
	FolderDestTMP = oFSO.BuildPath(ALArcFolder, oFSO.GetBaseName(file) & "_tmp")
	' ���������� ����� ����� �����������
	Zip.UnpackArchive file, FolderDestTMP
End Sub

' ���� ������ (XML) � ��������� �������� �� ���������� ���������
Sub FindLogs(Folder, syscheck, syscure) 'byref syscheck, byref syscure
	' - ������� ����� � ������ � ��������� ������
	' - ���� � ��������� ����� ������ "virusinfo_syscheck" ��� "virusinfo_syscure";
	' - ����� ���� avz_sysinfo.xml � ������������ virusinfo_syscheck/virusinfo_syscure;
	' - ����� ����� � ����� ����� ������� ����������� ZIP-������;
    ' - ����� ����� � ����� ����� ������� ����������� XML-�����.
    '
		Dim oFile

		Dim SysCheck_v1: SysCheck_v1 = oFSO.BuildPath(Folder, "\virusinfo_syscheck.zip")
		Dim SysCure_v1:  SysCure_v1  = oFSO.BuildPath(Folder, "\virusinfo_syscure.zip")
		Dim SysCheck_v2: SysCheck_v2 = oFSO.BuildPath(Folder, "\virusinfo_syscheck\avz_sysinfo.xml")
		Dim SysCheck_v2h:SysCheck_v2h= oFSO.BuildPath(Folder, "\virusinfo_syscheck\avz_sysinfo.htm")
		Dim SysCure_v2:  SysCure_v2  = oFSO.BuildPath(Folder, "\virusinfo_syscure\avz_sysinfo.xml")
		Dim SysCure_v2h: SysCure_v2h = oFSO.BuildPath(Folder, "\virusinfo_syscure\avz_sysinfo.htm")
		Dim SysCheck_v3: SysCheck_v3 = oFSO.BuildPath(Folder, "\avz_sysinfo.xml")
		Dim SysCheck_v3h:SysCheck_v3h= oFSO.BuildPath(Folder, "\avz_sysinfo.htm")

		syscheck = "": syscure = ""

		' ���� ������ � ����� ��������� �����
		if oFSO.FileExists(SysCheck_v1) then syscheck = SysCheck_v1
		if oFSO.FileExists(SysCure_v1)  then syscure  = SysCure_v1
		if len(syscheck) = 0 and len(syscure) = 0 then
			'�����, ������������, ��� ������ �����������
			if oFSO.FileExists(SysCheck_v2) then syscheck = SysCheck_v2
			if oFSO.FileExists(SysCheck_v2h) then syscheckH=SysCheck_v2h
			if oFSO.FileExists(SysCure_v2)  then syscure  = SysCure_v2
			if oFSO.FileExists(SysCure_v2h) then syscureH = SysCure_v2h
			if oFSO.FileExists(SysCheck_v3) then syscheck = SysCheck_v3
			if oFSO.FileExists(SysCheck_v3h) then syscheckH=SysCheck_v3h
			if len(syscheck) = 0 and len(syscheckH) = 0 then ' ���� ����� ���� �������
				For each oFile in oFSO.GetFolder(Folder).Files
					if lcase(oFSO.GetExtensionName(oFile.Name)) = "zip" then syscheck = oFile.Path: Exit For
				Next
				if len(syscheck) = 0 then 
					For each oFile in oFSO.GetFolder(Folder).Files ' ���� ����� ���� XML
						if lcase(oFSO.GetExtensionName(oFile.Name)) = "xml" then syscheck = oFile.Path: Exit For
					Next
					For each oFile in oFSO.GetFolder(Folder).Files ' ���� ����� ���� HTM
						if lcase(oFSO.GetExtensionName(oFile.Name)) = "htm" then syscheckH =oFile.Path: Exit For
					Next
				end if
			end if
		end if
		if len(syscheck) = 0 and len(syscheckH) = 0 then ' ������ �� ZIP, XML, HTM �� ����� � ��������� ��������
			if not Silent then msgbox "����� ZIP, XML, HTM �� ������� � ��������� ��������!", vbCritical,AppName
			WScript.Quit 1
		end if
		if not Silent then
			if msgbox("����������� ����� �����:" & vbLf & syscure & vbLf & syscheck, vbQuestion + vbYesNo,AppName & ": �������������") = vbNo then WScript.Quit 1
		end if
End Sub

Function GetFixedName(path) ' �������� ���� � ������������� ����
	if NoBackup then GetFixedName = path: Exit Function
	dim ext, RootFolder, BaseName
	on error resume next
    '��������� ���� �� ����������
	ext          = oFSO.GetExtensionName(path)
	RootFolder   = oFSO.GetParentFolderName(path)
	BaseName     = oFSO.GetBaseName(path)
	GetFixedName = oFSO.BuildPath(RootFolder, BaseName & "_fix." & ext)
End Function


' ***************************************
' "���������" ���-���� (����� ���� XML)
' ***************************************

Function FixIt(path)
	dim ext, RootFolder, BaseName, FixItH
	Dim XMLName, XMLFixedName, HTMName, HTMFixedName

	if not oFSO.FileExists(path) then Exit Function '���� ��� �� ���������� - �������

	' �������� ������ ���� � �����, ���� ����� �������������
	if instr(path, ":") = 0 then path = oFSO.GetAbsolutePathName(path)

    '��������� ���� �� ����������
	ext = oFSO.GetExtensionName(path)
	RootFolder = oFSO.GetParentFolderName(path)
	BaseName = oFSO.GetBaseName(path)

	' � ����������� �� ����, ��� ������������ - XML ��� �����
	Select case lcase(ext)
	case "zip"
		' ���� ����� - ��������� � ����� - "����� + fixed"

		'********************
		' ���������� ������
		'********************

		' ������� ����� ��� �����
		' �� ������ �� ����� ��������� "fixed"
		on error resume next
		Dim temp: temp = oFSO.BuildPath(RootFolder, "fixed")
		if oFSO.FolderExists(temp) then '���� ��������� ����� ��� ����������
			if Silent then
				oFSO.DeleteFolder temp, true
			else
				if msgbox("��������! ��������� ����� ��� ����������. ������� ��?" & vblf & vblf & _
					temp, vbQuestion + vbYesNo,AppName & ": �������������") = vbYes then
					oFSO.DeleteFolder temp, true
				else
					msgbox "���� �� ��� ��������.",,AppName
					WScript.Quit 1
				end if
			end if
		end if

		' ���� �� ������� ���� ��� �������� ��������� ����� ����������
		if Err.Number <> 0 then call Elevate("�� ������� ���� ��� �������� ��������� ����� ����������."): WScript.Quit 5
		Err.Clear: on error goto 0
		Zip.UnpackArchive path, temp

		Dim ArcFixed: ArcFixed = GetFixedName(path) '���� � �������������� ������

		'���� fixed-����� ��� ����������
		if oFSO.FileExists(ArcFixed) then 
			if (Silent or NoBackup) then
				oFSO.DeleteFile ArcFixed, true
			else
				if msgbox("��������! ����� ��� ����������. ������� ���?" & vblf & vblf & _
				ArcFixed, vbQuestion + vbYesNo,AppName & ": �������������") = vbYes then
					oFSO.DeleteFile ArcFixed, true
				else
					msgbox "����� �� ��� ������.",,AppName
					WScript.Quit 1
				end if
			end if
		end if

		'**************************************
		' ������� ��������� ���������� ������ 
		'**************************************

		' ����� ����� ����� ��� "������ ���_fix.zip"
		Zip.CreateArchive ArcFixed
		WScript.Sleep(500) ' �������� ������� ����� ��������� ����� ��� ���������� ZIP-�����

		' ***************************
		' ������ ������������� XML 
		' ***************************
		XMLName      = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo.xml")
		XMLFixedName = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo_fix.xml")
		HTMName      = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo.htm")
		HTMFixedName = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo_fix.htm")

		if oFSO.FileExists(XMLName) then FixIt = FixXML (XMLName, XMLFixedName)
		if oFSO.FileExists(HTMName) then FixItH= FixHTM (HTMName, HTMFixedName)

		' ���� ���� ��� "�������" ���� ������� "�����" �����, ���� ���� ��������� ������ ��������� �����
		if (FixIt or FixItH or Silent or NoBackup) then 
			' ������� ������ XML � ������������� ����� � ����
			if oFSO.FileExists(XMLFixedName) then
				oFSO.DeleteFile XMLName
				oFSO.MoveFile XMLFixedName, XMLName
			end if
			if oFSO.FileExists(XMLName) then
				' ***********************
				' ��������� XML � �����
				' ***********************
				Zip.CopyFileToArchive XMLName
				WScript.Sleep(200)
			end if
			' ���������� � HTM
			if oFSO.FileExists(HTMFixedName) then
				oFSO.DeleteFile HTMName
				oFSO.MoveFile HTMFixedName, HTMName
			end if
			if oFSO.FileExists(XMLName) then
				Zip.CopyFileToArchive HTMName
			end if

			Zip.CloseArchive

			' �������� � ����� ���� � ������ ������
			if CopytoBufer then CopyToClipBoard ArcFixed
		else
			if CopytoBufer then CopyToClipBoard path
			' ���� ���� �� �����, ������� ��������� ������
			if oFSO.FileExists(ArcFixed) then on error resume next: oFSO.DeleteFile ArcFixed: on error goto 0
		end if
		WScript.Sleep(200)

		'������� ��������� ����� ����������
		if oFSO.FolderExists(temp) then on error resume next: oFSO.DeleteFolder temp, true: on error goto 0

	case "xml" ' ���� �������� ���� - XML
		XMLFixedName = GetFixedName(path)
		' ������
		FixIt = FixXML(path, XMLFixedName)

		if (FixIt or Silent or NoBackup) then ' ���� ����, ��� "�������" ���� ������� "�����" �����, ���� ���� ��������� ������ ��������� �����
			if CopytoBufer then CopyToClipBoard XMLFixedName
		else
			if CopytoBufer then CopyToClipBoard path
			if oFSO.FileExists(XMLFixedName) then on error resume next: oFSO.DeleteFile XMLFixedName: on error goto 0
		end if

	case "htm" ' ���� �������� ���� - HTM
		HTMFixedName = GetFixedName(path)
		' ������
		FixIt = FixHTM(path, HTMFixedName)

		if (FixIt or Silent or NoBackup) then ' ���� ����, ��� "�������" ���� ������� "�����" �����, ���� ���� ��������� ������ ��������� �����
			if CopytoBufer then CopyToClipBoard HTMFixedName
		else
			if CopytoBufer then CopyToClipBoard path
			if oFSO.FileExists(HTMFixedName) then on error resume next: oFSO.DeleteFile HTMFixedName: on error goto 0
		end if

	case else
		if not Silent then
			msgbox "������ ����� " & ext & " �� �������������� !" & vblf & "��������� ������ .zip � .xml",vbCritical,AppName
		end if
		WScript.Quit 1
	End Select

	if FixItH then FixIt = true

End Function


' *********************
' ������ ����� HTM
' *********************
Function FixHTM(file_or, file_fix)
	Dim pos1, pos2, pos_q2, i, varStr, arStr, sLine, Section, v1

	if not ReadFileToArray(file_or, arStr, file_fix) then exit function

	' fix # 3. NUL
	' ��������: � ���� ������������ ������� Char(0)
	' �������:  ������� ��� ������� NUL
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if instr(sLine, chr(0)) <> 0 then
			FixHTM = true ' ����, ��� ���� ���������
			arStr(i) = replace(sLine, chr(0), vbNullString) '��������� ��������� � �������
		end if
	next

	'fix # 7. < char in 'process' section
	'��������: ���������������� ���� '<' � ���������� ��������� ������ ������ "������ ���������"
	'�������:  �������� ���� '<' �� &lt; � ���� '>' �� &gt; (�� �������)
	Section = false
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if instr(1, sLine, "������ ���������", 1) <> 0 then Section = true '��� ������ � ������ "������ ���������"
		if instr(1, sLine, "Process List", 1) <> 0     then Section = true
		if instr(1, sLine, "�� ��� �������� ��� ����������", 1) <> 0 then Section = false: Exit For ' ����� ������
		if instr(1, sLine, "recognized as trusted", 1) <> 0          then Section = false: Exit For
		if Section then
			'RU
			pos1 = instr(1, sLine, "<br>��������� ������:", 1)
			if pos1 <> 0 then
				pos1 = pos1 +26
			else
			'EN
				pos1 = instr(1, sLine, "<br>Command line:", 1)
				if pos1 <> 0 then pos1 = pos1 +22
			end if

			if pos1 <> 0 then
				if instr(pos1, sLine, "<") <> 0 or instr(pos1, sLine, ">") <> 0 then
					v1 = mid(sLine, pos1) 'CmdLine					
					arStr(i) = left(sLine, pos1 -1) & Replace(Replace(v1, "<", "&lt;"), ">", "&gt;")
					FixHTM = true ' ����, ��� ���� ���������
				end if
			end if
		end if
	next

	call WriteArrayToFile(arStr, file_fix)

End Function


' *********************
' ������ ����� XML
' *********************
Function FixXML(file_or, file_fix)
	Dim pos1, pos2, pos_q2, i, arStr, sLine, j, Section

	if not ReadFileToArray(file_or, arStr, file_fix) then exit function

	' fix # 1. Is64
    ' ��������: ������������ ��������� Is64
	' �������:  �������� ������� ���������
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		pos1 = instr(1, sLine, "Is64=""", vbtextcompare) '������� �� ��������� ����� Is64="
		if (pos1 > 0) then
			pos2 = instr(pos1 + 1, sLine, "Is64=""", vbtextcompare) '���� 2-� ����� �� ����� � ��� �� ������
			if (pos2 > 0) then
				FixXML = true ' ����, ��� ���� ���������
				pos_q2 = instr(pos2 + 6, sLine, """", vbtextcompare)
				arStr(i) = left(sLine, pos2 - 1) & mid(sLine, pos_q2 + 2) '��������� ��������� � �������
			end if
		end if
	next

	' fix # 2. Ampersand
	' ��������: ���������������� ���������� � ����� ������������
	' �������:  ������ & -> &amp;
	dim v1
	sLine = arStr(1) '������ 2-� ������
		pos1 = instr(1, sLine, "ProfileDir=""", vbtextcompare) '������� �� ��������� ����� ProfileDir="
		if (pos1 > 0) then
			pos2 = instr(pos1 + 12, sLine, """", vbtextcompare) '���� 2-� �������
			if (pos2 > 0) then
				v1 = mid(sLine, pos1 + 12, pos2 - (pos1 + 12)) '�������� ���������
				if instr(1, sLine, "&") <> 0 and instr(1, sLine, "&amp;", vbtextcompare) = 0 then
					FixXML = true ' ����, ��� ���� ���������
					arStr(1) = left(sLine, pos1 + 11) & replace(v1, "&", "&amp;") & mid(sLine, pos2) '��������� ��������� � �������
				end if
			end if
		end if

	' fix # 3. NUL + Spec. chars
	' ��������: � ���� ������������ ��������� �������
	' �������:  ������� ��� ��������� �������, ����� TAB, CR, LF
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		for j = 0 to 31
			if j <> 9 and j <> 10 and j <> 13 then
				if instr(sLine, chr(j)) <> 0 then
					FixXML = true ' ����, ��� ���� ���������
					sLine = replace(sLine, chr(j), vbNullString) '��������� ��������� � �������
				end if
			end if
		next
		arStr(i) = sLine
	next

	'fix # 4. CheckResult no-Space
	'��������: ����������� ������ ����� ���������� CheckResult � �������������� ��� ���������
	'�������:  �������� ������
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if instr(1, sLine, """CheckResult=""", vbtextcompare) <> 0 then
			FixXML = true ' ����, ��� ���� ���������
			arStr(i) = replace(sLine, """CheckResult=""", """ CheckResult=""", 1, -1, vbtextcompare) '��������� ��������� � �������
		end if
	next

	'fix # 5. <Value> not quoted
	'��������: ���������������� �������� <local> � ��������� VAL ������ Key
	'�������:  �������� �������� �� &lt;local&gt;
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if left(sLine, 6) = "  <Key" then '��� ������ � ������� Key
			if instr(1, sLine, "<local>", vbtextcompare) <> 0 then '���� �� � ������ <local>
				pos1 = instr(sLine, "VAL=""") '��� �������� VAL
				if pos1 <> 0 then
					pos2 = instr(pos1 + 5, sLine, """") '��� 2-� �������
					v1 = mid(sLine, pos1 + 5, pos2 - (pos1 + 5)) '��������� �������� ���������
					if instr(1, v1, "<local>", vbtextcompare) <> 0 then '���������, ��� <local> ��������� ������ � ��� ��������
						FixXML = true ' ����, ��� ���� ���������
						arStr(i) = left(sLine, pos1 + 4) & replace(v1, "<local>", "&lt;local&gt;", 1, -1, vbtextcompare) & mid(sLine, pos2) '��������� ��������� � �������
					end if
				end if
			end if
		end if
	next

	'fix # 6. " not replaced
	'��������: ���������������� ������� � ��������� "CmdLine" ������ <PROCESS>
	'�������:  �������� ������� �� &quot;
	Section = false
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if left(sLine, 10) = " <PROCESS>" then Section = true '��� ������ � ������ PROCESS
		if left(sLine, 11) = " </PROCESS>" then Section = false: Exit For
		if Section then
			pos1 = instr(sLine, "CmdLine=""")
			if pos1 <> 0 then
				pos1 = pos1 +9
				pos2 = instr(sLine, """ Size=""")
				if pos2 <> 0 then
					v1 = mid(sLine, pos1, pos2 - pos1) 'CmdLine
					if instr(v1, """") <> 0 then
						FixXML = true ' ����, ��� ���� ���������
						arStr(i) = left(sLine, pos1 -1) & replace(v1, """", "&quot;") & mid(sLine, pos2) '��������� ��������� � �������
					end if
				end if
			end if
		end if
	next

	call WriteArrayToFile(arStr, file_fix)

End Function

Function ReadFileToArray(file, arr, file_out)
	Dim varStr
    With CreateObject("ADODB.Stream")
        .Type = 1: .Open: .LoadFromFile file '��������� ���� � �������� ������
		if .Size > 0 then
			varStr = ByteArrayToString(.Read()) '��������������� � ������
			arr = Split(varStr, vbCrLf)			'��������������� � ������ ����� ��� ������� ������
		else
			oFSO.CopyFile file_or, file_out
			if (not Silent) then msgbox "������! ������ ���� = 0 ����.",vbCritical,AppName
			Exit Function
		end if
	end with
	ReadFileToArray = true
End Function

Function WriteArrayToFile(arr, file)
	Dim varStr
	'�������� ������ ����� � ����
    With CreateObject("ADODB.Stream")
		.Type = 1: .Open: .Position = 0
		varStr = join(arr, vbCrLf)
		.Write StringToByteArray(varStr)
		on error resume next
		.SaveToFile file, 2: .Close
		if Err.Number <> 0 then call Elevate("�� ������� ���� ��� �������� ����� " & file): WScript.Quit
		Err.Clear: on error goto 0
	end with
End function


' ����������� � �����
Sub CopyToClipBoard(Data)
	Dim oShell: set oShell = CreateObject("WScript.Shell")
	Dim mshta: mshta = oShell.ExpandEnvironmentStrings("%windir%") & "\system32\mshta.exe" '�������� ���� � ������� HTA
	oShell.Run mshta & " ""vbscript:document.parentwindow.clipboardData.SetData(""text"",replace(""" & replace(Data," ","$#@!~%") & """,""$#@!~%"",chr(32)))&close()""",0,false
End Sub
'Function GetFromClipBoard()
'	GetFromClipBoard = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
'End Function

' ������ ������ �����
Function OpenFileDialogue(StartFolder)
	on error resume next
    Dim oFolder: Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "����� ����� � ������� ��� ������ XML ���� AVZ", 16 + 16384, StartFolder)
    If not (oFolder is Nothing) Then OpenFileDialogue = oFolder.Self.Path
	if Err.Number <> 0 or len(OpenFileDialogue) = 0 then msgbox "�������� ����� ������ ����� !",,AppName: WScript.Quit 1
	set oFolder = Nothing
end Function

' ��������� ����������
Sub Elevate(msg)
	Const DQ = """"
	if msgbox(msg & vblf & "��������� � ����������������� ������������ ?", vbQuestion + vbYesNo,AppName & ": �������������") = vbNo then WScript.Quit 5
	Dim oShellApp: Set oShellApp = CreateObject("Shell.Application")
	' ������������ ����������
	Dim args, i: For i = 1 to WScript.Arguments.Count
		args = args & DQ & WScript.Arguments(i - 1) & DQ & " "
	Next
	if len(args) = 0 then args = DQ & syscure & DQ & " " & DQ & syscheck & DQ
	oShellApp.ShellExecute WScript.FullName, DQ & WScript.ScriptFullName & DQ & " " & args, "", "runas", 1
	set oShellApp = Nothing
End Sub

' ������� �������������� ��������� ������� � ������ � �������
Function ByteArrayToString(varByteArray)
    Dim rs: Set rs = CreateObject("ADODB.Recordset")
    rs.Fields.Append "temp", 201, LenB(varByteArray) 'adLongVarChar
    rs.Open: rs.AddNew: rs("temp").AppendChunk varByteArray: rs.Update
	ByteArrayToString = rs("temp"): rs.Close: Set rs = Nothing
End Function

Function StringToByteArray(sText)
    Dim BS: Set BS = CreateObject("ADODB.Stream")
    BS.Type = 1 'adTypeBinary
    BS.Open
    Dim TS: Set TS = CreateObject("ADODB.Stream")
    With TS
        .Type = 2: .Open: .Charset = "windows-1251": .WriteText sText: .Position = 0: .CopyTo BS: .Close
    End With
    BS.Position = 0: StringToByteArray = BS.Read()
    BS.Close: Set BS = Nothing: Set TS = Nothing
End Function

' ����� �������� ������� ZIP
Class ZipClass
        Private Shell, oShApp, FileSystemObject, ArchiveFolder, ItemsCount, SHCONTF_FILES_AND_FOLDERS
        Private Sub Class_Initialize() '������������� ��������
            Const SHCONTF_FOLDERS               = &H20
            Const SHCONTF_NONFOLDERS            = &H40
            Const SHCONTF_INCLUDEHIDDEN         = &H80
            Const SHCONTF_INCLUDESUPERHIDDEN    = &H10000 ' Windows 7 and Later
            SHCONTF_FILES_AND_FOLDERS = SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS Or SHCONTF_INCLUDEHIDDEN Or SHCONTF_INCLUDESUPERHIDDEN
            Set oShApp = CreateObject("Shell.Application")
            Set Shell = CreateObject("Shell.Application")
            Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
        End Sub
		Private Sub Class_Terminate
			Set Shell = Nothing: Set FileSystemObject = Nothing
		End Sub
		Function UnpackArchive(SourceArchive, DestPath) '���������� ������
			Dim objSource, objTarget
			Set objSource = Shell.NameSpace(SourceArchive).Items()
			on error resume next
			if not FileSystemObject.FolderExists(DestPath) then FileSystemObject.CreateFolder(DestPath)
			' ���� �� ������� ���� ��� �������� ��������� ����� ����������
			if Err.Number <> 0 then call Elevate("�� ������� ���� ��� �������� ��������� ����� ����������."): WScript.Quit 0
			on error goto 0
			Set objTarget = Shell.NameSpace(DestPath)
			objTarget.CopyHere objSource, 4+16
	  	    Do: Wscript.Sleep 200: Loop Until objTarget.Items().Count => objSource.Count
			Set objSource = Nothing
		End Function
        Function CreateArchive(ZipArchivePath) '���������� ZIP-�����
            If lcase(FileSystemObject.GetExtensionName(ZipArchivePath)) <> "zip" Then Exit Function
            Dim ZipFileHeader: ZipFileHeader = "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
			on error resume next
			Dim oTS: set oTS = FileSystemObject.OpenTextFile(ZipArchivePath, 2, True)
			' ���� �� ������� ���� ��� �������� ������
			if Err.Number <> 0 then call Elevate("�� ������� ���� ��� �������� ������."): WScript.Quit 0
			on error goto 0
			oTS.Write ZipFileHeader
			oTS.Close
			Do until FileSystemObject.FileExists(ZipArchivePath): WScript.Sleep(100): Loop '�������� �����, ���� ZIP-����� �� ����� �������
			WScript.Sleep(200)
            Set ArchiveFolder = Shell.NameSpace(ZipArchivePath)
            if Not (ArchiveFolder is Nothing) Then CreateArchive = True
        End Function
        Function CopyFileToArchive(FilePath) '�������� ���� � ZIP-�����
            ArchiveFolder.CopyHere FilePath
            ItemsCount = ItemsCount + 1
			Do: Wscript.Sleep 200: Loop Until ArchiveFolder.Items.Count => ItemsCount '�������� ���� ���-�� �������� � ZIP-����� ������ => ����������
        End Function
        Function CopyFolderToArchive(srcFolderPath) '�������� ���������� ����� � ZIP-�����
            Dim sFilter: set oFolderItems = oShApp.NameSpace(srcFolderPath).Items
			Dim isEmptyFolder, oFolderItem, oFolderItems, ArcItemsNewCount
            oFolderItems.Filter SHCONTF_FILES_AND_FOLDERS, "*" '�������� � ����� ������� �����
            For each oFolderItem in oFolderItems ' ����� ������ �����
                isEmptyFolder = false
                if oFolderItem.IsFolder then if oFolderItem.GetFolder.Items.Count = 0 then isEmptyFolder = true
                if not isEmptyFolder then sFilter = sFilter & ";" & replace(oFolderItem.Name, ";", "?") ' ����� ������ �������� ��� �������
            Next
            oFolderItems.Filter SHCONTF_FILES_AND_FOLDERS, mid(sFilter, 1)
            ArcItemsNewCount = ArchiveFolder.Items.Count + oFolderItems.Count
            ArchiveFolder.CopyHere oFolderItems
            Do: Wscript.Sleep 200: Loop Until ArchiveFolder.Items.Count => ArcItemsNewCount '�������� ���� ���-�� �������� � ZIP-������ ������ >= ���������� � ����
        End Function
        Function CloseArchive '��������� ������ � �������
	  	    Do: Wscript.Sleep 200: Loop Until ArchiveFolder.Items.Count => ItemsCount
	        ItemsCount = 0: Set ArchiveFolder = Nothing
        End Function
End Class
