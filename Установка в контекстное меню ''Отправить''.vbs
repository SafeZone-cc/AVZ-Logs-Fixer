Option Explicit

Dim MenuName, oFolder, oFile, s, oTS, oShortCut, fname

Dim oShell: Set oShell = CreateObject("Wscript.Shell")
Dim oFSO: set oFSO = CreateObject("Scripting.FileSystemObject")

'Путь к папке со скриптом
Dim cur: cur = oFSO.GetParentFolderName(WScript.ScriptFullName)

'Путь к папке "Отправить"
Dim SendTo: SendTo = oShell.SpecialFolders("SendTo")

Select case msgbox("Да - Установить" & vblf & "Нет - Удалить",vbYesNoCancel,"ALF: Контекстное меню ""Отправить""")
case vbYes
	MenuName = inputBox("Выберите имя для контекстного пункта","ALF","ALF - пофиксить лог AVZ")
	if len(MenuName) = 0 then msgbox "Установка отменена.",,"ALF": WScript.Quit
	MenuName = NormalizeFileName(MenuName)
	Dim MenuPath: MenuPath = SendTo & "\" & MenuName & ".lnk"
	if oFSO.FileExists(MenuPath) then
		select case (msgbox ("Ярлык " & MenuName & ".vbs" & " уже существует!" & vblf & "Удалить его?",vbInformation + vbYesNo,"ALF"))
		case vbYes
			oFSO.DeleteFile(MenuPath)
		case vbNo
			msgbox "Установка отменена.",vbCritical,"ALF"
			WScript.Quit 1
		End Select
	end if
	set oShortCut = oShell.CreateShortCut(MenuPath)
	oShortCut.TargetPath = cur & "\AVZ Logs Fixer.vbs"
	oShortCut.WorkingDirectory = cur
	oShortCut.Save
	msgbox "Установка завершена." & vblf & "Нажмите по файлу лога AVZ правым кликом мыши -> Отправить" & vblf & "-> " & MenuName,,"ALF"
case vbNo
	'on error resume next
	set oFolder = oFSO.GetFolder(SendTo)
	For Each oFile in oFolder.Files
		if lcase(right(oFile.Path, 4)) = ".lnk" then
			' Подключаемся к ярлыку и проверяем целевой путь (совпадение имени файла)
			set oShortCut = nothing
			set oShortCut = oShell.CreateShortCut(oFile.Path)
			fname = lcase(oFSO.GetFileName(oShortCut.TargetPath))
			if fname = lcase("AVZ Logs Fixer.vbs") or fname = lcase("AVZ Logs Is64 Fixer.vbs") then
				s = oFile.Name
				oFile.Delete(true)
				msgbox s & " был удален",,"ALF"
				WScript.Quit
			end if
		end if
	Next
	msgbox "Не найдено контекстного пункта меню. Удалите его самостоятельно.",,"ALF"
	oShell.Run("explorer.exe " & """" & SendTo & """")
case vbCancel
	WScript.Quit 1
End Select
WScript.Quit

Function NormalizeFileName(FileName) 'Удалить запрещенные для файловой системы знаки
	Dim Char
	For each Char in array("\","/",":","*","?",chr(34),"<",">","|")
		FileName = Replace(FileName, Char, "")
	Next
	NormalizeFileName = FileName
End Function