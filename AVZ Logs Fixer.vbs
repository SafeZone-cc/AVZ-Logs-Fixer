' AVZ Logs Fixer by Dragokas & regist
' Создан специально для ассоциации VirusNet [ SafeZone.cc ]
' ------------------------------------------------------------- ver. 1.14
Option Explicit: ver = "1.14"

' <<< ------ Укажите основную папку, где Вы сохраняете логи AVZ !!!
LogFolder = "c:\папка с Вашими логами"

' <<< ------ true - Копировать путь к исправленному логу в буфер обмена
CopytoBufer = false

' <<< ------ Silent-режим. Укажите значение true, если не хотите получать лишних окон: "Готово", "Архив уже существует..." и т.п.
Silent = false

' <<< ------ Если true, то не создавать новый архив. Фиксить сразу в исходном файле.
NoBackup = false

' <<< ------ Вызов внешнего парсера логов по завершению работы фикса (укажите полный путь к EXE)
AVZParser = ""
' <<< ------ Ключи и аргументы при вызове парсера. Макрос <LOG> будет заменен на полный путь к исправленному логу.
ParserSwitches = "<LOG>"

'--------------------------------------------------------
' Утилита ALF предназначена для исправления ошибок в логх XML от AVZ,
' которые не позволяют обработать его в парсерах или просмотреть в программе чтения XML-файлов.
' Позже в утилиту было добавлено и несколько фиксов для HTM (например, когда записи секции "Процессы" расположены все в одну строку).
'
' Настройки:
' Сперва нужно указать папку, где Вы обычно храните логи - переменная LogFolder (см. выше)
'
' Способы использования:
'  - перетянуть архив или XML-файл (или папку с ними) на скрипт или его ярлык
'  - запустить скрипт и указать папку с архивом или XML-файлом (или ZIP-каталог)
'  - через контекстное меню -> "Отправить". Для этого установите пункт через скрипт "Установка в контекстное меню "Отправить".vbs"
'
' -------------------------------------------------------
' Авторские права: Dragokas, regist
'
' Сторонние модули:
' Класс создания ZIP-файла средствами Windows by ALX_2002
' (модифицирован мною + добавлена функция распаковщика архива ZIP by Rob van der Woude (также модифицирована)
' -------------------------------------------------------
'
' ******* СПИСОК ИЗМЕНЕНИЙ *******
'
' Первый релиз - v.1.0
' Fix # 1. Исправляет дублирование параметра Is64.
'
' Изменения - v.1.1
' ------------------
' Добавлен файл настроек начальной папки с логами
' Копирование в буфер обмена лога с фиксом
' Запрет выбора файла - улучшено
'
' Изменения - v.1.2
' Убрал SetClipBoard.hta в папку %temp%
' Удалил файл настроек. Теперь папка с логами указывается в шапке основного кода.
' "Тихий" режим - задается опционально в шапке скрипта (по-умолчанию, отключен)
'
' Изменения - v.1.3
' Вернул SetClipBoard.hta в текущий каталог (были проблемы с антивирусом)
' Копирование в буфер пути к логу происходит, если выбрано 'CopytoBufer = true' в шапке кода (по-умолчанию, отключено)
'
' Изменения - v.1.4
' Новый лог не создается, если в нем нечего исправлять (если включен Silent-режим лог ВСЕГДА создается)
' Запрет выбора форматов файлов, отличных от .zip и .xml
'
' Изменение - v.1.5
' Теперь разрешается также выбирать в диалоговом окне файлы архивов ZIP.
' Если в указанной папке не было найдено файлов virusinfo_syscheck.zip, virusinfo_syscure.zip, avz_sysinfo.xml,
' предлагается первый найденный в этой папке архив ZIP, или файл XML.
' Исправлен баг с обработкой логов из корня диска.
' Если не достаточно прав для создания фикса, запрашивается повышение привилегий.
'
' Изменение - v.1.6
' Добавлен фикс # 2 (Неэкранированный амперсанд)
' Добавлен фикс # 3 (Символы NUL)
' Добавлен фикс # 4 (Параметр CheckResult написан слитно).
' Возможность запуска внешнего парсера логов с аргументами по завершению работы фикса (задается в настройках).
' Движок заменен на побайтовое чтение (защита скрипта от падения при чтении/записи символов ASCII с кодом < 32 в логе).
' Теперь через контекстное меню можно выбирать папку с логами, а не только архив / XML.
' Опция "Фиксить сразу исходный лог без бекапа" (по-умолчанию, отключена)
' Код усиленно прокомментирован и "почищен".
' Утилита переименована в ALF (установщик контекстного меню подправлен).
'
' Изменение - v.1.7
' В v.1.6 не работало копирование в буфер обмена и режим повышения привилегий. Это исправлено и улучшено.
' Скрипт "SetClipBoard.hta" убран в ресурсы.
' Небольшие правки в установщике контекстного меню.
'
' Изменение - v.1.8
' Добавлен фикс # 5. Неэкранированное значение <local> в параметре VAL секции Key.
'
' Изменение - v.1.8h2
' Обработка HTML-логов (только фикс # 3 (Символы NUL))
' Упрощена процедура копирования в буфер (экспериментально. Спасибо ComSpec)
'
' Изменение - v.1.9
' Добавлен фикс # 6. Неэкранированная кавычка в параметре "CmdLine" секции <PROCESS>.
'
' Изменение - v.1.10
' Добавлена возможность перетягивать на скрипт (или выбирать) архив, созданный Autologger-ом (имя архива должно содержать фразу "CollectionLog").
'
' Изменение - v.1.11-1.12
' Дополнен фикс #3 - теперь удаляет все служебные символы (за исключением TAB, CR, LF).
'
' Изменение - v.1.13
' Добавлен фикс # 7 (HTM). Неэкранированный знак '<' в аргументах командной строки секции "Список процессов"
' Добавлена возможность передавать скрипту относительный путь к пациенту
' Исправлен баг, когда не фиксился HTM в случае переноса на скрипт ZIP-архива
' Удалена ЭЦП (более быстрый запуск скрипта).

Dim ver, AppName, LogFolder, CopytoBufer, Silent, AVZParser, ParserSwitches, NoBackup, syscure, syscheck, syscheckH, syscureH, oFSO, Zip
Dim wasAutoLoggerArc, ALArcFile, ALArcFolder, NoBackup_AL, FolderDestTMP

AppName = "ALF ver." & ver
call main()

Sub main()
	Dim arg, Folder, s, isChanges, isChanges1, isChanges2, isChanges3, isChanges4, ALArcFileNew

	set oFSO = CreateObject("Scripting.FileSystemObject")
	Set Zip = New ZipClass

	'*******************************
	'  Проверка аргументов запуска
	'*******************************

	if WScript.Arguments.Count > 0 then
		'Если задано аргументом(ами) - фиксим каждый из файлов, указанных в них
		for each arg in WScript.Arguments
			syscheck = "": syscure = "": syscheckH = "": syscureH = ""
			if oFSO.FolderExists(arg) then 'если задана папка - определяем имена архивов или XML в ней
				FindLogs arg, syscheck, syscure
			else
				syscheck = arg 'иначе работаем с конкретным файлом
				if isAutoLoggerArchive(arg) then
					call ALUnpack(arg)	' распаковка CollectionLog
					FindLogs FolderDestTMP, syscheck, syscure
				end if
			end if
			isChanges1 = FixIt(syscheck)
			isChanges2 = FixIt(syscure)
			isChanges = (isChanges1 or isChanges2) or isChanges ' Применен ли хоть раз фикс за время работы скрипта
		next
	else
		'Если аргумент НЕ задан

		'Проверяем наличие папки с логами AVZ
		if not oFSO.FolderExists(LogFolder) then 
			msgbox "Папка с логами указана неверно." & vblf & "Пожалуйста, задайте другую в 7-й строке этого скрипта.",vbInformation,AppName
			CreateObject("WScript.Shell").Run "notepad.exe " & """" & WScript.ScriptFullName & """"
			WScript.Quit 1
		end if
		'Открываем диалог выбора файла/папки
		Folder = ""
		Folder = OpenFileDialogue(LogFolder)
		if len(Folder) = 0 then WScript.Quit 1 'ничего не выбрано

		if oFSO.FolderExists(Folder) then
			FindLogs Folder, syscheck, syscure ' определяем имена архивов или XML в указанной папке
		else 'выбран архив
			if oFSO.FileExists(Folder) then	
				syscure = Folder
				if isAutoLoggerArchive(Folder) then
					call ALUnpack(Folder)	' распаковка CollectionLog
					FindLogs FolderDestTMP, syscheck, syscure
				end if
			else
				WScript.Quit 1
			end if
		end if

		' Применяем фикс
		isChanges1 = FixIt(syscheck)
		isChanges2 = FixIt(syscure)
		isChanges3 = FixIt(syscheckH)
		isChanges4 = FixIt(syscureH)
		isChanges = (isChanges1 or isChanges2 or isChanges3 or isChanges4)
	end if

	' если работали с архивом Автологгера
	if wasAutoLoggerArc then
		' если были изменения
		if isChanges then
			if NoBackup_AL then ' если архив автологгера нужно пересохранить на старом месте
				ALArcFileNew = ALArcFile
				' удаляем старый архив
				oFSO.DeleteFile ALArcFile, true
			else 'иначе - с постфиксом _fixed
				ALArcFileNew = oFSO.BuildPath(ALArcFolder, oFSO.GetBaseName(ALArcFile) & "_fixed.zip")
			end if
			if Zip.CreateArchive (ALArcFileNew) then
				if oFSO.FolderExists(FolderDestTMP) then
			    	Zip.CopyFolderToArchive FolderDestTMP
					oFSO.DeleteFolder FolderDestTMP, true
				end if
			else
				if not Silent then msgbox "Ошибка при попытке создать исправленный архив на старом месте!",,AppName
			end if
		else ' изменений нет - удаляем временную папку распаковки архива Автологгера
			if oFSO.FolderExists(FolderDestTMP) then oFSO.DeleteFolder FolderDestTMP, true
		end if
	end if

	Set Zip = Nothing

	' Сообщение об отчете работы фикса.
	if isChanges then s = "Готово." else s = "В логе нет ошибок." & vblf & "Фикс не был создан."
	if CopytoBufer then s = s & vblf & "Путь скопирован в буфер обмена."
	if not Silent then msgbox s,,AppName

	' Запуск внешнего парсера логов AVZ
	if (Silent or isChanges) then ' если включен режим Silent или были изменения в процессе фикса
		if len(AVZParser) <> 0 then ' если парсер указан
			if len(syscheck) <> 0 then s = GetFixedName(syscheck) else s = GetFixedName(syscure)
			CreateObject("WScript.Shell").Run chr(34) & AVZParser & chr(34) & " " & chr(34) & replace(ParserSwitches, "<LOG>", s, vbtextcompare) & chr(34)
		end if
	end if

	set oFSO = Nothing
End Sub

Function isAutoLoggerArchive(file)
	isAutoLoggerArchive = (instr(1, oFSO.GetBaseName(file), "CollectionLog", 1) <> 0)
End Function

' Распаковка архива Collectionlog во временную папку -> тоже имя + постфикс _TMP
' Заполнение некоторых глобальных переменных (wasAutoLoggerArc, NoBackup_AL, ALArcFile, ALArcFolder, FolderDestTMP)
' Модификация глобальной переменной NoBackup (при необходимости)
Sub ALUnpack(file)
	wasAutoLoggerArc = true ' чтобы заново перепаковал архив
	NoBackup_AL = NoBackup ' флаг - чтобы знать, нужно ли фиксить в том же архиве CollectionLog, либо создавать копию _fixed
	NoBackup = true ' чтобы в распакованной папке не создавал копию архивов AVZ с постфиксом _fixed
	ALArcFile = file
	ALArcFolder = oFSO.GetParentFolderName(file)
	FolderDestTMP = oFSO.BuildPath(ALArcFolder, oFSO.GetBaseName(file) & "_tmp")
	' Распаковую архив логов Автологгера
	Zip.UnpackArchive file, FolderDestTMP
End Sub

' Ищет архивы (XML) в указанном каталоге по нескольким критериям
Sub FindLogs(Folder, syscheck, syscure) 'byref syscheck, byref syscure
	' - смотрим папку с логами и открываем диалог
	' - ищем в указанной папке архивы "virusinfo_syscheck" или "virusinfo_syscure";
	' - иначе ищем avz_sysinfo.xml в подкаталогах virusinfo_syscheck/virusinfo_syscure;
	' - иначе ведем в папке поиск первого попавшегося ZIP-архива;
    ' - иначе ведем в папке поиск первого попавшегося XML-файла.
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

		' ищем архивы в корне указанной папки
		if oFSO.FileExists(SysCheck_v1) then syscheck = SysCheck_v1
		if oFSO.FileExists(SysCure_v1)  then syscure  = SysCure_v1
		if len(syscheck) = 0 and len(syscure) = 0 then
			'иначе, предполагаем, что архивы распакованы
			if oFSO.FileExists(SysCheck_v2) then syscheck = SysCheck_v2
			if oFSO.FileExists(SysCheck_v2h) then syscheckH=SysCheck_v2h
			if oFSO.FileExists(SysCure_v2)  then syscure  = SysCure_v2
			if oFSO.FileExists(SysCure_v2h) then syscureH = SysCure_v2h
			if oFSO.FileExists(SysCheck_v3) then syscheck = SysCheck_v3
			if oFSO.FileExists(SysCheck_v3h) then syscheckH=SysCheck_v3h
			if len(syscheck) = 0 and len(syscheckH) = 0 then ' Веду поиск всех архивов
				For each oFile in oFSO.GetFolder(Folder).Files
					if lcase(oFSO.GetExtensionName(oFile.Name)) = "zip" then syscheck = oFile.Path: Exit For
				Next
				if len(syscheck) = 0 then 
					For each oFile in oFSO.GetFolder(Folder).Files ' Веду поиск всех XML
						if lcase(oFSO.GetExtensionName(oFile.Name)) = "xml" then syscheck = oFile.Path: Exit For
					Next
					For each oFile in oFSO.GetFolder(Folder).Files ' Веду поиск всех HTM
						if lcase(oFSO.GetExtensionName(oFile.Name)) = "htm" then syscheckH =oFile.Path: Exit For
					Next
				end if
			end if
		end if
		if len(syscheck) = 0 and len(syscheckH) = 0 then ' Ничего из ZIP, XML, HTM не нашли в указанном каталоге
			if not Silent then msgbox "Файлы ZIP, XML, HTM не найдены в указанном каталоге!", vbCritical,AppName
			WScript.Quit 1
		end if
		if not Silent then
			if msgbox("Подтвердите выбор логов:" & vbLf & syscure & vbLf & syscheck, vbQuestion + vbYesNo,AppName & ": Подтверждение") = vbNo then WScript.Quit 1
		end if
End Sub

Function GetFixedName(path) ' Получить путь к исправленному логу
	if NoBackup then GetFixedName = path: Exit Function
	dim ext, RootFolder, BaseName
	on error resume next
    'Разбираем путь на компоненты
	ext          = oFSO.GetExtensionName(path)
	RootFolder   = oFSO.GetParentFolderName(path)
	BaseName     = oFSO.GetBaseName(path)
	GetFixedName = oFSO.BuildPath(RootFolder, BaseName & "_fix." & ext)
End Function


' ***************************************
' "Пофиксить" лог-файл (архив либо XML)
' ***************************************

Function FixIt(path)
	dim ext, RootFolder, BaseName, FixItH
	Dim XMLName, XMLFixedName, HTMName, HTMFixedName

	if not oFSO.FileExists(path) then Exit Function 'Если лог не существует - выходим

	' получаем полный путь к файлу, если задан относительный
	if instr(path, ":") = 0 then path = oFSO.GetAbsolutePathName(path)

    'Разбираем путь на компоненты
	ext = oFSO.GetExtensionName(path)
	RootFolder = oFSO.GetParentFolderName(path)
	BaseName = oFSO.GetBaseName(path)

	' В зависимости от того, что обрабатываем - XML или архив
	Select case lcase(ext)
	case "zip"
		' Если архив - распакуем в папку - "Архив + fixed"

		'********************
		' Распаковка архива
		'********************

		' Готовим место под архив
		' Не занята ли папка временная "fixed"
		on error resume next
		Dim temp: temp = oFSO.BuildPath(RootFolder, "fixed")
		if oFSO.FolderExists(temp) then 'если временная папка уже существует
			if Silent then
				oFSO.DeleteFolder temp, true
			else
				if msgbox("Внимание! Временная папка уже существует. Удалить ее?" & vblf & vblf & _
					temp, vbQuestion + vbYesNo,AppName & ": Подтверждение") = vbYes then
					oFSO.DeleteFolder temp, true
				else
					msgbox "Фикс не был применен.",,AppName
					WScript.Quit 1
				end if
			end if
		end if

		' Если не хватает прав для удаления временной папки распаковки
		if Err.Number <> 0 then call Elevate("Не хватает прав для удаления временной папки распаковки."): WScript.Quit 5
		Err.Clear: on error goto 0
		Zip.UnpackArchive path, temp

		Dim ArcFixed: ArcFixed = GetFixedName(path) 'путь к исправаленному архиву

		'если fixed-архив уже существует
		if oFSO.FileExists(ArcFixed) then 
			if (Silent or NoBackup) then
				oFSO.DeleteFile ArcFixed, true
			else
				if msgbox("Внимание! Архив уже существует. Удалить его?" & vblf & vblf & _
				ArcFixed, vbQuestion + vbYesNo,AppName & ": Подтверждение") = vbYes then
					oFSO.DeleteFile ArcFixed, true
				else
					msgbox "Архив не был создан.",,AppName
					WScript.Quit 1
				end if
			end if
		end if

		'**************************************
		' Создаем заготовку финального архива 
		'**************************************

		' Архив будет иметь имя "Старое имя_fix.zip"
		Zip.CreateArchive ArcFixed
		WScript.Sleep(500) ' файловой системе нужно некоторое время для подготовки ZIP-папки

		' ***************************
		' Фиксим распакованный XML 
		' ***************************
		XMLName      = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo.xml")
		XMLFixedName = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo_fix.xml")
		HTMName      = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo.htm")
		HTMFixedName = oFSO.BuildPath(RootFolder, "fixed\avz_sysinfo_fix.htm")

		if oFSO.FileExists(XMLName) then FixIt = FixXML (XMLName, XMLFixedName)
		if oFSO.FileExists(HTMName) then FixItH= FixHTM (HTMName, HTMFixedName)

		' Если было что "фиксить" либо включен "тихий" режим, либо фикс создается поверх исходного файла
		if (FixIt or FixItH or Silent or NoBackup) then 
			' Удаляем старый XML и переименовуем новый в него
			if oFSO.FileExists(XMLFixedName) then
				oFSO.DeleteFile XMLName
				oFSO.MoveFile XMLFixedName, XMLName
			end if
			if oFSO.FileExists(XMLName) then
				' ***********************
				' Добавляем XML в архив
				' ***********************
				Zip.CopyFileToArchive XMLName
				WScript.Sleep(200)
			end if
			' Аналогично с HTM
			if oFSO.FileExists(HTMFixedName) then
				oFSO.DeleteFile HTMName
				oFSO.MoveFile HTMFixedName, HTMName
			end if
			if oFSO.FileExists(XMLName) then
				Zip.CopyFileToArchive HTMName
			end if

			Zip.CloseArchive

			' Копируем в буфер путь к новому архиву
			if CopytoBufer then CopyToClipBoard ArcFixed
		else
			if CopytoBufer then CopyToClipBoard path
			' Если фикс не нужен, удаляем заготовку архива
			if oFSO.FileExists(ArcFixed) then on error resume next: oFSO.DeleteFile ArcFixed: on error goto 0
		end if
		WScript.Sleep(200)

		'Удаляем временную папку распаковки
		if oFSO.FolderExists(temp) then on error resume next: oFSO.DeleteFolder temp, true: on error goto 0

	case "xml" ' если исходный файл - XML
		XMLFixedName = GetFixedName(path)
		' фиксим
		FixIt = FixXML(path, XMLFixedName)

		if (FixIt or Silent or NoBackup) then ' Если было, что "фиксить" либо включен "тихий" режим, либо фикс создается поверх исходного файла
			if CopytoBufer then CopyToClipBoard XMLFixedName
		else
			if CopytoBufer then CopyToClipBoard path
			if oFSO.FileExists(XMLFixedName) then on error resume next: oFSO.DeleteFile XMLFixedName: on error goto 0
		end if

	case "htm" ' если исходный файл - HTM
		HTMFixedName = GetFixedName(path)
		' фиксим
		FixIt = FixHTM(path, HTMFixedName)

		if (FixIt or Silent or NoBackup) then ' Если было, что "фиксить" либо включен "тихий" режим, либо фикс создается поверх исходного файла
			if CopytoBufer then CopyToClipBoard HTMFixedName
		else
			if CopytoBufer then CopyToClipBoard path
			if oFSO.FileExists(HTMFixedName) then on error resume next: oFSO.DeleteFile HTMFixedName: on error goto 0
		end if

	case else
		if not Silent then
			msgbox "Формат файла " & ext & " не поддерживается !" & vblf & "Допустимы только .zip и .xml",vbCritical,AppName
		end if
		WScript.Quit 1
	End Select

	if FixItH then FixIt = true

End Function


' *********************
' Движок фикса HTM
' *********************
Function FixHTM(file_or, file_fix)
	Dim pos1, pos2, pos_q2, i, varStr, arStr, sLine, Section, v1

	if not ReadFileToArray(file_or, arStr, file_fix) then exit function

	' fix # 3. NUL
	' Проблема: В логе присутствуют символы Char(0)
	' Решение:  удалить все символы NUL
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if instr(sLine, chr(0)) <> 0 then
			FixHTM = true ' флаг, что были изменения
			arStr(i) = replace(sLine, chr(0), vbNullString) 'сохраняем изменения в массиве
		end if
	next

	'fix # 7. < char in 'process' section
	'Проблема: Неэкранированный знак '<' в аргументах командной строки секции "Список процессов"
	'Решение:  Заменить знак '<' на &lt; и знак '>' на &gt; (на будущее)
	Section = false
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if instr(1, sLine, "Список процессов", 1) <> 0 then Section = true 'Ищу только в секции "Список процессов"
		if instr(1, sLine, "Process List", 1) <> 0     then Section = true
		if instr(1, sLine, "из них опознаны как безопасные", 1) <> 0 then Section = false: Exit For ' конец секции
		if instr(1, sLine, "recognized as trusted", 1) <> 0          then Section = false: Exit For
		if Section then
			'RU
			pos1 = instr(1, sLine, "<br>Командная строка:", 1)
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
					FixHTM = true ' флаг, что были изменения
				end if
			end if
		end if
	next

	call WriteArrayToFile(arStr, file_fix)

End Function


' *********************
' Движок фикса XML
' *********************
Function FixXML(file_or, file_fix)
	Dim pos1, pos2, pos_q2, i, arStr, sLine, j, Section

	if not ReadFileToArray(file_or, arStr, file_fix) then exit function

	' fix # 1. Is64
    ' Проблема: Дублирование параметра Is64
	' Решение:  удаление лишнего параметра
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		pos1 = instr(1, sLine, "Is64=""", vbtextcompare) 'парсинг по ключевому слову Is64="
		if (pos1 > 0) then
			pos2 = instr(pos1 + 1, sLine, "Is64=""", vbtextcompare) 'ищем 2-е такое же слово в той же строке
			if (pos2 > 0) then
				FixXML = true ' флаг, что были изменения
				pos_q2 = instr(pos2 + 6, sLine, """", vbtextcompare)
				arStr(i) = left(sLine, pos2 - 1) & mid(sLine, pos_q2 + 2) 'сохраняем изменения в массиве
			end if
		end if
	next

	' fix # 2. Ampersand
	' Проблема: Неэкранированные амперсанды в имени пользователя
	' Решение:  замена & -> &amp;
	dim v1
	sLine = arStr(1) 'Читаем 2-ю строку
		pos1 = instr(1, sLine, "ProfileDir=""", vbtextcompare) 'парсинг по ключевому слову ProfileDir="
		if (pos1 > 0) then
			pos2 = instr(pos1 + 12, sLine, """", vbtextcompare) 'ищем 2-ю кавычку
			if (pos2 > 0) then
				v1 = mid(sLine, pos1 + 12, pos2 - (pos1 + 12)) 'Значение параметра
				if instr(1, sLine, "&") <> 0 and instr(1, sLine, "&amp;", vbtextcompare) = 0 then
					FixXML = true ' флаг, что были изменения
					arStr(1) = left(sLine, pos1 + 11) & replace(v1, "&", "&amp;") & mid(sLine, pos2) 'сохраняем изменения в массиве
				end if
			end if
		end if

	' fix # 3. NUL + Spec. chars
	' Проблема: В логе присутствуют служебные символы
	' Решение:  удалить все служебные символы, кроме TAB, CR, LF
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		for j = 0 to 31
			if j <> 9 and j <> 10 and j <> 13 then
				if instr(sLine, chr(j)) <> 0 then
					FixXML = true ' флаг, что были изменения
					sLine = replace(sLine, chr(j), vbNullString) 'сохраняем изменения в массиве
				end if
			end if
		next
		arStr(i) = sLine
	next

	'fix # 4. CheckResult no-Space
	'Проблема: Отсутствует пробел между параметром CheckResult и предшествующем ему значением
	'Решение:  добавить пробел
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if instr(1, sLine, """CheckResult=""", vbtextcompare) <> 0 then
			FixXML = true ' флаг, что были изменения
			arStr(i) = replace(sLine, """CheckResult=""", """ CheckResult=""", 1, -1, vbtextcompare) 'сохраняем изменения в массиве
		end if
	next

	'fix # 5. <Value> not quoted
	'Проблема: Неэкранированное значение <local> в параметре VAL секции Key
	'Решение:  Заменить значение на &lt;local&gt;
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if left(sLine, 6) = "  <Key" then 'Ищу только в секциях Key
			if instr(1, sLine, "<local>", vbtextcompare) <> 0 then 'Есть ли в строке <local>
				pos1 = instr(sLine, "VAL=""") 'Ищу параметр VAL
				if pos1 <> 0 then
					pos2 = instr(pos1 + 5, sLine, """") 'Ищу 2-ю кавычку
					v1 = mid(sLine, pos1 + 5, pos2 - (pos1 + 5)) 'Вытягиваю значение параметра
					if instr(1, v1, "<local>", vbtextcompare) <> 0 then 'Убеждаюсь, что <local> находится именно в его значении
						FixXML = true ' флаг, что были изменения
						arStr(i) = left(sLine, pos1 + 4) & replace(v1, "<local>", "&lt;local&gt;", 1, -1, vbtextcompare) & mid(sLine, pos2) 'сохраняем изменения в массиве
					end if
				end if
			end if
		end if
	next

	'fix # 6. " not replaced
	'Проблема: Неэкранированная кавычка в параметре "CmdLine" секции <PROCESS>
	'Решение:  Заменить кавычку на &quot;
	Section = false
	for i = 0 to Ubound(arStr)
		sLine = arStr(i)
		if left(sLine, 10) = " <PROCESS>" then Section = true 'Ищу только в секции PROCESS
		if left(sLine, 11) = " </PROCESS>" then Section = false: Exit For
		if Section then
			pos1 = instr(sLine, "CmdLine=""")
			if pos1 <> 0 then
				pos1 = pos1 +9
				pos2 = instr(sLine, """ Size=""")
				if pos2 <> 0 then
					v1 = mid(sLine, pos1, pos2 - pos1) 'CmdLine
					if instr(v1, """") <> 0 then
						FixXML = true ' флаг, что были изменения
						arStr(i) = left(sLine, pos1 -1) & replace(v1, """", "&quot;") & mid(sLine, pos2) 'сохраняем изменения в массиве
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
        .Type = 1: .Open: .LoadFromFile file 'загружаем файл в байтовый массив
		if .Size > 0 then
			varStr = ByteArrayToString(.Read()) 'преобразовываем в строку
			arr = Split(varStr, vbCrLf)			'преобразовываем в массив строк для удобной работы
		else
			oFSO.CopyFile file_or, file_out
			if (not Silent) then msgbox "Ошибка! Размер лога = 0 байт.",vbCritical,AppName
			Exit Function
		end if
	end with
	ReadFileToArray = true
End Function

Function WriteArrayToFile(arr, file)
	Dim varStr
	'Сохраняю массив строк в файл
    With CreateObject("ADODB.Stream")
		.Type = 1: .Open: .Position = 0
		varStr = join(arr, vbCrLf)
		.Write StringToByteArray(varStr)
		on error resume next
		.SaveToFile file, 2: .Close
		if Err.Number <> 0 then call Elevate("Не хватает прав для создания фикса " & file): WScript.Quit
		Err.Clear: on error goto 0
	end with
End function


' Копирование в буфер
Sub CopyToClipBoard(Data)
	Dim oShell: set oShell = CreateObject("WScript.Shell")
	Dim mshta: mshta = oShell.ExpandEnvironmentStrings("%windir%") & "\system32\mshta.exe" 'Получаем путь к серверу HTA
	oShell.Run mshta & " ""vbscript:document.parentwindow.clipboardData.SetData(""text"",replace(""" & replace(Data," ","$#@!~%") & """,""$#@!~%"",chr(32)))&close()""",0,false
End Sub
'Function GetFromClipBoard()
'	GetFromClipBoard = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
'End Function

' Диалог выбора файла
Function OpenFileDialogue(StartFolder)
	on error resume next
    Dim oFolder: Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "Выбор папки с архивом или файлом XML лога AVZ", 16 + 16384, StartFolder)
    If not (oFolder is Nothing) Then OpenFileDialogue = oFolder.Self.Path
	if Err.Number <> 0 or len(OpenFileDialogue) = 0 then msgbox "Выбирать можно только папки !",,AppName: WScript.Quit 1
	set oFolder = Nothing
end Function

' Повышение привилегий
Sub Elevate(msg)
	Const DQ = """"
	if msgbox(msg & vblf & "Запустить с Административными привилегиями ?", vbQuestion + vbYesNo,AppName & ": Подтверждение") = vbNo then WScript.Quit 5
	Dim oShellApp: Set oShellApp = CreateObject("Shell.Application")
	' Конкатенация аргументов
	Dim args, i: For i = 1 to WScript.Arguments.Count
		args = args & DQ & WScript.Arguments(i - 1) & DQ & " "
	Next
	if len(args) = 0 then args = DQ & syscure & DQ & " " & DQ & syscheck & DQ
	oShellApp.ShellExecute WScript.FullName, DQ & WScript.ScriptFullName & DQ & " " & args, "", "runas", 1
	set oShellApp = Nothing
End Sub

' Функции преобразования байтового массива в строку и обратно
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

' Класс создания архивов ZIP
Class ZipClass
        Private Shell, oShApp, FileSystemObject, ArchiveFolder, ItemsCount, SHCONTF_FILES_AND_FOLDERS
        Private Sub Class_Initialize() 'Инициализация объектов
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
		Function UnpackArchive(SourceArchive, DestPath) 'Распаковка архива
			Dim objSource, objTarget
			Set objSource = Shell.NameSpace(SourceArchive).Items()
			on error resume next
			if not FileSystemObject.FolderExists(DestPath) then FileSystemObject.CreateFolder(DestPath)
			' Если не хватает прав для создания временной папки распаковки
			if Err.Number <> 0 then call Elevate("Не хватает прав для создания временной папки распаковки."): WScript.Quit 0
			on error goto 0
			Set objTarget = Shell.NameSpace(DestPath)
			objTarget.CopyHere objSource, 4+16
	  	    Do: Wscript.Sleep 200: Loop Until objTarget.Items().Count => objSource.Count
			Set objSource = Nothing
		End Function
        Function CreateArchive(ZipArchivePath) 'Подготовка ZIP-папки
            If lcase(FileSystemObject.GetExtensionName(ZipArchivePath)) <> "zip" Then Exit Function
            Dim ZipFileHeader: ZipFileHeader = "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
			on error resume next
			Dim oTS: set oTS = FileSystemObject.OpenTextFile(ZipArchivePath, 2, True)
			' Если не хватает прав для создания архива
			if Err.Number <> 0 then call Elevate("Не хватает прав для создания архива."): WScript.Quit 0
			on error goto 0
			oTS.Write ZipFileHeader
			oTS.Close
			Do until FileSystemObject.FileExists(ZipArchivePath): WScript.Sleep(100): Loop 'выжидаем время, пока ZIP-папка не будет создана
			WScript.Sleep(200)
            Set ArchiveFolder = Shell.NameSpace(ZipArchivePath)
            if Not (ArchiveFolder is Nothing) Then CreateArchive = True
        End Function
        Function CopyFileToArchive(FilePath) 'Копируем файл в ZIP-папку
            ArchiveFolder.CopyHere FilePath
            ItemsCount = ItemsCount + 1
			Do: Wscript.Sleep 200: Loop Until ArchiveFolder.Items.Count => ItemsCount 'Выжидаем пока кол-во объектов в ZIP-папке станет => копируемым
        End Function
        Function CopyFolderToArchive(srcFolderPath) 'Копируем содержимое папки в ZIP-архив
            Dim sFilter: set oFolderItems = oShApp.NameSpace(srcFolderPath).Items
			Dim isEmptyFolder, oFolderItem, oFolderItems, ArcItemsNewCount
            oFolderItems.Filter SHCONTF_FILES_AND_FOLDERS, "*" 'включаем в архив скрытые файлы
            For each oFolderItem in oFolderItems ' поиск пустых папок
                isEmptyFolder = false
                if oFolderItem.IsFolder then if oFolderItem.GetFolder.Items.Count = 0 then isEmptyFolder = true
                if not isEmptyFolder then sFilter = sFilter & ";" & replace(oFolderItem.Name, ";", "?") ' белый список объектов для фильтра
            Next
            oFolderItems.Filter SHCONTF_FILES_AND_FOLDERS, mid(sFilter, 1)
            ArcItemsNewCount = ArchiveFolder.Items.Count + oFolderItems.Count
            ArchiveFolder.CopyHere oFolderItems
            Do: Wscript.Sleep 200: Loop Until ArchiveFolder.Items.Count => ArcItemsNewCount 'Выжидаем пока кол-во объектов в ZIP-архиве станет >= копируемым в него
        End Function
        Function CloseArchive 'завершаем работу с архивом
	  	    Do: Wscript.Sleep 200: Loop Until ArchiveFolder.Items.Count => ItemsCount
	        ItemsCount = 0: Set ArchiveFolder = Nothing
        End Function
End Class
