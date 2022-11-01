<%'/*
'**********************************************
'      /\      | (_)
'     /  \   __| |_  __ _ _ __  ___
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/
'**********************************************
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************
'*/
Class Zip_File_Management_Plugin
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Variables
	'---------------------------------------------------------------
	'*/
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Variables
	'---------------------------------------------------------------
	'*/
	Dim BlankZip, NoInterfaceYesToAll
	Dim fso, curArquieve, created, saved
	Dim files, m_path, zipApp, zipFile	
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class
	'---------------------------------------------------------------
	'*/
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		'/*
		'---------------------------------------------------------------
		' Check Register
		'---------------------------------------------------------------
		'*/
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If
		'/*
		'---------------------------------------------------------------
		' Plugin Settings
		'---------------------------------------------------------------
		'*/
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE&"_")
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "Zip_File_Management_Plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "345")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)
		'/*
		'---------------------------------------------------------------
		' Register Settings
		'---------------------------------------------------------------
		'*/
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Register Class End
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Settings Panel
	'---------------------------------------------------------------
	'*/
	Public sub LoadPanel()
		'/*
		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		'*/
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("number", ""& PLUGIN_CODE &"_OPTION_2", "Buraya Title", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write 			QuickSettings("tag", ""& PLUGIN_CODE &"_OPTION_3", "Buraya Title", "", TO_DB)
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:CachedFiles"" class=""btn btn-sm btn-primary"">"
			.Write "        	Önbelleklenmiş Dosyaları Göster"
			.Write "        </a>"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=DELETE:CachedFiles"" class=""btn btn-sm btn-danger"">"
			.Write "        	Tüm Önbelleği Temizle"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Settings Panel
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Initialize
	'---------------------------------------------------------------
	'*/
	Private Sub class_initialize()
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: PluginTemplate Main Variables
		'-----------------------------------------------------------------------------------
		'*/
    	PLUGIN_CODE  			= "ZIPFILE_MANAGER"
    	PLUGIN_NAME 			= "Zip File Management Plugin"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Zip-File-Management-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-archive"
    	PLUGIN_CREDITS 			= "Coded By @RCDMK <rcdmk@rcdmk.com> ReDeveloped By @badursun Anthony Burak DURSUN - The MIT License (MIT)"
    	PLUGIN_FOLDER_NAME 		= "Zip-File-Management-Plugin"
    	PLUGIN_DB_NAME 			= ""
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= True
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
		'/*
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
		'*/
		'Create the blank file structure
		BlankZip 			= Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
		'http://msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
		NoInterfaceYesToAll = 4 or 16 or 1024
		
		Set fso = createObject("scripting.filesystemobject")
		Set files = createObject("Scripting.Dictionary")
		
		Set zipApp = CreateObject("Shell.Application")
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: Register Plugin to CMS
		'-----------------------------------------------------------------------------------
		'*/
		class_register()
		'/*
		'-----------------------------------------------------------------------------------
		' REQUIRED: Hook Plugin to CMS Auto Load Location WEB|API|PANEL
		'-----------------------------------------------------------------------------------
		'*/
		If PLUGIN_AUTOLOAD_AT("WEB") = True Then 
			' Cms.BodyData = Init()
			' Cms.FooterData = "<add-footer-html>Hello World!</add-footer-html>"
		End If
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Initialize
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Terminate
	'---------------------------------------------------------------
	'*/
	Private sub class_terminate()
		Set curArquieve = nothing
		Set zipApp = nothing
		Set files = nothing
		' If we created the file but did not saved it, delete it since its empty
		If created and not saved then
			On Error Resume Next
			fso.deleteFile m_path
			On Error Goto 0
		End If
		Set fso = nothing
	End Sub
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Class Terminate
	'---------------------------------------------------------------
	'*/
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Manager Exports
	'---------------------------------------------------------------
	'*/
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'/*
	'---------------------------------------------------------------
	' REQUIRED: Plugin Manager Exports
	'---------------------------------------------------------------
	'*/

	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
	Public Property Get Count()
		Count = files.Count
	End Property
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
	Public Property Get Path
		Path = m_path
	End Property
	'/*
	'---------------------------------------------------------------
	' Opens or creates the arquieve
	'---------------------------------------------------------------
	'*/
	Public Sub OpenArquieve(byval path)
		Dim file
		path 	= replace(path, "/", "\")
		m_path 	= Server.MapPath(path)
		
		If not fso.fileexists(m_path) then
			Set file = fso.createTextFile(m_path)
				file.write BlankZip
				file.close()
			Set file = Nothing
			Set curArquieve = zipApp.NameSpace(m_path)
			created = true
		Else
			Dim cnt
			Set curArquieve = zipApp.NameSpace(m_path)
			cnt = 0
			For Each file in curArquieve.Items
				cnt = cnt + 1
				files.add file.path, cnt
			Next
		End If
		saved = false
	End Sub
	'/*
	'---------------------------------------------------------------
	' Add a file or folder to the list
	'---------------------------------------------------------------
	'*/
	Public Sub Add(byval path)
		path = replace(path, "/", "\")		
		If instr(path, ":") = 0 Then path = Server.mappath(path)
		
		If Not fso.fileExists(path) AND Not fso.folderExists(path) Then
			err.raise 1, "File not exists", "The input file name doen't correspond to an existing file"
			
		ElseIf Not files.exists(path) Then
			files.add path, files.Count + 1
		End If
	End Sub
	'/*
	'---------------------------------------------------------------
	' Remove a file or folder from the to be added list (currently it only works for new files)
	'---------------------------------------------------------------
	'*/
	Public Sub Remove(byval path)
		If files.exists(path) Then files.Remove(path)
	End Sub
	'/*
	'---------------------------------------------------------------
	' Clear the to be added list
	'---------------------------------------------------------------
	'*/
	Public Sub RemoveAll()
		files.RemoveAll()
	End Sub
	'/*
	'---------------------------------------------------------------
	' Writes the to the arquieve
	'---------------------------------------------------------------
	'*/
	Public Sub CloseArquieve()
		Dim filepath, file, initTime, fileCount
		Dim cnt
		cnt = 0
		For Each filepath In files.keys
			If instr(filepath, m_path) = 0 Then
				curArquieve.Copyhere filepath, NoInterfaceYesToAll
				fileCount = curArquieve.items.Count
				On Error Resume Next
				'Do Until fileCount < curArquieve.Items.Count
					wscript.sleep(10)
					cn = cnt + 1
				'Loop
				On Error GoTo 0
			End If
		Next
		saved = true
	End Sub
	'/*
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	'*/
	Public Sub ExtractTo(byval path)
		If typeName(curArquieve) = "Folder3" Then
			path = Server.MapPath(path)
			
			if not fso.folderExists(path) then
				fso.createFolder(path)
			end if
			
			zipApp.NameSpace(path).CopyHere curArquieve.Items, NoInterfaceYesToAll
		end if
	End Sub
End Class 
%>
