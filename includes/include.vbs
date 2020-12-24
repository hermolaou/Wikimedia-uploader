'
' Gloria in excelsis Dei, et in terra pax, hominibus bonae voluntatis.
'


'====================================================================


'====================================================================

'Returns file contents As a binary data
Function ReadFile(File)
	Dim Stream: Set Stream = CreateObject("ADODB.Stream")
	Stream.Type = 1 'Binary
	Stream.Open

	on error resume next
	Stream.LoadFromFile File.Path
	if err then
		echo "Reading file error:", File.Path, vblf, err.Description
		ReadFile=null
	else
		ReadFile = Stream.Read	
	end if
	on error goto 0
	
	'WScript.Echo Stream.read
	Stream.Close
	Set Stream=nothing
End Function

'utf-8 write
'Sub WriteFile(file)

'	Set inStream=WScript.CreateObject("ADODB.Stream")
'	Set outStream=WScript.CreateObject("ADODB.Stream")
							
'	inStream.Open
'	inStream.type=adTypeBinary
'	inStream.LoadFromFile(file)
'	inStream.Position=3
	
'	outStream.Open
'	outStream.type=adTypeBinary
'	outStream.Write instream.Read
	
'	inStream.Close()
'	outStream.SaveToFile file, adSaveCreateOverWrite						
			
'	outstream.Close
'End sub	

'

Function dump(var)
	Set table=newtable
	For Each key In var
		table.appendChild NewTR(Array(key, var(key)))
	Next
	document.body.insertbefore table,tabs
End Function

'======

Function newdict()
	Set newdict = CreateObject("scripting.dictionary")
End Function


If Len(settingsFile)=0 Then
	Const settingsFile="./settings.inc"
ElseIf Not(fso.FileExists(settingsFile)) Then
	call fso.CreateTextFile (settingsFile,, True)
End if

Function SaveSetting(name, key, value)
	Dim settings, nokey	'to ensure there isn't a global of this name
		
	'find and replace or append
			
	On Error Resume Next
	a=(key="")
	If Err or key="" Then nokey=True
	On Error Goto 0
		
	If nokey Then
		setting=name & "="
		settingPattern = "(" & name & "="").*"""
			
		'also assign the variable immediately:
		statement= format(name, "=""", value, """")	
	Else	
		setting=name & "(""" & key & """)="
		settingPattern = "(" & name & "\(""" & key & """\)="").*"""
		decl = format("Set ", name, "=newdict")
		
		'also assign the variable immediately:
		If IsObject(Eval(name))=0 Then
			'declaration not found, declare the variable now.
			ExecuteGlobal decl
		End if
		statement= format(name, "(""", key, """)=""", value, """")
	End If
	ExecuteGlobal statement
	
	Set oSF=fso.OpenTextFile(settingsFile, ForReading,, TristateTrue)
	If oSF.AtEndOfStream=False Then settings=oSF.ReadAll
	osf.Close
	
	If InStr(settings, setting) Then
		settings=regexreplace(settings, settingPattern, "$1" & value & """")
	Else
		settings = settings & setting & """" & value & """" & vbCrLf
	End If
	
	If Len(decl) and InStr(settings, decl)=0 Then
		declPattern=format("(Set .+=newdict[\S\s]+Set .+=newdict)\s")
		settings=decl & vbLf & settings
	End If
		
	set oSF=fso.OpenTextFile(settingsFile, ForWriting,, TristateTrue)
	oSF.Write settings
	oSF.Close
	
End Function

'=================================================
'Checks if there is a connected network adapter.
'=================================================
Function CheckInternet()
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)
	For Each objItem in colItems
   		If objItem.NetConnectionStatus=2 Then
   			CheckInternet=True
   			Exit function
   		End If
    Next
    CheckInternet=False
End function

'================================================================================
'SetLocale("ru-ru")

'wshell


'
'= So weird :) 
' How to otherwise get the path of a Shell folder object?
'
Function FolderPath(shFolder)
	FolderPath = shfolder.ParentFolder.ParseName(shFolder.title).Path
End Function


'Includes another script.
'Sub Include(file)
'	Set fso=CreateObject("scripting.filesystemobject")
'		
'	If fso.GetDriveName(file)="" Then
'		file = fso.BuildPath(scriptPath, file)
'	End if

'	If fso.GetFileName(file)="adovbs.inc" Then
'		UnicodeState=TristateFalse
'	Else	
'		UnicodeState=TristateTrue
'	End If

	'alert file
'	On Error Resume Next
'	code = fso.OpenTextFile(file,,, UnicodeState).ReadAll()
'	If Err Then
'		call msgbox format("Include file ", file, " error.", vbLf, Err.Description)
'	End If
'	
'	If fso.GetFileName(file)<>"strings.vbs" Then
'		code=RegExReplace(code, "^<%|%>$", "")
'	End If
'	
'	ExecuteGlobal (code)
'End Sub



'====================================

