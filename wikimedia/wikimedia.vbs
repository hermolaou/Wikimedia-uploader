'Glory to God in the highest, and peace on the earth, in humans benevolence.
'==============================================================================

'
' MediaWiki API.
' Glory to God for everything.
'

Const mediaWikiUA = "Wikimedia uploader/page editor (kostjaermolaev@gmail.com)"

' Set user-agent for API calls.
userAgent = mediaWikiUA

Dim runAllTasks


Const wikimediaPath = "d:\wikimedia"

tasksFile=fso.buildpath(scriptPath, "tasks.xml")
tasksFile="tasks.xml"


Set fieldRusLabels = newdict
Set fieldTypes = newdict
Set selectOptions= newdict

fieldTypes("action")	="hidden"
fieldTypes("title")		="text"
fieldTypes("subtitle")	="hidden"
fieldTypes("author")	="text"
fieldTypes("editor")	="text"
fieldTypes("translator")="text"
fieldTypes("date")		="text"
fieldTypes("city")		="hidden"
fieldTypes("source")	="text"
fieldTypes("language")	="select"
fieldTypes("license")	="select"
'fieldTypes("file")="hidden"
fieldTypes("text")	="hidden"
fieldTypes("description")="text"
fieldTypes("category")	="text"
fieldTypes("pageid")	="hidden"

fieldRusLabels("file")			= 	"Файл"
fieldRusLabels("title")			=	"Название"
fieldRusLabels("subtitle")		=	"Подзаголовок"
fieldRusLabels("author")		=	"Автор"
fieldRusLabels("editor")		=	"Редактор"
fieldRusLabels("description")	=	"Описание"
fieldRusLabels("date")			=	"Дата"
fieldRusLabels("source")		=	"Источник"
fieldRusLabels("language")		=	"Язык"
fieldRusLabels("license")		=	"Лицензия"
fieldRusLabels("category")		=	"Категория"	
fieldRusLabels("translator")	=	"Переводчик"			

selectOptions("language")=Array("ru", "el", "en", "fr", "la")
selectOptions("license")=Array("PD-old-1923", "PD-1996", "PD-Russia-2008", "PD-Russia")


Set tasks = CreateObject("adodb.recordset")

If fso.FileExists(tasksFile) Then
	
	tasks.Open tasksFile


Else

	With tasks.Fields
		.Append "action", adVarChar, 12	
			.Append "file", adVarWChar, 1000
			.Append "author", adVarWChar, 100
		.Append "title", adVarWChar, 200
		.Append "date", adBSTR, 15
		.Append "language", adVarChar, 10

		.Append "source", adVarWChar, 500
		.Append "license", adVarChar, 20	
		.append "category", adVarWChar, 1000
		.append "description", adVarWChar, 2000

	
		.Append "editor",adVarWChar, 100

	
		.Append "pageid", adVarChar, 15
			.Append "text", adVarWChar, 5000
	
	End With
	tasks.Open
End If


Function RunTask()
	'assuming that tasks record poiner is to the next task.

'categories to separate them with ;
'not to expose 'undefined' values
	
	'alert "RunNextTask needs attention. see comment above."
		
	action=tasks.fields("action")
	Select Case action
		Case "edit"
		'!!! pageid or title. check what is given.
			params=	"minor=1&bot=1"
			For Each field In tasks.fields
				If Len(Trim(field)) And field<>"undefined" then
					params = params + "&" + field.name + "=" + encodeURIComponent(field)
				End If
			Next
						
		Case "upload"
			
			If Not(fso.FileExists(tasks.fields("file"))) Then
				'echo "File not exists", tasks.fields("file")
				RunNextTask
				Exit Function
			End If
		
					
			filedesc = "{{Book" & vbLf
			For Each field In Array("author", "title", "editor", "date", _
									"language", "source", "description")
				If Len(Trim(tasks.fields(field))) And tasks.fields(field)<>"undefined" then
					filedesc = filedesc & "|" & field & "=" & tasks.fields(field) & vbLf
				End if
			Next
			filedesc = filedesc & "}}"
						
			For Each category In Split(tasks.fields("category"), ";")
				If regextest(category, "\[\[Category:.+\]\]") Then
					categories = categories & vbLf & category
				Else
					categories = categories & vbLf & "[[Category:" & category & "]]"
				End If
			Next
		
			license="{{" & tasks.fields("license") & "}}"
			
			wikitext = "=={{int:filedesc}}==" & vbLf & _
					filedesc & vbLf & vbLf & _
					"=={{int:license-header}}==" & vbLf & _
					license & vbLf & vbLf & _
					categories
					
			Set params=newdict
			params("text")=wikitext
			params("filename")=tasks.fields("title") & "." & fso.GetExtensionName(tasks.fields("file"))
			Set params("file")=fso.GetFile(tasks.fields("file"))
			
		'	MsgBox wikitext
			
			
	End Select
	
	result = MediaWikiAPI (action, params)
	
End Function


'=
'=
'=
Function MediaWikiAPI(action, params)

	Const apiUrlBase = "https://commons.wikimedia.org/w/api.php?format=json&action="
	
	addr= apiUrlBase & action

	If dbg Then	
		echo "API action " & action
		'dump params
	End If

	Select Case action
		Case "login"
		
			Set resp = JSON.parse(httpPost(addr, params))
		
		case "edit", "move"
			'==============
			'asynchronous.
			'==============
			If csrfToken="" Then GetCSRFToken
			params = params & "&assert=user&token=" & encodeURIComponent(csrfToken)
			
			MediaWikiAPI= httpPostAsync(addr, params, GetRef("xmlhttp_OnReadyStateChange"))
			Exit function
			
		Case "upload"
			'=============================
			' upload is asynchronous
			'=============================
			If csrfToken="" Then GetCSRFToken
			params("token")=csrfToken
			
			With xmlhttp
			
		     	.open "POST", addr, True 
		    	.setRequestHeader "User-Agent", userAgent
		    	.setRequestHeader "Content-Type", "multipart/form-data; boundary=" + MPFDBoundary
		    	
		    	'set request header cookies.	
		    	domain=htmla(addr).hostname
		    	path=htmla(addr).pathname
		    	cookiesHeader=GetCookies(domain, path)
		    	If Len(cookiesHeader) Then
			   		.setRequestHeader "Cookie", cookiesHeader
				End If
		    	
		    	SetORSC		'Set OnReadyStateChange
		   
		    	MediaWikiAPI = .send (BuildFormData(params))
		    
 			 End With
 			 Exit function
	
		Case "query"
			addr = addr & "&" & params
			Set resp = JSON.parse(httpGet(addr))
						
		Case Else
			arguments = arguments & "&assert=user"
			addr = addr & "&" & params
			Set resp = JSON.parse(httpGet(addr))
			
	End Select
	
	set MediaWikiAPI = JSHelperErrorCheck (resp)

End Function




Dim csrfToken

Function GetCSRFToken()
	Set resp = MediaWikiAPI("query", "meta=tokens&assert=user")
	
	csrfToken  = resp.query.tokens.csrftoken
'	SaveSetting "csrfToken", , csrfToken
End Function

'	Set wShell = CreateObject("WScript.Shell")
'	Set shApp = CreateObject("Shell.Application")
'	Set fso = CreateObject("scripting.filesystemobject")
'	Set xmlHttp = CreateObject("MSXML2.Ser verXMLHTTP.6.0")
'	Set rs=CreateObject("ado db.recordset")





Sub SaveTasks()

	MsgBox "SaveTasks needs to be corrected to save tasks.xml only in the script path"
	bakFile=fso.BuildPath(fso.GetParentFolderName(tasksfile), "tasks.bak")

	tasks.Filter=""
	tasks.Save tasksFile, adPersistXML
	
	fso.CopyFile tasksfile, bakFile, True
	
	tasks.Close
	tasks.Open tasksFile
	
End Sub


Function UploadFolder(oFolder)

	For Each fItem In oFolder.Items
		If (fItem.IsFolder) Then
			'... upload subfolder.
			UploadFolder fItem.GetFolder
		Else
			UploadFile fItem, oFolder.Title
		End if
	Next
	
	'save recordset.
	'savetasks

End Function

Function UploadFile(file, category)
	
	If file.IsLink Then
		Set url=file
		urltitle=fso.GetBaseName(file.Name)
		Set file = file.GetLink.Target
	End If
	
	If Not(fso.FileExists(file.Path)) Then
		If MsgBox( "Upload doesn't exist, delete link?" & vbLf &  file.Path, vbYesNo) Then
			fso.DeleteFile(url.path)
		End If
		Exit function
	End If
	
	ext=fso.GetExtensionName(file.name)
	If ext<>"pdf" And ext<>"djvu" Then Exit function
	
	if file.Size > (100*1024*1024) Then
		'>100 mb needs to be uploaded in chunks.
		echo "File too big:", file.name
		If IsObject(url) Then fso.DeleteFile url.path
		Exit Function
	End if

	'title of url. filename as 'title' can all be used to search.
	
	filename=fso.GetFileName(file.Name)
	
	criteria=array("file like '*" & urltitle & "*'", _
		"file like '*" & basename & "*'", _
		"title like '*" & urltitle & "*'", _
		"title like '*" & basename & "*'")
	
	If Not(tasks.BOF) Then
	
		'see if this file is already in tasks.
		crit="file like '*" & filename & "*'"
		tasks.MoveFirst
			
		echo crit
		tasks.Find crit
		If Not(tasks.EOF) And Not(tasks.BOF)  Then
			'On Error Goto 0
			'found such file
			fileAlreadyInTasks=true
		End If
		On Error Goto 0
		
'		For Each criterion In criteria
'			tasks.MoveFirst
			
'			On Error Resume Next
'			tasks.Find criterion
'			If Not(Err) And Not(tasks.EOF)  Then
'				On Error Goto 0
'				msg=format("Found by criterion ", criterion, vbLf, _
'							tasks.GetString(,1), vbLf, "updating it.")
				
'				echo msg			
'				tasks.Fields("file")=file.path
'				tasks.Fields("category")=category
'				tasks.Update
'				Exit Function
				
'			End If
'			On Error Goto 0
'		Next
	End if
	
	'not found task with that file or title, add new
	If Not(fileAlreadyInTasks) then
		tasks.addnew Array("action", "file", "category"), _
			array("upload", file.Path, category)
			
		tasks.Update
	End if
	
	
End Function
	

'Check if everything is all right with tasks, if all files exist.
Function CheckTasks()
	If tasks.bof Then Exit Function 	
 	tasks.MoveFirst
  
 	While Not(tasks.eof)
 	
 		If tasks.fields("action")="upload" then
 
			If fso.FileExists(tasks.fields("file"))=False Then
	
				altFile=fso.BuildPath(fso.GetParentFolderName(tasks.fields("file")), _
										tasks.fields("title") & "." & _
										fso.GetExtensionName(tasks.fields("file")))
										
				If fso.FileExists(altFile) then
					tasks.fields("file")=altFile
					modified=true
				
				Else	
					tasks.delete
					
				End If
 			End If
 		End If
 		
 		tasks.MoveNext
 	Wend
 	
 	If modified Then SaveTasks
	
End function
