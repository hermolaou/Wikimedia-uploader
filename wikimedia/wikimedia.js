/*
Glory to the Only True God in the Highest.
*/

function SetORSC()
{
	xmlhttp.onreadystatechange=xmlhttp_OnReadyStateChange;
}

function xmlhttp_OnReadyStateChange()
{
	if ( xmlhttp.readyState != READYSTATE_COMPLETE ) return;
	
	SetCookies (xmlhttp.getAllResponseHeaders, "commons.wikimedia.org", "/w/api.php")
	
	respText = xmlhttp.responseText
	
	if ( respText.match(/Request Entity Too Large/im)) {
		//file larger than allowed (100mb)
		echo ('File larger than 100mb:')
   		RunNextTask();
		return
	}
	
	var resp=JSON.parse(respText);
	
	if ( resp.upload.result=="Success") {
		
		echo ("Success");
		
		tasks.delete();
		tasks.update()
		SaveTasks()
		
	} else {
     	echo (respText);
	}
	
	RunNextTask();
	
}

function RunNextTask()
{
 	if ( runAllTasks ) {
		tasks.movenext()
		if ( tasks.eof)
			echo('Выполнено.');
		else
			RunTask()
	}
}


function WikimediaLogin()
{
	var resp = MediaWikiAPI("query", "meta=tokens&type=login")
	alert(JSON.stringify(resp));

	lgToken = resp.api.query.tokens.getAttribute("logintoken")
	lgToken = encodeURIComponent(lgToken)
					
	if (lgName=="" || lgPassword=="")	{
		//'need credentials
		alert("Need credentials")
	}
	var resp = MediaWikiAPI("login", format("lgname=", lgName, "&lgpassword=", lgPassword, "&lgtoken=", lgToken))

	var login = resp.api.login
	loginResult  = login.getAttribute("result");
	if (loginResult != "Success") {
		echo ("Login not success.")
		echo (JSON.stringify(resp));
	}
}

function JSHelperErrorCheck(resp)
{
	if (!resp.hasOwnProperty ('error')) return resp;
	var errCheck = resp.error
		
	errCode = errCheck(0).getAttribute("code")
	switch (errCode)
	{
		case "assertuserfailed":
			//'Not logged in. Let's log in.
			WikimediaLogin()
			return MediaWikiAPI(action, params)
		
		case "missingtitle":
			//it's not a problem,
			//the calling side should be prepared to handle this.
			return resp
			
		default:
			Echo ("Error on wikimedia api request.",vbLf, "action:", action, vblf, "params:", params)
			echo (resp.xml)
			
	}
	return resp
}
