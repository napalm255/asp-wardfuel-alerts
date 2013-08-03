<%
	Dim objConn
	Set objConn = Server.CreateObject ("ADODB.Connection")
	objConn.Open "Provider=SQLOLEDB.1;User ID=EjwDBAdmin;Password=EjwDBAdmin;Data Source=SQL01-SRV"

	if request.querystring("m") = "es" then
		if request.querystring("a") = "upd" then
			objConn.Execute("UPDATE EmailSettings SET sWaitMinutes = " & request.form("txtEmailWait") & ", sEmailCount = " & request.form("txtEmailCount") & " WHERE sID = 1")
			response.redirect("default.asp")
		end if
	elseif request.querystring("m") = "eu" then
		if request.querystring("a") = "upd" then
			objConn.Execute("UPDATE EmailUsers SET sEmail = '" & request.form("txtDesktop") & "' WHERE sType = 'desktop'")
			objConn.Execute("UPDATE EmailUsers SET sEmail = '" & request.form("txtCell") & "' WHERE sType = 'cell'")
			response.redirect("default.asp")
		end if
	end if
' Header
	write("<html>")
	write("<head>")
	write("<link rel=stylesheet href=""/alerts/alerts.css"" type=text/css>")
	write("<title>Fuel System E-Mail Alert Manager</title>")
	write("</head>")
	write("<body>")
	write("<div id=""mainTitle"">Fuel System E-Mail Alert Manager</div>")

' E-Mail Settings
	Set eSet = objConn.Execute("SELECT * FROM EmailSettings WHERE sID = 1")
	write("<center>")
	write("<form method=""post"" action=""default.asp?m=es&a=upd"">")
	write("<div id=""emailSettingsBox"">")
	write("<div id=""emailSettingsTitle"">E-Mail Settings</div>")
	write("<div id=""emailSettingsInside"">")
	write("<table cellpadding=""2"" cellspacing=""2"" border=""0"">")
	write("<tr><td id=""emailSettingsHeader"">E-Mail Count:</td><td id=""emailSettingsHeader""><input id=""emailSettingsTextField"" type=""text"" size=""20"" name=""txtEmailCount"" value=""" & eSet("sEmailCount") & """></td></tr>")
	write("<tr><td id=""emailSettingsInfo"" colspan=""2"">'E-Mail Count' defines the number of e-mails you will recieve while the alarm is active.</td></tr>")
	write("<tr><td id=""emailSettingsHeader"">E-Mail Wait Period:</td><td id=""emailSettingsHeader""><input id=""emailSettingsTextField"" type=""text"" size=""20"" name=""txtEmailWait"" value=""" & eSet("sWaitMinutes") & """> minutes</td></tr>")
	write("<tr><td id=""emailSettingsInfo"" colspan=""2"">'E-Mail Wait Period' defines the time in minutes that an alarm has to be inactive before sending new alerts. Once the 'E-Mail Count' has been reached, this wait period must be reach before the 'E-Mail Count' gets reset.</td></tr>")
	write("</table>")
	write("<div id=""emailSettingsButton""><input type=""submit"" name=""setSubmit"" value=""Update Settings""></div>")
	write("</div>")
	write("</div>")
	write("</form>")
	write("</center>")
	Set eSet = nothing

' E-Mail Users
	Set eDsk = objConn.Execute("SELECT * FROM EmailUsers WHERE sType = 'desktop'")
	Set eCel = objConn.Execute("SELECT * FROM EmailUsers WHERE sType = 'cell'")
	write("<center>")
	write("<form method=""post"" action=""default.asp?m=eu&a=upd"">")
	write("<div id=""emailSettingsBox"">")
	write("<div id=""emailSettingsTitle"">E-Mail Addresses</div>")
	write("<div id=""emailSettingsInside"">")
	write("<table cellpadding=""2"" cellspacing=""2"" border=""0"">")
	write("<tr><td id=""emailSettingsHeader"">Desktop E-Mail Addresses:</td><td id=""emailSettingsHeader""><input id=""emailSettingsTextField"" type=""text"" size=""60"" name=""txtDesktop"" value=""" & eDsk("sEmail") & """></td></tr>")
	write("<tr><td id=""emailSettingsInfo"" colspan=""2"">'Desktop' e-mail type displays entire alarm message. The field is a comma seperated list of e-mail addresses. Make sure there is no comma at the end of the list.</td></tr>")
	write("<tr><td id=""emailSettingsHeader"">Cell E-Mail Addresses:</td><td id=""emailSettingsHeader""><input id=""emailSettingsTextField"" type=""text"" size=""60"" name=""txtCell"" value=""" & eCel("sEmail") & """></td></tr>")
	write("<tr><td id=""emailSettingsInfo"" colspan=""2"">'Cell' e-mail type displays a shortend message that is viewable on most cell phones. The field is a comma seperated list of e-mail addresses. Make sure there is no comma at the end of the list.</td></tr>")
	write("</table>")
	write("<div id=""emailSettingsButton""><input type=""submit"" name=""setSubmit"" value=""Update E-Mail Addresses""></div>")
	write("</div>")
	write("</div>")
	write("</form>")
	write("</center>")
	Set eDsk = nothing
	Set eCel = nothing

' Footer
	write("</body>")
	write("<html>")
%>

<%
 'Functions
	Function write(strText)
		response.write(strText & vbCRLF)
	End Function
%>