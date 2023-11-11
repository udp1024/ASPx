<%@ language="VBScript" %>
<%
	Option Explicit
	  
	Dim obj_ASPError, b_ErrorWritten 
	Dim str_HTML, str_URL, str_Method, l_Pos, str_QueryString, strKey
	Dim obj_Mail, str_EmailToList, str_EmailSubject, str_EmailFrom, str_EmailCCList
	Dim wsh_Network, str_MachineName
	
	Const l_MaxFormBytes = 200
	
	<!-- Change Email Info -->
	str_EmailFrom = "IPSentry@kubra.com"
	str_EmailToList ="orion@kubra.com"
	str_EmailCCList ="aleslie@kubra.com; support@kubra.com;"
	str_EmailSubject ="Server Error on e-bill: "
	  
	If Response.Buffer Then
		Response.Clear
		Response.Status = "200"
	   Response.ContentType = "text/html"
	   Response.Expires = 0
	End If
	
	Set obj_ASPError = Server.GetLastError
	
	<!-- Message displayed on screen to user who experiences error. -->
	Response.Write "<html><body>"
	Response.Write "<table align=center><tr><td><h3>An unexpected error has occurred!</h3></td></tr>"
	Response.Write "<tr><td><i>This may be due to resource unavailability or a technical glitch."
	Response.Write "<br>The problem has been logged and the support team has been notified via e-mail." 
	Response.Write "<br><br>Please try your request again. We are sorry for any inconvenience this may have caused you."
	Response.Write "</i></td></tr></table>"
	Response.Write "</body></html>"
	
	<!-- HTML email message -->
	str_HTML = str_HTML & "<td width='400' colspan='2'>"
	str_HTML = str_HTML & "<font style='COLOR:000000; FONT: 8pt/11pt verdana'>"
	str_HTML = str_HTML & "<h2 style='font:8pt/11pt verdana; color:000000'>Kubra e-bill v1.0<br>"
	str_HTML = str_HTML & "HTTP 500.100 - Internal Server Error<br>"
	str_HTML = str_HTML & FormatDateTime(Now(), 1) & ", " & FormatDateTime(Now(), 3) & "</h2>"
	str_HTML = str_HTML & "<hr color='#C0C0C0' noshade>"
	str_HTML = str_HTML & "<p>Technical Information</p>"
	str_HTML = str_HTML & "<ul>"
	str_HTML = str_HTML & "<li>Error Type:<br>"
	
	str_HTML = str_HTML & obj_ASPError.Category

	If obj_ASPError.ASPCode > "" Then 
		<!-- ASP Error -->
		str_HTML = str_HTML & ", " & obj_ASPError.ASPCode
		str_HTML = str_HTML & " (0x" & Hex(obj_ASPError.Number) & ")" & "<br>"
		str_HTML = str_HTML & "<b>" & obj_ASPError.Description & "</b><br>"
		str_HTML = str_HTML & obj_ASPError.ASPDescription & "<br>"
		str_HTML = str_HTML & obj_ASPError.File
	
		If obj_ASPError.Line > 0 Then 
			str_HTML = str_HTML & ", line " & obj_ASPError.Line
		End if
		If obj_ASPError.Column > 0 Then 
			str_HTML = str_HTML & ", column " & obj_ASPError.Column
		End if

		str_HTML = str_HTML & "<br>"
		str_HTML = str_HTML & "<font style=''COLOR:000000; FONT: 8pt/11pt courier new''><b>"
		str_HTML = str_HTML & Server.HTMLEncode(obj_ASPError.Source) & "<br>"
		If obj_ASPError.Column > 0 Then 
			str_HTML = str_HTML & String((obj_ASPError.Column - 1), "-") & "^<br>"
		End If
	  	str_HTML = str_HTML & "</b></font>"

	Else
		<!-- Other cateogory of error -->
		str_HTML = str_HTML & ", (0x" & Hex(obj_ASPError.Number) & ")" 
		str_HTML = str_HTML & "<br> <b>" & obj_ASPError.Description & "</b><br>"
		str_HTML = str_HTML & obj_ASPError.File
		If obj_ASPError.Line > 0 Then 
			str_HTML = str_HTML & ", line " & obj_ASPError.Line
		End If
		If obj_ASPError.Column> 0 Then 
			str_HTML = str_HTML & ", column " & obj_ASPError.Column
		End If
		str_HTML = str_HTML & "</b><br>"
	End If
	
	str_HTML = str_HTML & "</li>"
	str_HTML = str_HTML & "<p><li>Browser Type:<br>"
	str_HTML = str_HTML & Request.ServerVariables("HTTP_USER_AGENT") & " "
	str_HTML = str_HTML & "</li>"
	
	str_HTML = str_HTML & "<p><li>Page:<br>"
	str_Method =Request.ServerVariables("REQUEST_METHOD")
	str_HTML = str_HTML & str_Method & " "
	If str_Method = "POST" Then
		str_HTML = str_HTML & Request.TotalBytes & " bytes to "
	End If
	str_HTML = str_HTML & Request.ServerVariables("SCRIPT_NAME")
	l_Pos = InStr(Request.QueryString, "|")
	If l_Pos > 1 Then
		str_HTML = str_HTML & "?" & Left(Request.QueryString, (l_Pos - 1))
	End If
	str_HTML = str_HTML & "</li>"
	
	If str_Method = "POST" Then
		str_HTML = str_HTML & "<p><li>POST Data:<br>"
		If Request.TotalBytes > l_MaxFormBytes Then
			str_HTML = str_HTML & Server.HTMLEncode(Left(Request.Form, l_MaxFormBytes)) & " . . ."
		Else
			str_HTML = str_HTML & Server.HTMLEncode(Request.Form)
		End If
		str_HTML = str_HTML & "</li>"
	End If
	
	str_HTML = str_HTML & "<p>"
	str_HTML = str_HTML & "<li>Kubra e-bill Info:<br>"
	str_HTML = str_HTML & "UserID <b>" & Session("LogonID") &"</b><br>"
	str_HTML = str_HTML & "Account <b>" & Session("Merchant") &"</b><br>"
	str_HTML = str_HTML & "</li></p>"

	str_HTML = str_HTML & "<p><li>More information:<br>"
	str_QueryString = "prd=iis&sbp=&pver=5.0&ID=500;100&cat=" & _
	Server.URLEncode(obj_ASPError.Category) & _
	"&os=&over=&hrd=&Opt1=" & Server.URLEncode(obj_ASPError.ASPCode) & "&Opt2=" & _
	Server.URLEncode(obj_ASPError.Number) & _
	"&Opt3=" & Server.URLEncode(obj_ASPError.Description) 
	str_URL = "http://www.microsoft.com/ContentRedirect.asp?" & str_QueryString
	
	str_HTML = str_HTML & "<a href=" & str_URL & ">Microsoft Support</a>"

	str_HTML = str_HTML & "<p><li>Server Variables:<br>"
	For Each strKey In Request.ServerVariables 
		str_HTML = str_HTML & "<b>" & strKey &" </b>" & Request.ServerVariables(strKey) & " <br>"
	Next

	<!-- Get Machine name for diplay in Email message -->
	Set wsh_Network = server.CreateObject("WScript.Network") 
	str_MachineName=wsh_Network.ComputerName

	<!-- Send Email message -->
	Set obj_Mail = Server.CreateObject("CDONTS.NewMail")
	
	obj_Mail.To = str_EmailToList 
	obj_Mail.From = str_EmailFrom 
	obj_Mail.CC = str_EmailCCList 
	obj_Mail.Subject = str_EmailSubject & str_MachineName
	obj_Mail.BodyFormat=0
	obj_Mail.MailFormat=0
	obj_Mail.Body = str_HTML
	obj_Mail.Send
	Set obj_Mail = Nothing
	
	Response.End
%>
