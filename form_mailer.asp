<%@ Language=VBScript %>
<%
	' DEFAULT values
	
	' Mail relay host - outgoing mail server address. 
	dim mailRelayHost
	mailRelayHost = "newint01.newinteractive.com"
	
	Email_address = "someone@precision-construction-company.com"
	Email_name = "Sales"
	
	'set form vars
	dim mail
	
	errorMessage = ""
	noResultsMessage = "We are unable to deliver your email request at this time. Please try again later.<br>"
	resorts = ""
	
	' Create the email message body
	messageBody = "Sales Inquiry" & vbcrlf & vbcrlf &_
			"First name: " & request.form("firstName") & vbcrlf &_
			"Last name: " & request.form("lastName") & vbcrlf &_
			"Address 1: " & request.form("addr1") & vbcrlf &_
			"Address 2: " & request.form("addr2") & vbcrlf &_
			"City: " & request.form("city") & vbcrlf &_
			"State: " & request.form("state") & vbcrlf &_
			"Zip: " & request.form("zip") & vbcrlf &_
			"Country: " & request.form("country") & vbcrlf &_
			"E-mail address: " & request.form("email") & vbcrlf &_
			"Home phone: " & request.form("homePhone") & vbcrlf &_
			"Business phone: " & request.form("businessPhone") & vbcrlf &_
			"Best time to contact: " & request.form("contactTime") & vbcrlf &_
			"Contact by: " & request.form("contactBy") & vbcrlf & vbcrlf &_
			"Would like info about: " & resorts & vbcrlf &_
			"Number of bedrooms: " & request.form("bedrooms") & vbcrlf &_
			"Activities interested in: " & request.form("activities") & vbcrlf &_
			"Are they a Timeshare owner? " & request.form("timeown") & vbcrlf & vbcrlf &_
			"Which Timeshare(s) do they own: " & request.form("timeshares") & vbcrlf
			
	
	' Use a fake email address if the user did not provide one
	if (request.form("email") = "") then
		fromEmail = "website@precision-construction-company.com"
	else
		fromEmail = request.form("email")
	end if
	
	Set mail = Server.CreateObject("SMTPsvg.Mailer")
	mail.RemoteHost = mailRelayHost
	mail.FromName = request.form("firstName") & " " & request.form("lastName")
	mail.FromAddress = fromEmail
	mail.AddRecipient Email_name, Email_address
	mail.Subject = "Web Site: Sales Inquiry"
	mail.BodyText = messageBody
	
	' Send the email message. If there is an error, set the error message
	if not mail.SendMail then
		errorMessage = noResultsMessage
	end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Precision Construction Company</title>
</head>

<body>
Thanks!
</body>
</html>