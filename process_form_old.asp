<%@ Language=VBScript %>
<%
	' DEFAULT values
	
	' Mail relay host - outgoing mail server address. 
	dim mailRelayHost
	dim Email_address
	dim Email_name
	dim mail
	dim errorMessage
	
	mailRelayHost = "mail.thenewgroup.com"
	'mailRelayHost = "newint01.newinteractive.com"
	

	Email_address = "Troy-Weller@hoffmancorp.com"
	Email_name = "Troy Weller"
	
	'set form vars
	
	errorMessage = ""
	noResultsMessage = "We are unable to deliver your email request at this time. Please try again later.<br>"
	resorts = ""
	
	' Create the email message body
	messageBody = "First name: " & request.form("firstname") & vbcrlf &_
			"Last name: " & request.form("lastname") & vbcrlf &_
			"Address: " & request.form("address") & vbcrlf &_
			"City: " & request.form("city") & vbcrlf &_
			"State: " & request.form("state") & vbcrlf &_
			"Zip: " & request.form("zipcode") & vbcrlf &_
			"Phone: " & request.form("phone") & vbcrlf &_
			"Condominium: " & request.form("condo") & vbcrlf &_
			"Comments: " & request.form("comments") & vbcrlf
	
	' Use a fake email address if the user did not provide one
	if (request.form("email") = "") then
		fromEmail = "website@precision-construction-company.com"
	else
		fromEmail = request.form("email")
	end if
	
	'Set mail = Server.CreateObject("SMTPsvg.Mailer")
	'mail.RemoteHost = mailRelayHost
	'mail.FromName = request.form("firstName") & " " & request.form("lastName")
	'mail.FromAddress = fromEmail
	'mail.AddRecipient Email_name, Email_address
	'mail.Subject = "Precision Construction Contact Form Submission"
	'mail.BodyText = messageBody
	
	' Send the email message. If there is an error, set the error message
	'if not mail.SendMail then
		'errorMessage = noResultsMessage
		'Response.Write("!!!")
	'end if

	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = "Precision Construction Contact Form Submission"
	objMessage.From = fromEmail
	objMessage.To = Email_address
	objMessage.TextBody = messageBody
	objMessage.Send
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>About Precision Construction Company</title>


<link href="precision.css" rel="stylesheet" type="text/css" /></head>

<body>
<table width="628" border="0" align="center" cellpadding="0" cellspacing="0">
 
  <tr>
    <td width="12" height="17"></td>
    <td width="367"></td>
    <td width="184"></td>
    <td width="65" rowspan="3" align="right" valign="top" class="padded_cell"><p>&nbsp;</p>      <a href="About_Us.html">ABOUT US </a><br />
      <span class="style6">CONTACT US</span> </td>
  </tr>
  <tr>
    <td height="52"></td>
    <td valign="top" class="padded_cell"><br />
      <a href="index.html"><img src="images/home/logo.gif" width="308" height="19" border="0" /></a></td>
  <td></td>
  </tr>
  <tr>
    <td height="5"></td>
    <td></td>
    <td></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#F1F1F1" class="graybar">
  
  <tr>
    <td width="683" height="17"  id="graybar" ></td>
  </tr>
</table>
<tr>
				
				
				
</tr>
<table width="632" height="172" border="0" align="center" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
 
  <tr valign="top">
    <td width="15" rowspan="3" valign="top"><img src="images/spacer.gif" height="1" width="1" /></td>
    <td width="383" rowspan="2" valign="top" class="padded_cell2"><span class="style6"><a href="our_advantage.html">OUR ADVANTAGE</a></span>   <img src="images/spacer.gif" width="25" height="1" />|<img src="images/spacer.gif" width="25" height="1" />  <span class="style6"><a href="Construction.html">CONSTRUCTION</a> </span><img src="images/spacer.gif" width="25" height="1" />|<img src="images/spacer.gif" width="25" height="1" />  <span class="style6"><a href="condominium.html">CONDOMINIUMS</a></span><a href="construction.html"> <br />
    </a>
      <span class="style7"><br />
      <br />
        Thanks!</span><br />
        Your information has been successfully submitted. We will be in touch with you shortly. Click here to return to the <a href="index.html">home page</a>. </td>
    <td width="11" height="145"></td>
    <td width="221" valign="top" class="picture3"><img src="images/contact_us/img-main.jpg" width="221" height="145" border="0" /></td>
  </tr>
  <tr valign="top">
    <td height="16"></td>
    <td></td>
  </tr>
  
  <tr valign="top">
    <td height="9"></td>
    <td></td>
    <td></td>
  </tr>
  <tr valign="top">
    <td height="12"></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
</table>

</body>
</html>