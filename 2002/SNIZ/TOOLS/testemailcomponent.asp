<% @ Language="VBScript" %>
<% Option Explicit %>
<%
Dim theComponent(9)
Dim theComponentName(9)
Dim theotherComponent(1)
Dim theotherComponentName(1)

' the components
theComponent(0) = "SMTPsvg.Mailer"
theComponent(1) = "Persits.MailSender"
theComponent(2) = "SMTPsvg.Mailer"
theComponent(3) = "CDONTS.NewMail"
theComponent(4) = "CDONTS.NewMail"
theComponent(5) = "dkQmail.Qmail"
theComponent(6) = "Geocel.Mailer"
theComponent(7) = "iismail.iismail.1"
theComponent(8) = "Jmail.smtpmail"
theComponent(9) = "SmtpMail.SmtpMail.1"

'the other components
theotherComponent(0) = "ocxQmail.ocxQmailCtrl.1"
theotherComponent(1) = "SoftArtisans.SMTPMail"

' the name of the components
theComponentName(0) = "ASPMail"
theComponentName(1) = "ASPEMail"
theComponentName(2) = "ASPQMail"
theComponentName(3) = "CDONTS"
theComponentName(4) = "Chili!Mail"
theComponentName(5) = "dkQMail"
theComponentName(6) = "GeoCel"
theComponentName(7) = "IISMail"
theComponentName(8) = "JMail"
theComponentName(9) = "SMTP"

'the name of the other components
theotherComponentName(0) = "OCXQMail"
theotherComponentName(1) = "SA-Smtp Mail"


Function IsObjInstalled(strClassString)
 On Error Resume Next
 ' initialize default values
 IsObjInstalled = False
 Err = 0
 ' testing code
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 ' cleanup
 Set xTestObj = Nothing
 Err = 0
End Function
%>

<HTML>
<HEAD>
<TITLE>发信组件测试</TITLE>
</HEAD>
<BODY bgColor="white" text="midnightblue" link="darkblue" aLink="red" vLink="red">
<font face="Verdana, Arial, Helvetica">
<table border=0 cellspacing=0 cellpadding=0 align=center>
  <tr valign="top">
    <td bgcolor="black">
    <table border=0 cellspacing=1 cellpadding=4>
      <tr valign="top">
        <td bgcolor="midnightblue" colspan="2" align="center"><font size="2" color="mintcream"><b>发信组件测试</b></font></td>
      </tr>
      <tr valign="top">
        <td bgcolor="midnightblue" colspan="2" align="center"><font size="2" color="mintcream">The following components are currently<br>available choices in the latest<br>release of Snitz Forums 2000</font></td>
      </tr>
        <% Dim i
           For i=0 to UBound(theComponent)
             Response.Write "<tr><td bgColor=""A0B0E0"" align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2""><strong>" & theComponentName(i) & ":&nbsp;</strong></td><td bgColor=""A0B0E0"" align=""center"">"
	     If Not IsObjInstalled(theComponent(i)) Then 
	       Response.Write "not installed"
	     Else
	       Response.Write "<strong>installed!</strong>"
	     End If
	     Response.Write "</td></font></tr>" & vbCrLf
	   Next %>
    </table>
    </td>
  </tr>
</table>
<br>
<table border=0 cellspacing=0 cellpadding=0 align=center>
  <tr valign="top">
    <td bgcolor="black">
    <table border=0 cellspacing=1 cellpadding=4>
      <tr valign="top">
        <td bgcolor="midnightblue" colspan="2" align="center"><font size="2" color="mintcream"><b>Email Component Test</b></font></td>
      </tr>
      <tr valign="top">
        <td bgcolor="midnightblue" colspan="2" align="center"><font size="2" color="mintcream">The following components are not currently<br>available choices in the latest<br>release of Snitz Forums 2000<br>but mods are available to add them.</font></td>
      </tr>
        <% Dim j
           For j=0 to UBound(theotherComponent)
             Response.Write "<tr><td bgColor=""A0B0E0"" align=""right""><font face=""Verdana, Arial, Helvetica"" size=""2""><strong>" & theotherComponentName(j) & ":&nbsp;</strong></td><td bgColor=""A0B0E0"" align=""center"">"
	     If Not IsObjInstalled(theotherComponent(j)) Then 
	       Response.Write "not installed"
	     Else
	       Response.Write "<strong>installed!</strong>"
	     End If
	     Response.Write "</td></font></tr>" & vbCrLf
	   Next %>
    </table>
    </td>
  </tr>
</table>
</font>
</BODY>
</HTML>