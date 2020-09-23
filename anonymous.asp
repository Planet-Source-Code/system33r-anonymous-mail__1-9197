<%
' Anonymous Mail!
' version 1.0
' (c) 2000 Secone
' http://www.secone.com
' info@secone.com
' file: anonymous.asp
' programming: System33r / s.r.t
<%

<% If request.form("flag")=""then %>
    <HTML>
    <H1>Anonymous Mailer by System33r</H1>
    <H4>Please enter all the required fields or it won't be send.</H4>
    <FORM action=anonymous.asp method=post>
    <TABLE border=0 cellPadding=1 cellSpacing=1 style="HEIGHT: 184px; WIDTH: 457px" 
    width="75%">
    <style>
<--
a:link { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: underlined }
a:visited { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: underlined }
a:active { color: #000080; font: 7.5pt verdana; font-weight: none; text-decoration: none }
a:hover { color: #000080; font: 7.5pt verdana; font-weight: none; text-decoration: none }
td { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: none }
body { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: none }
input { color:#000000; font: 7.5pt verdana; font-weight: none; text-decoration: none; background: #c0c0c0; border: 1 solid #000000; }
-->
</style>
    <TR>
    <TD>From: </TD>
    <TD><INPUT name=From size=30 style="HEIGHT: 22px; WIDTH: 321px"></TD></TR>
    <TR>
    <TD>To: </TD>
    <TD><INPUT name=To size=30 style="HEIGHT: 22px; WIDTH: 321px"></TD></TR>
    <TR>
    <TD>Subject: </TD>
    <TD><INPUT name=Subject size=50 style="HEIGHT: 22px; WIDTH: 321px"></TD></TR>
    <TR>
    <TD>Body of Message: <BR></TD>
    <TD><TEXTAREA cols=30 name=Body rows=5 style="HEIGHT: 86px; WIDTH: 322px" wrap=virtual>Message Body</TEXTAREA></TD></TR></TABLE>  <BR>  <BR> <BR><BR>
    <INPUT type="submit" value="Send Mail">
    <INPUT type="hidden" name="flag" value="1">
    </HTML>
    <%
    Else
    Dim anonFrom,anonTo,anonSubj,anonBody
    anonFrom = request.form("From")
    anonTo = request.form("To")
    anonSubj = request.form("Subject")
    anonBody = request.form("Body")
    Set objMail = CreateObject("CDONTS.NewMail")
    objMail.From=anonFrom
    objMail.To=anonTo
    objMail.Subject=anonSubj
    objMail.Body=anonBody
    intReturn=objMail.Send()
    %>
    <HTML>
<style>
<--
a:link { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: underlined }
a:visited { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: underlined }
a:active { color: #000080; font: 7.5pt verdana; font-weight: none; text-decoration: none }
a:hover { color: #000080; font: 7.5pt verdana; font-weight: none; text-decoration: none }
td { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: none }
body { color: #000000; font: 7.5pt verdana; font-weight: none; text-decoration: none }
input { color:#000000; font: 7.5pt verdana; font-weight: none; text-decoration: none; background: #c0c0c0; border: 1 solid #000000; }
-->
</style>
    <H1>The message sent successfully - Please Link to Secone on your site! - www.Secone.com</H1>
    <INPUT type='button' value='Back' onclick=history.back()>
    </HTML>
    <%
    End If
    %>
