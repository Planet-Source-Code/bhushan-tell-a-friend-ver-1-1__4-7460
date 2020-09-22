<div align="center">

## Tell a friend ver 1\.1


</div>

### Description

Tell a friend
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bhushan\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bhushan.md)
**Level**          |Advanced
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bhushan-tell-a-friend-ver-1-1__4-7460/archive/master.zip)





### Source Code

```
<a href="mail.asp?URL=www.http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=4&grpCategories=&txtMaxNumberOfEntriesPerPage=10&optSort=Alphabetical&chkThoroughSearch=&blnTopCode=False&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=1&intLastRecordOnPage=10&intMaxNumberOfEntriesPerPage=10&intLastRecordInRecordset=14&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=bhushan+paranjpe&cmdGoToPage=2&lngMaxNumberOfEntriesPerPage=10">Tell a Friend</a>
So once he clicks on that hyperlink the URL is carried to mail.asp and let us see what mail.asp will do for us.
<html>
<head>
<title>Tell a Friend</title>
</head>
<body>
<%
URL = Request.QueryString("URL")
If Len(URL) = 0 Then URL = "http://www.planet-source-code.com" ' The default URL
name=request.form("Sendersname")
from=request.form("SendersEmail")
message=Request.Form("Message")
If Len(Request.Form("SendersEmail")) > 0 Then ' Time to send the emails
Dim objMail,FriendEmail,I
sBody = "This Page at " & URL & " has been recommended by " & name & " at " & Request.Form("SendersEmail") & vbCrLf & " <-- Message For You--> " & vbcrlf & message
I=0
' Loop until no more email addresses are given.
Do While True
FriendEmail = Request.Form("FriendEmail" & I)
If Len(FriendEmail) = 0 Then
Exit Do
Else
Set objMail = CreateObject("CDONTS.NewMail")
objMail.From = name
objMail.Subject = "Recommended Page"
objMail.Importance=1
objMail.Body = sBody
objMail.To = FriendEmail
objMail.Send()
End If
I=I+1
Loop
Set objMail = Nothing%>
<%
Response.write "<center> <font color=#FFFFFF><H1>Thank you for Recommending us to Your Friends.</H1>"%>
<%
Response.write "<a href=" & URL & " style='color: #ffffff'>Click here to return to " & URL & "</a></font></center>"
Else
%>
<form method="POST" action="mail.asp?URL=<%= URL %>">
<div align="center">
<center>
<table border="1" cellpadding="0" cellspacing="0">
<tr>
<td><b><font color="#FFFFFF">Recommended URL:</font></b> &nbsp;</td>
<td><p><font color="#FFFFFF"> <%= URL %></font></p>
</td>
</tr>
<tr>
<td><b><font color="#FFFFFF">Your Name:</font></b> </td>
<td><input type="text" name="Sendersname" size="25"></td>
</tr>
<tr>
<td><b><font color="#FFFFFF">Your Email:</font></b> </td>
<td><input type="text" name="SendersEmail" size="25"></td>
</tr>
<tr>
<td><b><font color="#FFFFFF">Your friends emails.</font></b></td>
<td>&nbsp;</td>
</tr>
<tr>
<td><font color="#FFFFFF">1.</font> </td>
<td> <input type="text" name="FriendEmail0" size="29"></td>
</tr>
<tr>
<td><font color="#FFFFFF">2.</font></td>
<td><input type="text" name="FriendEmail1" size="29"></td>
</tr>
<tr>
<td><font color="#FFFFFF">2.</font></td>
<td><input type="text" name="FriendEmail2" size="29"></td>
</tr>
<tr>
<td><font color="#FFFFFF">2.</font></td>
<td><input type="text" name="FriendEmail3" size="29"></td>
</tr>
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td><b><font color="#FFFFFF">Message</font></b></td>
<td><textarea rows="6" name="Message" cols="41"></textarea></td>
</tr>
<tr>
<td colspan="2">
<p align="center"><br>
<input type="submit" value="Tell a Friend"><br>
<br>
</p>
</td>
</tr>
</table>
</form>
<% End If %>
</td>
</tr>
</table>
</body>
</html>
```

