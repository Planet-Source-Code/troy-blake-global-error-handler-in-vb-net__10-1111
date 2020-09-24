<div align="center">

## Global Error Handler in VB\.Net


</div>

### Description

I saw the earier example of a global

error handler written in C#, but

needed it written in VB for my

company. I translated the earlier

work into my version in VB. It was

suggested by a couple of people

that I provide my VB version, so

here it is. I just hope you find

it useful.

You can visit the C# version at:

http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=948&lngWId=10

It was submitted by Joel

Thoms on 2/5/2003. Thanks to all

that asked me to post the VB version.

Special thanks to Charles Richardson

for helping me track down a bug.

When you paste the code into the IDE,

most of the formatting should return.
 
### More Info
 
'Sample Use

Try

'Regular Code Here

Catch MyErr As Exception

' Catch Errors

' Send Email Only

SendHtmlError(MyErr, "yourname@yourBusiness.com")

' Show Error Only

Response.Write(GetHTMLError(MyErr))

End Try

'Code Statement

SendHtmlError(MyErr, "yourname@yourBusiness.com")

'To Show error without email

Response.Write(GetHTMLError(MyErr))

I hard-coded some of the more basic

aspects of the email. For example,

I hard-coded the from address, email

subject, etc. because I knew in

advance what I wanted them to be.

You could pass these values as part

of the call to the sub.

I just needed an easy way to email

the formatted error message to me

in the event of an error. This

works great in ASP.Net and VB.Net

applications.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Blake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-blake.md)
**Level**          |Advanced
**User Rating**    |4.6 (46 globes from 10 users)
**Compatibility**  |VB\.NET, ASP\.NET
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__10-6.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-blake-global-error-handler-in-vb-net__10-1111/archive/master.zip)

### API Declarations

```
Original submission was by Joel
Thoms, this is just the VB.Net
conversion. I didn't write this
code, I just "translated" it to
VB.
```


### Source Code

'Code Module
Imports System
Imports System.Data
Imports System.Web
Imports System.Web.Mail
Imports System.Collections.Specialized
Module ModError
 Public Sub SendHtmlError(ByVal Ex As Exception, ByVal EmailAddress As String)
 Dim Mail As New MailMessage()
 Mail.From = "ERROR_HANDLER"
 Mail.To = EmailAddress
 Mail.Subject = "Custom Intranet Error"
 Mail.Body = GetHTMLError(Ex)
 Mail.BodyFormat = MailFormat.Html
 SmtpMail.SmtpServer = "100.1.1.1"
 SmtpMail.Send(Mail)
 End Sub
 Public Function GetHTMLError(ByVal Ex As Exception) As String
 'Returns HTML an formatted error message.
 Dim Heading As String
 Dim MyHTML As String
 Dim Error_Info As New NameValueCollection()
 Heading = "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD bgcolor=""black"" COLSPAN=""2""><FONT face=""Arial"" color=""white""><B> <!--HEADER--></B></FONT></TD></TR></TABLE>"
 MyHTML = "<FONT face=""Arial"" size=""4"" color=""red"">Error - " & Ex.Message & "</FONT><BR><BR>"
 Error_Info.Add("Message", CleanHTML(Ex.Message))
 Error_Info.Add("Source", CleanHTML(Ex.Source))
 Error_Info.Add("TargetSite", CleanHTML(Ex.TargetSite.ToString()))
 Error_Info.Add("StackTrace", CleanHTML(Ex.StackTrace))
 MyHTML += Heading.Replace("<!--HEADER-->", "Error Information")
 MyHTML += CollectionToHtmlTable(Error_Info)
 '// QueryString Collection
 MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "QueryString Collection")
 MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.QueryString)
 '// Form Collection
 MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Form Collection")
 MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Form)
 '// Cookies Collection
 MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Cookies Collection")
 MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Cookies)
 '// Session Variables
 MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Session Variables")
 MyHTML += CollectionToHtmlTable(HttpContext.Current.Session)
 '// Server Variables
 MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Server Variables")
 MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.ServerVariables)
 Return MyHTML
 End Function
 Public Function CollectionToHtmlTable(ByVal Collection As NameValueCollection) As String
 Dim TD As String
 Dim MyHTML As String
 Dim i As Integer
 TD = "<TD><FONT face=""Arial"" size=""2""><!--VALUE--></FONT></TD>"
 MyHTML = "<TABLE width=""100%"">" & _
  " <TR bgcolor=""#C0C0C0"">" & _
  TD.Replace("<!--VALUE-->", " <B>Name</B>") & _
  " " & TD.Replace("<!--VALUE-->", " <B>Value</B>") & "</TR>"
 'No Body? -> N/A
 If (Collection.Count <= 0) Then
 Collection = New NameValueCollection()
 Collection.Add("N/A", "")
 Else
 'Table Body
 For i = 0 To Collection.Count - 1
 MyHTML += "<TR valign=""top"" bgcolor=""#EEEEEE"">" & _
  TD.Replace("<!--VALUE-->", Collection.Keys(i)) & " " & _
  TD.Replace("<!--VALUE-->", Collection(i)) & "</TR> "
 Next i
 End If
 'Table Footer
 Return MyHTML & "</TABLE>"
 End Function
 Private Function CollectionToHtmlTable(ByVal Collection As HttpCookieCollection) As String
 'Converts HttpCookieCollection to NameValueCollection
 Dim NVC = New NameValueCollection()
 Dim i As Integer
 Dim Value As String
 Try
 If Collection.Count > 0 Then
 For i = 0 To Collection.Count - 1
  NVC.Add(i, Collection(i).Value)
 Next i
 End If
 Value = CollectionToHtmlTable(NVC)
 Return Value
 Catch MyError As Exception
 MyError.ToString()
 End Try
 End Function
 Private Function CollectionToHtmlTable(ByVal Collection As System.Web.SessionState.HttpSessionState) As String
 'Converts HttpSessionState to NameValueCollection
 Dim NVC = New NameValueCollection()
 Dim i As Integer
 Dim Value As String
 If Collection.Count > 0 Then
 For i = 0 To Collection.Count - 1
 NVC.Add(i, Collection(i).ToString())
 Next i
 End If
 Value = CollectionToHtmlTable(NVC)
 Return Value
 End Function
 Private Function CleanHTML(ByVal HTML As String) As String
 If HTML.Length <> 0 Then
 HTML.Replace("<", "<").Replace("\r\n", "<BR>").Replace("&", "&").Replace(" ", " ")
 Else
 HTML = ""
 End If
 Return HTML
 End Function
End Module

