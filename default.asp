<!--#include file="inc_browserrecon.asp"-->

<%

Dim sMode
Dim sResult

sMode = Request.QueryString("mode")

If Not (LenB(sMode) = 0 Or sMode = "simple" Or sMode = "besthitdetail" Or sMode = "list") Then
	sMode = "simple"
End If
	
sResult = (BrowserRecon(GetFullHeaders(), sMode, "C:\Inetpub\wwwroot\browserrecon\"))
Response.Write sResult

%>
