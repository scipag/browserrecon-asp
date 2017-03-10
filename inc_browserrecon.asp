<%@ Language=VBScript EnableSessionState=False %>
<%

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	browserrecon 1.0-asp
'
'	(c) 2008 by Marc Ruef
'	marc.ruef@computec.ch
'	http://www.computec.ch/projekte/browserrecon/
'
'	Released under the terms and conditions of the
'	GNU General Public License 3.0 (http://gnu.org).
'
'	Installation:
'	Extract the zip archive in a folder accessible by your
'	web browser. Include the browserrecon script with
'	the following statement:
'		<!--#include file="inc_browserrecon.asp"-->
'   There mightsome access privileges for the following
'	objects required (check the registry permissions):
'		Scripting.FileSystemObject (local file access)
'		Scripting.Dictionary       (dictionaries)
'
'	Use:
'	Use the function BrowserRecon() to do a web browser
'	fingerprinting with the included utility. The first
'	argument of the function call is the raw http headers
'	sent by the client. You might use the following
'	call to do a live fingerprinting of visiting users:
'		Response.Write BrowserRecon(GetFullHeaders());
'
'	It is also possible to get the data from another
'	source. For example a local file named header.txt:
'		Response.Write BrowserRecon(ReadFile("header.txt")));
'
'	Or the data sent via a http post form:
'		Response.Write BrowserRecon(Request.Form("header"));
'
'	Reporting:
'	You are able to change the behavior of the reports
'	sent back by BrowserRecon(). As second argument you
'	might use the following parameters:
'		- simple: Identified implementation only
'		- besthitdetail: Additional hit detail
'		- list: Unordered list of all matches
'
'	Limitations:
'	The ASP implementation of browserrecon might not provide
'	all features of the PHP version. Please check the project
'	web site for more details.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Main Function
Public Function BrowserRecon(ByRef sRawHeader, ByRef sMode, ByRef sDataBase)
	BrowserRecon = AnnounceFingerprintMatches(GenerateMatchStatistics(IdentifyGlobalFingerprint(sDataBase, sRawHeader)), sMode, CountHitPossibilities(sRawHeader))
End Function

' Header Extraction
Public Function GetFullHeaders()
	Dim sHeaders
	Dim iHeadersCount
	Dim i
	Dim sNewHeaders

	sHeaders = Split(Request.ServerVariables("ALL_RAW"), vbCrLf, 64, 0)

	iHeadersCount = UBound(sHeaders)
	For i = 0 To iHeadersCount
		sNewHeaders = sNewHeaders & sHeaders(iHeadersCount-i)

		If (i <= iHeadersCount) Then
			sNewHeaders = sNewHeaders & vbCrLf
		End If
	Next

	GetFullHeaders = sNewHeaders
End Function

Public Function GetHeaderValue(ByRef sRawHeader, ByRef sHeaderName)
	Dim sHeaders
	Dim sHeaderSmall
	Dim sHeader
	Dim sHeaderData

	sHeaders = Split(sRawHeader, vbCrLf, 64, 0)
	sHeaderSmall = Lcase(sHeaderName)

	For Each sHeader in sHeaders
		sHeaderData = Split(sHeader, ":", 2, 0)
		If (Ubound(sHeaderData) > 0) Then
			If (Lcase(sHeaderData(0)) = sHeaderSmall) Then
				GetHeaderValue = Trim(sHeaderData(1))
				Exit Function
			End If
		End if
	Next
End Function

Public Function GetHeaderOrder(ByRef sRawHeader)
	Dim sHeaders
	Dim iHeadersCount
	Dim i
	Dim sHeaderData
	Dim sHeaderOrder

	sHeaders = Split(sRawHeader, vbCrLf, 64)
	iHeadersCount = Ubound(sHeaders)

	For i = 0 To iHeadersCount
		sHeaderData = Split(sHeaders(i), ":", 2, 0)

		If (Ubound(sHeaderData) > 0) Then
			sHeaderOrder = sHeaderOrder & Trim(sHeaderData(0))

			If (LenB(sHeaders(i+1)) <> 0) Then
				sHeaderOrder = sHeaderOrder & ", "
			End If
		End if
	Next

	GetHeaderOrder = sHeaderOrder
End Function

Public Function CountHitPossibilities(ByRef sRawHeader)
	Dim iCount

	If (LenB(GetHeaderValue(sRawHeader, "User-Agent")) <> 0)		Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Accept")) <> 0)			Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Accept-Language")) <> 0)	Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Accept-Encoding")) <> 0)	Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Accept-Charset")) <> 0)	Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Keep-Alive")) <> 0)		Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Connection")) <> 0)		Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "Cache-Control")) <> 0)		Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "UA-Pixels")) <> 0)			Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "UA-Color")) <> 0)			Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "UA-OS")) <> 0)				Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "UA-CPU")) <> 0)			Then iCount = iCount + 1 End If
	If (LenB(GetHeaderValue(sRawHeader, "TE")) <> 0)				Then iCount = iCount + 1 End If
	If (LenB(GetHeaderOrder(sRawHeader)) <> 0)						Then iCount = iCount + 1 End If

	CountHitPossibilities = iCount
End Function

' Database Search
Public Function IdentifyGlobalFingerprint(ByRef sDataBase, ByRef sRawHeader)
	Dim sMatchList

	sMatchList = FindMatchInDataBase(sDataBase & "user-agent.fdb", GetHeaderValue(sRawHeader, "User-Agent"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "accept.fdb", GetHeaderValue(sRawHeader, "Accept"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "accept-language.fdb", GetHeaderValue(sRawHeader, "Accept-Language"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "accept-encoding.fdb", GetHeaderValue(sRawHeader, "Accept-Encoding"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "accept-charset.fdb", GetHeaderValue(sRawHeader, "Accept-Charset"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "keep-alive.fdb", GetHeaderValue(sRawHeader, "Keep-Alive"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "connection.fdb", GetHeaderValue(sRawHeader, "Connection"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "cache-control.fdb", GetHeaderValue(sRawHeader, "Cache-Control"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "ua-pixels.fdb", GetHeaderValue(sRawHeader, "UA-Pixels"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "ua-color.fdb", GetHeaderValue(sRawHeader, "UA-Color"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "ua-os.fdb", GetHeaderValue(sRawHeader, "UA-OS"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "ua-cpu.fdb", GetHeaderValue(sRawHeader, "UA-CPU"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "te.fdb", GetHeaderValue(sRawHeader, "TE"))
	sMatchList = sMatchList & FindMatchInDataBase(sDataBase & "header-order.fdb", GetHeaderOrder(sRawHeader))

	IdentifyGlobalFingerprint = sMatchList
End Function

Public Function FindMatchInDataBase(ByRef sDataBaseFile, ByRef sFingerprint)
	Dim sDataBase
	Dim sEntry
	Dim sEntryArray
	Dim sMatches

	sDataBase = Split(ReadFile(sDataBaseFile), vbCrLf)

	For Each sEntry	in sDataBase
		sEntryArray = Split(sEntry, ";", 2, 0)

		If (UBound(sEntryArray) > 0) Then
			If (sFingerprint = sEntryArray(1)) Then
				sMatches = sMatches & sEntryArray(0) & ";"
			End If
		End If
	Next

	FindMatchInDataBase = sMatches
End Function

Public Function ReadFile(ByRef sFileName)
	Dim sFileContents
	Dim oFS
	Dim oTextStream

	Set oFS = Server.CreateObject("Scripting.FileSystemObject")

	If oFS.FileExists(sFileName) = True Then
		Set oTextStream = oFS.OpenTextFile(sFileName, 1)
		sFileContents = oTextStream.ReadAll
		oTextStream.Close
		Set oTextStream = nothing
	End If 
   
	Set oFS = nothing

	ReadFile = sFileContents
End Function

Public Function AnnounceFingerprintMatches(ByRef sFullMatchList, ByRef sMode, ByRef iHitPossibilities)
	Dim sResultArray
	Dim sResult
	Dim sEntry
	Dim sScanBestHitName
	Dim iScanBestHitCount
	Dim sScanResultList
	Dim iScanHitAccuracy

	sResultArray = Split(sFullMatchList, vbCrLf)

	For Each sResult in sResultArray
		sEntry = Split(sResult, "=", 2, 0)

		If (Ubound(sEntry) = 1) Then
			If (iScanBestHitCount < sEntry(1)) Then
				sScanBestHitName = sEntry(0)
				iScanBestHitCount = sEntry(1)
			End If
			sScanResultList = sScanResultList & sEntry(0) & ":" & sEntry(1) & vbCrLf
		End If
	Next

	If (sMode = "list") Then
		AnnounceFingerprintMatches = sScanResultList
	Elseif (sMode = "besthitdetail") Then
		If (iHitPossibilities > 0) Then
			iScanHitAccuracy = Round(((100 / iHitPossibilities) * iScanBestHitCount), 2)
		Else
			iScanHitAccuracy = 100
		End If

		AnnounceFingerprintMatches = sScanBestHitName & " (" & iScanHitAccuracy & " % with " & iScanBestHitCOunt & " hits)"
	Else
		AnnounceFingerprintMatches = sScanBestHitName
	End If
End Function

Public Function GenerateMatchStatistics(ByRef sMatchList)
	Dim sMatchesArray
	Dim sMatches
	Dim sMatchStatistic
	Dim sMatch

	sMatchesArray = Split(sMatchList, ";")
	sMatches = RemoveDuplicatesFromArray(sMatchesArray)

	For Each sMatch In sMatches
		sMatchStatistic = sMatchStatistic & sMatch & "=" & CountIf(sMatchesArray, sMatch) & vbCrLf
	Next

	GenerateMatchStatistics = sMatchStatistic
End Function

Public Function CountIf(ByRef sInput, ByRef sSearch)
	Dim iSum
	Dim sEntry

	For Each sEntry In sInput
		If (sEntry = sSearch) Then
			iSum = iSum + 1
		End If
	Next

	CountIf = iSum
End Function

Public Function RemoveDuplicatesFromArray(ByRef sArray)
	Dim dDictionary
	Dim sItem
	Dim sTheKeys

	Set dDictionary = CreateObject("Scripting.Dictionary")
	dDictionary.removeall
	dDictionary.CompareMode = 0

	For Each sItem In sArray
		If Not dDictionary.Exists(sItem) Then dDictionary.Add sItem, sItem
	Next

	sTheKeys = dDictionary.keys
	Set dDictionary = Nothing
	RemoveDuplicatesFromArray= sTheKeys
End Function

%>
