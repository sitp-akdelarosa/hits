<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sUserID
Dim conn, rs
Dim sContE, sContN
Dim sLastContNo
Dim nlen
Dim sSQL
Dim sErrMsg

sErrMsg = ""
sContE = ""
sContN = ""

Dim sPhoneType
sPhoneType = GetPhoneType()

sUserID = Trim(Request.QueryString("UserID"))

sErrMsg = CheckUserID(sUserID)

If sErrMsg = "" Then
	ConnectSvr conn, rs

	' [Ue[uðõµA¼OÉìµ½ReiÔðæ¾·é
	sSQL = "SELECT lUserTable.BeforeCntnrNo FROM lUserTable WHERE lUserTable.UserID='" & sUserID & "'"
	rs.Open sSQL, conn, 0, 1
	If rs.Eof Then
		rs.Close
		rs.Open "lUserTable", conn, 2, 2
		rs.AddNew
		rs("UserID") = sUserID
		rs("CompanyName") = "Unknown"
		rs.Update
		rs.Close
	Else
		If Not IsNull(rs("BeforeCntnrNo")) Then
			' ReiÔðpªÆªÉª·é
			sLastContNo = rs("BeforeCntnrNo")
			sContE = "value=""" & Left(sLastContNo, 4) & """ "
			nlen = Len(sLastContNo)
			If 4 < nlen Then
				sContN = "value=""" & Right(sLastContNo, nlen - 4) & """ "
			End If
		End If
		rs.Close
	End If
	conn.Close
End If

' LogoÍ
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
If sErrMsg="" Then
	WriteLogM oFS, sUserID, "6200", "gÑ-OC", "10",sPhoneType, sUserID & "," & "üÍàeÌ³ë:0(³µ¢)"
	WriteLogM oFS, sUserID, "6201", "gÑ-®¹ReiÔüÍ", "00",sPhoneType, ","
Else
	WriteLogM oFS, sUserID, "6200", "gÑ-OC", "10",sPhoneType, sUserID & "," & "üÍàeÌ³ë:1(ëè)" & sErrMsg
End If
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWebp^OðÒW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
<%
	If sErrMsg <> "" Then
%>
		<display title="®¹üÍ">
			<center>
			y®¹üÍz<br><br>
			<center>
			<%=sErrMsg%><br>
			<center>
			<a task="gosub" dest="index.asp">ÒÆ­°</a>
		</display>
<%
	Else
%>
		<entry name="p1" key="cont_e" format="*A" title="®¹üÍ">
			<action type="accept" task="go" dest="#p2">
			<center>
			y®¹üÍz<br>
			ºÝÃÅÔ<br>
			æªp4:
		</entry>

		<entry name="p2" key="cont_s" format="*N">
			<action type="accept" task="go" dest="mrung02.asp?UserID=<%=sUserID%>&cont_e=$cont_e&cont_s=$cont_s">
			<center>
			y®¹üÍz<br>
			ºÝÃÅÔ<br>
			ª7:
		</entry>
<%
	End If
%>
	</hdml>
<%
Else
	' EzWebÈOÌ^OðÒW
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("®¹üÍ")%>
	</head>
	<body>
	<center>
	y®¹üÍz
	<hr>
<%
	If sErrMsg <> "" Then
%>
		<br>
		<%=sErrMsg%><br>
		<br>
		<form action="index.asp" method="get">
			<input type="submit" value="ÒÆ­°">
		</form>
<%
	Else
%>
		<form action="mrung02.asp" method="get">
			ºÝÃÅÔüÍ<br>
			<table boreder="0">
				<tr><td>
					p4:
					<input type="text" name="cont_e" <%=sContE%> maxlength="4" <%=GetTextSizeMode(4, "A")%>><br>
				</td></tr>
				<tr><td>
					:
					<input type="text" name="cont_s" <%=sContN%> maxlength="8" <%=GetTextSizeMode(8, "N")%>><br>
				</td></tr>
			</table>
			<input type="hidden" name="UserID" value="<%=sUserID%>">
			<input type="submit" value="è">
		</form>
<%
	End If
%>
	<hr>
	</body>
	</html>
<%
End If
%>
