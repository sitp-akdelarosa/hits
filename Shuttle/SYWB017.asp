<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB017.inc"-->
<html>

<head>
<title>���p�񐔃��j�^</title>
</head>

<body>
<%
	Dim sYMD, sChassisID, sDispChassis1, sDispChassis2  
	Dim conn, rsd, sql
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator,sMonthStart
	Dim sNMonth, sBMonth1, sBMonth2, sBMonth3
	Dim sDisp_Date1, sDisp_Date2, sDisp_Date3,sDisp_Date4
	Dim i, sGroupName, sTrgDate 

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'���[�U���̎擾
	sql = "SELECT GroupID,GroupName FROM sMGroup" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
			rsd.Open sql, conn, 0, 1, 1
	if not rsd.EOF then
		sGroupName = rsd("GroupName")
	end if
	rsd.Close

	sGroupName = sGroupName & "�@�a"

	'�w����t�擾
	sTrgDate = TRIM(Request.QueryString("YMD"))

	'���x�J�n���̎擾
	sMonthStart= GetEnv(conn, rsd, "MonthStart")

	'�ߋ��R�����̔N���擾
	
	Call GetBefore3Month(date(), sMonthStart, sNMonth, sBMonth1, sBMonth2, sBMonth3)
		
	sDisp_Date1 = left(sNMonth,4) & "�N" & mid(sNMonth,5) & "��"
	sDisp_Date2 = left(sBMonth1,4) & "�N" & mid(sBMonth1,5) & "��"
	sDisp_Date3 = left(sBMonth2,4) & "�N" & mid(sBMonth2,5) & "��"
	sDisp_Date4 = left(sBMonth3,4) & "�N" & mid(sBMonth3,5) & "��"


%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title31.gif" width="236" height="34"><p>
<b><u><font size=3><%=sGroupName %></font></u></b><br><br>

<FORM ACTION="SYWB034.asp?TDATE=<%=sTrgDate%>" METHOD="post">
<b><font size=3>�N���I���i�ߋ��R�����j</font></b>

<SELECT NAME="SELECT1">
<OPTION VALUE="No" >�@
<OPTION VALUE=<%=sNMonth%>><%=sDisp_Date1%>
<OPTION VALUE=<%=sBMonth1%>><%=sDisp_Date2%>
<OPTION VALUE=<%=sBMonth2%>><%=sDisp_Date3%>
<OPTION VALUE=<%=sBMonth3%>><%=sDisp_Date4%>
</select>
<input type="submit" value="��    ��" id=submit4></form>
<br><br>

<form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
<input type="submit" value="��    ��"id=submit4 name=submit4>
</form>

</center>
</body>     
</html>     
