<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>�\��I�����</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
/* �u�b�L���O�ԍ��̃N���b�N */
function ClickBook() {

	if (document.VPBook.VPBookNo.value == "") {
		window.alert("�u�b�L���O�ԍ�����͂��Ă��������B");
		return false;
	}
}

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sHH, sHHName, sYMD_I
	Dim conn, rsd, sql
	Dim iCnt, i, sOpeNo
	Dim iTimeCnt, TimeSlot(40), TimeName(40)

	'�w����t�擾
	sYMD_I = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD_I, 9, 2)
	sYMD = Left(sYMD_I, 8)
	sHHName = TRIM(Request.QueryString("NAME"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title32.gif" width="236" height="34"><p>

	<tr ALIGN=middle>
		<td width="360" ><font size=5><u><%=ChgYMDStr2(sYMD)%>�@<%=sHHName%></u></font></td>
	</tr>

</center>
<br>

<font face="�l�r �S�V�b�N">
   
<center>
<table>
<form  METHOD="post" NAME="REV" ACTION="SYWB01A.asp?YMD=<%=sYMD_I%>&NAME=<%=sHHName%>">
	<tr><td align=left><b>(1)���o���\��(�ܑO��)</b></td>
	<td></td>
	</tr>
	<tr><td align=right>�Ώۂb�x�E�u�o�I��</td><td align=right>
		<SELECT NAME="SELECT" width="100">
				<OPTION VALUE="KA">���łb�x
				<%	
					sql = "SELECT * FROM sTerminal WHERE Umu <> '1' And Terminal <> 'KA ' "
					sql = sql & "  Order By Terminal"
					rsd.Open sql, conn, 0, 1, 1
			
					if not rsd.eof then
						do while not rsd.EOF%>
							<OPTION VALUE=<%=rsd("Terminal")%>><%=rsd("Name")%>
							<%rsd.MoveNext
						loop
					end if
					rsd.Close
				%></select></td>
		<td><input type="submit" value="�@���s�@" id=submit2 name=submit2></td>
	</tr>

</form>
<tr></tr><tr></tr><tr></tr><tr></tr><tr></tr>
<form  METHOD="post" NAME="VPBook" ACTION="SYWB060.asp?YMD=<%=sYMD_I%>&NAME=<%=sHHName%>" onSubmit="return ClickBook()">
	<tr><td align=left><b>(2)��o���\��</b>
	<td></td>
	</tr>
	<tr><td align=right>�u�b�L���O�ԍ�����</td><td  align=right>
         <INPUT NAME="VPBookNo" SIZE="28" MAXLENGTH="16" STYLE="ime-mode:disabled">
		</td>
		<td>
			<input type="submit" value="�@���s�@" id=submit4 name=submit4>
		</td>
	</tr>
</form>
</table>
</center>

<center>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="�@���~�@" id=submit6 name=submit6></td>
	</form>
</center>

</body>
</html>
