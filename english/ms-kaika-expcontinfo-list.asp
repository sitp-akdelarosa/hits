<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	�y�R���e�i�����́z	�f�[�^�ꗗ�\��
	'
%>

<%
	' �����ꗗ�\���ő�l
	Dim sUser,sUserNo
    sUser   = UCase(Trim(Request.form("user")))
    sUserNo = UCase(Trim(Request.form("userno")))

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "nyuryoku-kaika.asp"             '���j���[��ʂ�
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �ڍו\���s�̃f�[�^�̎擾
    Dim strData()
    LineNo=0
    Do While Not ti.AtEndOfStream
        strTemp=ti.ReadLine
        ReDim Preserve strData(LineNo)
        strData(LineNo) = strTemp
        LineNo=LineNo+1
    Loop
    ti.Close
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function formSend(formname){
		window.document.forms[formname].submit();
	}

<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------�������烍�O�C�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika5t.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=GetUpdateTime(fs)%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td> 

            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�X�V�Ώۈꗗ</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>

<table border=0 cellpadding=0 cellspacing=0><tr><td>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŃR���e�i�����͕ύX</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j���h�@�Ɋւ��댯���̗L��</font></td>
          </tr>
        </table>

</td></tr><tr><td>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>�D��</td>
                <td nowrap>Voyage No.</td>
                <td nowrap>�׎�R�[�h</td>
                <td nowrap>�׎�Ǘ��ԍ�</td>
                <td nowrap>Booking No.</td>
                <td nowrap>�R���e�iNo.<font size="-1"><sup>(��1)</sup></font></td>
                <td nowrap>�V�[��No.</td>
                <td nowrap>�ݕ�<BR>�d��(t)</td>
                <td nowrap>��<BR>�d��(t)</td>
                <td nowrap>���[�t�@�[</td>
                <td nowrap>�댯��<BR><font size="-1"><sup>(��2)</sup></font></td>
              </tr>

<%

	For i = 1 to LineNo
	    '�g�����U�N�V�����t�@�C���쐬
	    anyTmp=Split(strData(i-1),",")
%>
              <tr bgcolor="#FFFFFF"> 
				<td nowrap align=center valign=middle><%=anyTmp(0)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(1)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(2)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(3)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(4)%></td>

<% If anyTmp(5)="" Then %>
			<form method=post action="ms-kaika-expcontinfo-new.asp?kind=1">
<% Else %>
			<form method=post action="ms-kaika-expcontinfo-new.asp?kind=0">
<% End If %>
				<td nowrap align=center valign=middle>
					<input type=hidden name="vslcode"	 value="<%=anyTmp(0)%>">
					<input type=hidden name="voyctrl"	 value="<%=anyTmp(1)%>">
					<input type=hidden name="user"		 value="<%=anyTmp(2)%>">
					<input type=hidden name="userno"	 value="<%=anyTmp(3)%>">
					<input type=hidden name="booking"	 value="<%=anyTmp(4)%>">
					<input type=hidden name="cont"		 value="<%=anyTmp(5)%>">
					<input type=hidden name="seal"		 value="<%=anyTmp(6)%>">
					<input type=hidden name="cargow"	 value="<%=anyTmp(7)%>">
					<input type=hidden name="contw"		 value="<%=anyTmp(8)%>">
					<input type=hidden name="rifer"		 value="<%=anyTmp(9)%>">
					<input type=hidden name="danger"	 value="<%=anyTmp(10)%>">
					<input type=hidden name="lineno"	 value="<%=i%>">

<% If anyTmp(5)="" Then %>
					<a href="JavaScript:formSend(<%=i%>)">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a>
<% Else %>
					<a href="JavaScript:formSend(<%=i%>)"><%=anyTmp(5)%></a>
<% End If %>
				</td>
			</form>

<%
		For j = 0 to 10
			If anyTmp(j)="" Then
				anyTmp(j) = "<BR>"
			End If
		Next
%>
				<td nowrap align=center valign=middle><%=anyTmp(6)%></td>
				<td nowrap align=right  valign=middle><%=anyTmp(7)%></td>
				<td nowrap align=right  valign=middle><%=anyTmp(8)%></td>

<% If anyTmp(9)="R" Then %>
				<td nowrap align=center valign=middle>��</td>
<% Else %>
				<td nowrap align=center valign=middle><BR></td>
<% End If %>

<% If anyTmp(10)="H" Then %>
				<td nowrap align=center valign=middle>��</td>
<% Else %>
				<td nowrap align=center valign=middle><BR></td>
<% End If %>

			  </tr>
<%
	Next
%>
			</table>

</td></tr></table>
		  </td>
		</tr>
	  </table>

    <br>
		</td>
	</tr>
	<tr>
		<td valign="bottom"> 
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------���O�C����ʏI���--------------------------->
<%
    DispMenuBarBack "ms-kaika-expcontinfo.asp"
%>
</body>
</html>

<%
	' Log�쐬
    WriteLog fs, "4108","�C�ݓ��͗A�o�R���e�i���-�X�V�Ώۈꗗ","00", ","
%>
