<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	�������o�V�X�e���y�C�ݗp�z	�f�[�^�ꗗ�\��
%>

<%
	' �Z�b�V�����̃`�F�b�N
	CheckLogin "sokuji.asp"

	' �C�݃R�[�h�擾
	sForwarder = Trim(Session.Contents("userid"))

	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

  ' �\���t�@�C���̎擾
	Dim strFileName
  strFileName = Session.Contents("tempfile")

  If strFileName="" Then
		Response.Redirect "sokuji-kaika-updtchk.asp"
		Response.End
	End If
	strFileName="./temp/" & strFileName

  ' �\���t�@�C����Open
  Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	Dim strData()
	LineNo=0
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)
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
<title>�������o�\���ݏ��ꗗ�i�C�݁j</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Pragma" content="no-cache">
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
          <td rowspan=2><img src="gif/sokuji1t.gif" width="506" height="73"></td>
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
<%
	strNowTime = Year(Now) & "�N" & _
		Right("0" & Month(Now), 2) & "��" & _
		Right("0" & Day(Now), 2) & "��" & _
		Right("0" & Hour(Now), 2) & "��" & _
		Right("0" & Minute(Now), 2) & "�����݂̏��"

%>
					&nbsp;&nbsp;<%=strNowTime%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td nowrap align=left> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>�i�C�ݗp�j�������o�\���ݏ��ꗗ</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>

            <br>

			<table border=0 cellpadding=0>
			  <tr>
				<td align=center colspan=2>

					<table border=0 cellpadding=0 cellspacing=2>
					  <tr>
						<td colspan=4 nowrap>
					�\���݃f�[�^���X�V����ꍇ�́A�ΏۂƂȂ�׎���N���b�N���ĉ������B <BR>
					�V�K�ɐ\���ޏꍇ�́A'�V�K����' ���N���b�N���ĉ������B<BR><BR>
						</td>
					  </tr>
					</table>

					<table border=0 cellpadding=0 cellspacing=2 width=500>
					  <tr>
						<td colspan=4 nowrap>
							�����؎����̎��{���@��
						</td>
					  </tr>
					  <tr>
						<td width=20 rowspan=5><BR></td>
						<td nowrap valign=top>
							�P�D�C�� �� �^�[�~�i���̘A��
						</td>
						<td valign=top nowrap> �F </td>
						<td valign=top>
							���O�ɑΏۂ̑D���AVoyage No.�A�R���e�iNo.���^�[�~�i���ɓd�b�ŘA������
						</td>
					  </tr>
					  <tr>
						<td nowrap valign=top>
							�Q�D�C�� �� �^�[�~�i���̐\������
						</td>
						<td valign=top nowrap> �F </td>
						<td valign=top>
							Web��ʏ�ɓ��͂Ɠ����ɓd�b�ŘA��
						</td>
					  </tr>
					  <tr>
						<td nowrap valign=top>
							�R�D�^�[�~�i�� �� �C�݂̉�
						</td>
						<td valign=top nowrap> �F </td>
						<td valign=top>
							Web��ʏ�ɓ��͂Ɠ����ɓd�b�ŘA��
						</td>
					  </tr>
					  <tr>
						<td nowrap colspan=3>
							�S�DOK�Ȃ�A�C�݌o�R�ŒS�����闤�^��Ђ�HITS�ŃV���g���ւ�\��<BR>
							�T�D�V���g���ւŃR���e�i���o
						</td>
					  </tr>
					</table>
					<BR>
				</td>
			  </tr>
			  <tr>
				<td width=30><BR></td>
				<td nowrap>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>�׎�</td>
                <td nowrap>�D��</td>
                <td nowrap>�D��</td>
                <td nowrap>BL�^�R���e�iNo.</td>
                <td nowrap>�Ή��`�^</td>
                <td nowrap>�Ή��`�^<BR>TEL</td>
                <td nowrap>�Ή�<BR>��</td>
                <td nowrap>�����m�F�\�莞��</td>
              </tr>

<%

	If LineNo>0 Then
		For i = 1 to LineNo
		    anyTmp=Split(strData(i-1),",")
%>
              <tr bgcolor="#FFFFFF"> 

				<form method=post action="sokuji-kaika-new.asp?kind=0">
					<td nowrap align=center valign=middle>
						<input type=hidden name="shipper"	 value="<%=anyTmp(9)%>">
						<input type=hidden name="shipline"	 value="<%=anyTmp(10)%>">
						<input type=hidden name="vslcode"	 value="<%=anyTmp(11)%>">

<% If Trim(anyTmp(3))<>"" Then %>
						<input type=hidden name="bl"	 value="<%=anyTmp(3)%>">
<% Else %>
						<input type=hidden name="cont"	 value="<%=anyTmp(4)%>">
<% End If %>

						<input type=hidden name="ope"		 value="<%=anyTmp(5)%>">
						<input type=hidden name="opetel"	 value="<%=anyTmp(6)%>">
						<input type=hidden name="reject"	 value="<%=anyTmp(7)%>">
						<input type=hidden name="recschtime" value="<%=anyTmp(8)%>">
						<input type=hidden name="lineno"	 value="<%=i%>">

<% If Trim(anyTmp(0))<>"" Then %>
						<a href="JavaScript:formSend(<%=i%>)"><%=anyTmp(0)%></a>
<% Else %>
						<a href="JavaScript:formSend(<%=i%>)"><%=anyTmp(9)%></a>
<% End If %>

					</td>
				</form>

<% If Trim(anyTmp(1))<>"" Then %>
				<td nowrap align=center valign=middle><%=anyTmp(1)%></td>
<% Else %>
				<td nowrap align=center valign=middle><%=anyTmp(10)%></td>
<% End If %>

<% If Trim(anyTmp(2))<>"" Then %>
				<td nowrap align=center valign=middle><%=anyTmp(2)%></td>
<% Else %>
				<td nowrap align=center valign=middle><%=anyTmp(11)%></td>
<% End If %>

<% If Trim(anyTmp(3))<>"" Then %>
				<td nowrap align=center valign=middle><%=anyTmp(3)%></td>
<% Else %>
				<td nowrap align=center valign=middle><%=anyTmp(4)%></td>
<% End If %>

<%
			For j = 0 to 8
				If anyTmp(j)=""Then
					anyTmp(j) = "<BR>"
				End If
			Next

			If Not anyTmp(8)="<BR>" Then
				anyTmp(8) = Right(anyTmp(8),11)
			End If
%>
				<td nowrap align=center valign=middle><%=anyTmp(5)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(6)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(7)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(8)%></td>
			  </tr>
<%
		Next
	End If
%>
			  <tr>
				<td nowrap align=center valign=middle>
					<a href="sokuji-kaika-new.asp?kind=1">�V�K����</a>
				</td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
				<td nowrap align=center valign=middle><BR></td>
			  </tr>
			</table>

			<form method=get action="sokuji-kaika-updtchk.asp">
				<input type=submit value="�\���f�[�^�̍X�V">
			</form>

				</td>
			  </tr>
			</table>

		  </td>
		</tr>
	  </table>

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
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log�쐬
    WriteLog fs, "7001", "�������o�V�X�e��-�C�ݗp���ꗗ", "00", ","
%>
