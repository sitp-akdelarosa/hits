<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	�������o�V�X�e���y�`�^�p�z	�f�[�^�ꗗ�\��

%>

<%
	' �Z�b�V�����̃`�F�b�N
	CheckLogin "sokuji.asp"

	' �`�^�R�[�h�擾
	sOpe = Trim(Session.Contents("userid"))

	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	Dim strFileName
	strFileName = Session.Contents("tempfile")

  If strFileName="" Then
		Response.Redirect "sokuji-koun-updtchk.asp"
		Response.End
	End If
	strFileName="./temp/" & strFileName

  ' �\���t�@�C����Open
  Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' �e���|�����t�@�C���̓ǂݍ���
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
	Session.Contents("ChkCount")=LineNo

%>
<html>
<head>
<title>�������o�\���ݏ��ꗗ�i�`�^�j</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Pragma" content="no-cache">
<SCRIPT Language="JavaScript">
	function checkFormValue(){
<%
	If LineNo>0 Then
		For i=1 to LineNo
			If LineNo=1 Then
				Response.Write "if(document.koun.chk" & i & ".checked==false)"
			ElseIf i=1 Then
				Response.Write "if((document.koun.chk" & i & ".checked==false)"
			ElseIf i=LineNo Then
				Response.Write "&&(document.koun.chk" & i & ".checked==false))"
			Else
				Response.Write "&&(document.koun.chk" & i & ".checked==false)"
			End If
		Next
%>
		{ return showAlert("�`�F�b�N",true); }
		return true;
<%
	End If
%>
	}
	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "�������͂ł��B");
		} else {
			window.alert(strAlert + "���s���ł��B");
		}
		return false;
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
          <td rowspan=2><img src="gif/sokuji2t.gif" width="506" height="73"></td>
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
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>�i�`�^�p�j�������o�\���ݏ��ꗗ</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>
			<center>

            <br>
			<table border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td nowrap align=left >
			�Ή��\�ȏꍇ�͖ړI�̃f�[�^�i�E�̎l�p�̘g���j�Ƀ`�F�b�N���ė\�莞�����͂������ĉ������B<BR>
			�Ή��s�̏ꍇ�͖ړI�̃f�[�^�i�E�̎l�p�̘g���j�Ƀ`�F�b�N���đΉ��s�������ĉ������B<BR><BR>
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

			<table border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td align=center nowrap>

					<table border=0 cellpadding=0 cellspacing=0>
					<tr>
					<td nowrap align=right>

				    <form method=post action="sokuji-koun-new.asp" name="koun">

		            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
		              <tr align="center" bgcolor="#FFCC33"> 
		                <td nowrap>�C��</td>
		                <td nowrap>�D��</td>
		                <td nowrap>�D��</td>
		                <td nowrap>BL�^�R���e�iNo.</td>
		                <td nowrap>�Ή�<BR>��</td>
		                <td nowrap>�����m�F<BR>�\�莞��</td>
		                <td nowrap><BR></td>
		              </tr>

<%
	For i = 1 to LineNo
	    anyTmp=Split(strData(i-1),",")

		For j = 5 to 6
			If anyTmp(j)="" Then
				anyTmp(j) = "<BR>"
			End If
		Next

		If Not anyTmp(6)="<BR>" Then
			anyTmp(6) = Right(anyTmp(6),11)
		End If
%>
		              <tr bgcolor="#FFFFFF"> 

<% If Trim(anyTmp(0))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(0)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(7)%></td>
<% End If %>
<% If Trim(anyTmp(1))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(1)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(8)%></td>
<% End If %>
<% If Trim(anyTmp(2))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(2)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(9)%></td>
<% End If %>
<% If Trim(anyTmp(3))<>"" Then %>
						<td nowrap align=center valign=middle><%=anyTmp(3)%></td>
<% Else %>
						<td nowrap align=center valign=middle><%=anyTmp(4)%></td>
<% End If %>
						<td nowrap align=center valign=middle><%=anyTmp(5)%></td>
						<td nowrap align=center valign=middle><%=anyTmp(6)%></td>
						<td nowrap align=center valign=middle>
						  <input type=checkbox name=chk<%=i%>>
						</td>

<%
		If anyTmp(6)="<BR>" Then anyTmp(6)=""
		If anyTmp(7)="<BR>" Then anyTmp(7)=""
		If anyTmp(8)="<BR>" Then anyTmp(8)=""
		If anyTmp(8)="<BR>" Then anyTmp(9)=""
		If anyTmp(3)="<BR>" Then anyTmp(3)=""
		If anyTmp(4)="<BR>" Then anyTmp(4)=""
		If anyTmp(5)="<BR>" Then anyTmp(5)=""
%>
						<input type=hidden name=shipper<%=i%> value=<%=anyTmp(7)%>>
						<input type=hidden name=shipline<%=i%> value=<%=anyTmp(8)%>>
						<input type=hidden name=vslcode<%=i%> value=<%=anyTmp(9)%>>
						<input type=hidden name=forwarder<%=i%> value=<%=anyTmp(10)%>>
<% If Trim(anyTmp(3))<>"" Then %>
						<input type=hidden name=bl<%=i%> value=<%=anyTmp(3)%>>
<% Else %>
						<input type=hidden name=cont<%=i%> value=<%=anyTmp(4)%>>
<% End If %>
						<input type=hidden name=reject<%=i%> value=<%=anyTmp(5)%>>
						<input type=hidden name=recschtime<%=i%> value=<%=anyTmp(6)%>>

					  </tr>
<%
	Next
%>
					</table>
					<BR>
					<div align=left>
					<input type=button value="�\���f�[�^�̍X�V" onclick="window.location.href='sokuji-koun-updtchk.asp'">
					</div>
					<input type=submit name=timeset value="�\�莞������" onClick="return checkFormValue()">
					<input type=submit name=corrfail value=" �� �� �s �� " onClick="return checkFormValue()">

					</td>
					</tr>

					</form>

					<tr><td align="left">
					</td></tr>

					</table>

				</td>
			  </tr>
			</table>

		  </center>
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
    WriteLog fs, "7003", "�������o�V�X�e��-�`�^�p���ꗗ", "00", ","
%>
