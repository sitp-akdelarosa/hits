<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:errer.asp				_/
'_/	Function	:�G���[���				_/
'_/	Date		:2003/06/18				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'�G���[���擾
  dim ObjConn, ObjRS,WinFlag,dispId,wkID,wkName,errerCd,etc
  WinFlag= Session.Contents("WinFlag")
  dispId = Session.Contents("dispId")
  wkID   =  Session.Contents("wkID")
  wkName =  Session.Contents("wkName")
  errerCd=  Session.Contents("errerCd")
  etc    =  Session.Contents("etc")
'�Z�b�V�����N���A
  Session.Contents.Remove("WinFlag")
  Session.Contents.Remove("dispId")
  Session.Contents.Remove("wkID")
  Session.Contents.Remove("wkName")
  Session.Contents.Remove("errerCd")
  Session.Contents.Remove("etc")

'�G���[���b�Z�[�W�擾
  dim ErrerM1,ErrerM2
  dim ObjFSO,ObjTS,tmpStr,tmp
  Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
  Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./INI/PREDEFERRER.INI"),1,false)
  '--- �t�@�C���f�[�^�̓Ǎ��� ---
  Do Until ObjTS.AtEndofStream
    tmpStr = ObjTS.ReadLine
    If Left(tmpStr,3) = errerCd Then
      tmp=Split(tmpStr,":",3,1)
      ErrerM1 = tmp(1)
      ErrerM2 = tmp(2)
      Exit Do
    End If
  Loop
  ObjTS.Close
  Set ObjTS = Nothing
  Set ObjFSO = Nothing

'�{�^���\������
  dim Button
  If WinFlag = 0 Then
    Button="'���O�C����ʂɖ߂�' onClick='submit()'"
  ElseIf WinFlag = 1 Then
    Button="'����' onClick='window.close()'"
  Else
    Button="'�߂�' onClick='window.history.back()'"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�G���[</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�G���[���--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="90%" align=center>
  <TR><TD colspan=2><DIV class=alert>�G���[</DIV></TD></TR>
  <TR><TD>�G���[���ID�F���ID</TD><TD>�F<%=dispId%>�F<%=wkId%></TD></TR>
  <TR><TD>��Ɩ�</TD><TD>�F<%=wkName%></TD></TR>
  <TR><TD>�G���[�R�[�h</TD><TD>�F<%=errerCd%></TD></TR>
  <TR><TD>���b�Z�[�W</TD><TD>�F<%=ErrerM1%><BR></TD></TR>
  <TR><TD>�Ώ�</TD><TD>�F<%=ErrerM2%><BR></TD></TR>
  <TR><TD colspan=2><%=etc%></TD></TR>
  <TR><TD colspan=2 align=center>
        <FORM action="../Userchk.asp" target="_top">
          <INPUT type=hidden name="link" value="predef/dmi000F.asp">
          <INPUT type=button value=<%=Button%>>
        </FORM>
      </TD></TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
