<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi320.asp				_/
'_/	Function	:�폜����				_/
'_/	Date		:2003/05/29				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b402", "���������O������","15",""

'�T�[�o���t�̎擾
  dim DayTime, YY,Yotei
  getDayTime DayTime

'���[�U�f�[�^����
  dim USER
  USER   = UCase(Session.Contents("userid"))

'�f�[�^���擾
  dim WkCNo,SakuNo
  WkCNo = Request("WkCNo")
  SakuNo = Request("SakuNo")

'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
'CW-010	ADD START ��������������
  dim ret, ErrerM
  ret=true
'20030912�@Del START ������������������������������
 '�����`�F�b�N
'  StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
'         "Where WkContrlNo="& WkCNo &" AND Process='R' AND WkType='3'"
'  ObjRS.Open StrSQL, ObjConn
'  if err <> 0 then
'    ObjRS.Close
'    Set ObjRS = Nothing
'    jampErrerPDB ObjConn,"1","b401","15","�������F�f�[�^�폜","101","SQL:<BR>"&StrSQL
'  end if
'  If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
'    ret=false
'    ErrerM="�w��̍�Ƃ͉�ʑ��쒆�ɍ�Ƃ������������߁A�폜�����̓L�����Z������܂����B"
'  End If
'  ObjRS.close
'20030912�@Del END ����������������������������

  If ret Then
'CW-010	End ADD ��������������
    StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
           "UpdtTmnl='"& USER &"', UpdtUserCode='"& USER &"', Status='0', Process='D' " &_
           "Where WkContrlNo="& WkCNo &" AND Process='R' AND WkType='3'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b401","15","�������F�f�[�^�폜","104","SQL:<BR>"&StrSQL
    end if
    StrSQL = "UPDATE hITWkNo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
           "UpdtTmnl='"& USER &"', Status='4' " &_
           "Where WkNo='"& SakuNo &"'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b401","15","�������F�f�[�^�폜","104","SQL:<BR>"&StrSQL
    end if
'    StrSQL = "DELETE FROM hITReference Where WkContrlNo="& WkCNo 
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      jampErrerPDB ObjConn,"1","b401","15","�������F�f�[�^�폜","105","SQL:<BR>"&StrSQL
'    end if
  End If		'CW-010
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�폜����</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�폜����--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR align=center><TD>
<% If ret Then %>
    �폜�������ł��B<BR>���΂炭���҂����������B<P>��ʂ͎����I�ɕ����܂��B
    <SCRIPT language=JavaScript>
     try{
        window.opener.parent.DList.location.href="./dmo310L.asp"
     }catch(e){
     }
      window.close();
    </SCRIPT>
<% Else %>
    <DIV class=alert><%=ErrerM%></DIV><BR>
    <INPUT type=button value="����" onClick="window.close()">
<% End If%>
  </TD></TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
