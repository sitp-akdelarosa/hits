<%@LANGUAGE = VBScript%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi920.asp				_/
'_/	Function	:�R���e�i���(�ꗗ)��ʐڑ�		_/
'_/	Date		:2003/05/26				_/
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
  dim inPutStr
'�f�[�^�擾
  dim CONnum,Flag,i,j,Num
  CONnum = Request("CONnum")
  Flag   = Request("Flag")

'�R���e�i�ԍ��擾
  dim param
  If Flag=3 Then
    For Each param In Request.Form
      If Left(param, 6) = "CONnum" Then
        If param <> "CONnum" Then
          CONnum=CONnum &","& Request.Form(param)
        End If
      End If
    Next
    inPutStr="<INPUT type=hidden name='cntnrno' value='"& CONnum &"'>"
  Else
    inPutStr="<INPUT type=hidden name='blno' value='"& Request("BLnum") &"'>"
  End If
'CW-059  Session.Contents("route") = "�A���R���e�i���Ɖ�i��ƑI���j "	'CW-011
  Session.Contents("route") = "Top > �A���R���e�i���Ɖ�i��ƑI���j "	'CW-059
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�]����</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function opnewin(){
  window.focus();
  document.dmi920F.submit();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="opnewin()">
<P>�]����...���΂炭���҂����������B</P>
<FORM action="../impcntnr.asp" name="dmi920F">
<%= inPutStr %>
</FORM>
</BODY></HTML>
