<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo900.asp				_/
'_/	Function	:�����o���͏��擾			_/
'_/	Date		:2003/12/17				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
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

'�f�[�^�擾
  dim CONnum,Flag,BLnum,SakuNo
  dim inPutStr,strNums
  CONnum = Request("CONnum")
  Flag   = Request("flag")
  SakuNo = Request("SakuNo")

'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  Select Case Flag
    Case "1"		'�w��L��
      inPutStr="<INPUT type=hidden name='cntnrno' value='"& CONnum &"'>"
	Case "2"		'�w��Ȃ�
      StrSQL = "SELECT ITC.BLNo FROM hITCommonInfo AS ITC " &_
               "WHERE ITC.ContNo='"& CONnum &"' AND ITC.WkNo='"& SakuNo &"' AND ITC.Process='R' AND ITC.WkType='1'"
      ObjRS.Open StrSQL, ObjConn
	  inPutStr="<INPUT type=hidden name='blno' value='"& Trim(ObjRS("BLNo")) &"'>"
      ObjRS.close
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB�ؒf
        jampErrerP "1","b101","99","�����o:�ڍחp�f�[�^�擾","102","SQL:<BR>"&strSQL
      end if
	Case "3"		'�ꗗ����I��
		strNums=CONnum
	   '�ΏۃR���e�i�ԍ��ꗗ�擾
      StrSQL = "SELECT ITF.ContNo FROM hITCommonInfo AS ITC " &_
               "LEFT JOIN hITFullOutSelect AS ITF ON ITC.WkContrlNo = ITF.WkContrlNo " &_
               "WHERE ITC.ContNo='"& CONnum &"' AND ITC.WkNo='"& SakuNo &"' AND ITC.Process='R' AND ITC.WkType='1'"
	    ObjRS.Open StrSQL, ObjConn
	    Do Until ObjRS.EOF
	      If CONnum <> Trim(ObjRS("ContNo")) Then 
	        strNums = strNums & "," & Trim(ObjRS("ContNo"))
	      End If
	      ObjRS.MoveNext
	    Loop
	    ObjRS.close
	    if err <> 0 then
	      DisConnDBH ObjConn, ObjRS	'DB�ؒf
	      jampErrerP "1","b101","99","�����o:�ڍחp�f�[�^�擾","102","SQL:<BR>"&strSQL
	    end if
        inPutStr="<INPUT type=hidden name='cntnrno' value='"& strNums &"'>"
	Case "4"		'BL
	  inPutStr="<INPUT type=hidden name='blno' value='"& CONnum &"'>"
  End Select

  if Flag=1 Then
	Session.Contents("route") = "�A���R���e�i���Ɖ�i��ƑI���j "
  Else
	Session.Contents("route") = "Top > �A���R���e�i���Ɖ�i��ƑI���j "
  End If
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
  document.dmi900F.submit();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="opnewin()">
<P>�]����...���΂炭���҂����������B</P>
<FORM action="../impcntnr.asp" name="dmi900F">
<%= inPutStr %>
</FORM>
</BODY></HTML>

