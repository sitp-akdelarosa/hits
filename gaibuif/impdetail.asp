<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "IMPORT", "impentry.asp"

    ' �w������̎擾
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

    ' �\�����[�h�̎擾
    Dim bDispMode          ' true=�R���e�i���� / false=BL����
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
    Else
        bDispMode = false
    End If

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "impentry.asp"             '�A���R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �ڍו\���s�̃f�[�^�̎擾
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLineNo=LineNo Then
           Exit Do
        End If
    Loop
    bSingle = false                    '�P�ƌ������ʃt���O
    If iLineNo=1 And LineNo=1 Then
        '�P�ƌ������ʂ��ǂ����`�F�b�N����
        if ti.AtEndOfStream Then
            bSingle = true
        End If
    End If
    ti.Close

' 2010/08/19 del-s ���O�o�͓��e�ύX
'    ' �A���R���e�i�Ɖ�ڍ�
'	WriteLog fs, "4006","�d���n���Ɖ�", "00", anyTmp(1) & ","
' 2010/08/19 del-e ���O�o�͓��e�ύX

    Session.Contents("dispcntnr")=anyTmp(1)     ' �\���R���e�iNo.���L��

' 2009/05/09 add-s �`�����ǉ�
    Dim sPortCode
    sPortCode = Session.Contents("usercodeex")
' 2009/05/09 add-e �`�����ǉ�

' 2010/08/19 del-s ���O�o�͓��e�ύX
    ' �A���R���e�i�Ɖ�ڍ�
    If sPortCode = "HUANG" Then
        WriteLog fs, "4006","�d���n���Ɖ�(����)","10", anyTmp(1) & ","
    ElseIf sPortCode = "NANSH" Then
        WriteLog fs, "4006","�d���n���Ɖ�(�썹)","20", anyTmp(1) & ","
    Else
        WriteLog fs, "4006","�d���n���Ɖ�","00", anyTmp(1) & ","
    End If
' 2010/08/19 del-e ���O�o�͓��e�ύX
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="./index.css" rel="stylesheet" type="text/css">
<SCRIPT language="javascript" type="text/javascript" src="./index.js"></SCRIPT>
<SCRIPT Language="JavaScript">
<!--
function FancBack()
{
        window.history.back();
}

function Submit(formName){
    document.forms[formName].submit();
}
// -->
<%
    DispMenuJava
%>
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function winOpen(winName,url,W,H){
  var WinD11=window.open(url,winName,'scrollbars=yes,resizable=yes,width='+W+',height='+H+'');
  WinD11.focus();
  WinD11.document.close();
}
</Script>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#0000ff" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000ff">
<!-------------��������ڍ׉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/impdetailt_<%=sPortCode%>.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.17
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.17
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.07 -->

<!-- mod by nics 2009.02.09 -->
<!--		<table width=95% cellpadding=3>-->
		<table width=95% cellpadding=0>
<!-- end of mod by nics 2009.02.09 -->
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
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' BL No
    If Not bDispMode Then
        Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>BL No</b></font></td>"
        Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
    End If
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�R���e�iNo</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �R���e�iNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--			<BR>-->
			<font size="-2"><BR></font>
<!-- end of mod by nics 2009.02.09 -->
<!---------------��{���------------------------------------------- commented by nics 2009.02.09 -->
<!---------------��{���--------------------------------------------->
<!-- commented by nics 2009.02.09
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>��{���</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.09 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<!-- add by nics 2009.02.09 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;��{���&nbsp;</td>
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td valign="top" nowrap>����</td>
end of comment by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">�T�C�Y</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap bgcolor="#FFCC33">����<font size="-1"><sup>(��4)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">����<font size="-1"><sup>(��1)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">���[�t�@�[</td>
                <td nowrap bgcolor="#FFCC33">���d��(t)</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td valign="top" nowrap>�댯�i<font size="-1"><sup>(��5)</sup></font></td>-->
                <td valign="top" nowrap>�댯�i<font size="-1"><sup>(��2)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap bgcolor="#FFCC33">���o�^�[�~�i��</td>-->
                <td nowrap bgcolor="#FFCC33">���o�^�[�~�i��(���u�ꏊ�R�[�h)</td>
<!-- end of mod by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">�{�D�S��<br>�I�y���[�^</td>
<!-- end of add by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">�X�g�b�N���[�h���p</td>
                <td nowrap bgcolor="#FFCC33">�ԋp�ꏊ</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFFFCC" nowrap>���</td>
end of comment by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' �T�C�Y
    If anyTmp(23)<>"" Then
        Response.Write anyTmp(23)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ����
    If anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ���[�t�@�[
    If anyTmp(25)="R" Then
        Response.Write "��"
    ElseIf anyTmp(25)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ���d��
    If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
        dWeight=anyTmp(26) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �댯��
    If anyTmp(27)="H" Then
        Response.Write "��"
    ElseIf anyTmp(27)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td align="center" nowrap>
<% ' ���o�^�[�~�i��
    If anyTmp(5)<>"" Then
        Response.Write anyTmp(5)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' ���o�^�[�~�i��(���u�ꏊ�R�[�h)
    strDisp = "<br>"
    If anyTmp(5) <> "" Then
        strDisp = anyTmp(5)
        If anyTmp(43) <> "" Then
            strDisp = strDisp & "(" & anyTmp(43) & ")"
        End If
    End If
    Response.Write strDisp
%>
                </td>
                <td align="center" nowrap>
<% ' �S���I�y���[�^
    If anyTmp(45)<>"" Then
        Response.Write anyTmp(45)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- end of add by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h���p $�ǉ�
    If anyTmp(35)>="1" And anyTmp(35)<="4" Then
        Response.Write "��"
    Else
        Response.Write "�~"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �ԋp�ꏊ
'' 2009/07/09 mod-s '-'�Œ��
'    If anyTmp(10)<>"" Then
'        Response.Write anyTmp(10)
'    Else
'        Response.Write "<br>"
'    End If
    Response.Write "�|"
'' 2009/07/09 mod-e '-'�Œ��
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <table border="0" cellspacing="1" cellpadding="3">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.09 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap><font color="#000000" size="-1">�i��4�j96=HC</td>-->
                <td nowrap><font color="#000000" size="-1">�i��1�j96=HC</td>
<!-- end of mod by nics 2009.02.09 -->
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap><font color="#000000" size="-1">�i��5�j���h�@�Ɋւ��댯�i�̗L��</td>-->
                <td nowrap><font color="#000000" size="-1">�i��2�j���h�@�Ɋւ��댯�i�̗L��</td>
<!-- end of mod by nics 2009.02.09 -->
              </tr>
            </table>
<!-- commented by nics 2009.02.09
            <br>
end of comment by nics 2009.02.09 -->
<!---------------�{�D���------------------------------------------- commented by nics 2009.02.09 -->
<!---------------�{�D���--------------------------------------------->
<!-- commented by nics 2009.02.09
            <table>
              <tr> 
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�{�D���&nbsp;&nbsp;</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.09 -->
            <table border=1 cellpadding="3" cellspacing="1">
<!-- mod by nics 2009.02.09 -->
<!--              <tr> -->
              <tr align="center" bgcolor="#FFCC33"> 
<!-- end of mod by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td rowspan="2" nowrap bgcolor="#6495ED">&nbsp;�{�D���&nbsp;</td>
<!-- end of add by nics 2009.02.09 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">�D��</font></td>
<!-- add by nics 2009.02.09 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">�D��</font></td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFCC33" nowrap>�d�o�`</td>
                <td bgcolor="#FFCC33" nowrap>�O�`</td>
              </tr>
              <tr align="center"> 
<!-- end of add by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' �D��
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap><font color="#000000">�D��</font></td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' �D��
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' ���q
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap>�d�o�`</td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' �d�o�`
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap>�O�`</td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' �O�`
'' 2009/07/09 mod-s '-'�Œ��
'    If anyTmp(38)<>"" Then
'        Response.Write anyTmp(38)
'    Else
'        Response.Write "<br>"
'    End If
    Response.Write "�|"
'' 2009/07/09 mod-e '-'�Œ��
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <br>-->
            <font size="-1"><br></font>
<!-- end of mod by nics 2009.02.09 -->
<!---------------�ʒu���------------------------------------------- commented by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�ʒu���@</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.09 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.02.09 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;�ʒu���&nbsp;</td>
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td nowrap align="center" bgcolor="#FFCC33">�ꏊ</td>
end of comment by nics 2009.02.09 -->
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap bgcolor="#FFCC33">�d�o�`<font size="-1"><sup>(��1)</sup></font></td>-->
<!-- del-s by mes 2009/05/12 -->
<!--                <td nowrap bgcolor="#FFCC33">�d�o�`<font size="-1"><sup>(��3)</sup></font></td>-->
<!-- del-e by mes 2009/05/12 -->
<!-- end of mod by nics 2009.02.09 -->
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap bgcolor="#FFCC33">�O�`<font size="-1"><sup>(��1)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">�O�`<font size="-1"><sup>(��3)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
                <td colspan="4" nowrap bgcolor="#FFCC33">�^�[�~�i��</td>
                <td nowrap bgcolor="#FFCC33">�X�g�b�N���[�h</td>
                <td colspan="3" nowrap bgcolor="#FFCC33">����A��</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.02.09
                <td nowrap rowspan="2" bgcolor="#FFFFCC">�H��</td>
end of comment by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
<!-- del-s by mes 2009/05/12 -->
<!--
                <td rowspan="4" align="center"><table border="0" cellspacing="5">
                    <tr>
                      <td nowrap><a href="javascript:Submit('Form1')" class="splink" onClick="javascript:winOpen('win1','./cct/index.html',560,500)">&nbsp;�Ԙp&nbsp;</a></td>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                    <tr>
                      <td><a href="javascript:Submit('queryForm')" class="splink" onClick="javascript:winOpen('win1','./sct/index.html',560,500)">&nbsp;�֌�&nbsp;</a></td>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                    <tr>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                </table></td>
-->
<!-- del-e by mes 2009/05/12 -->
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td nowrap rowspan="2" bgcolor="#FFFFCC">���݊���</td>
end of comment by nics 2009.02.09 -->
                <td nowrap rowspan="2" bgcolor="#FFFFCC">���݊���</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">����</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">���[�h</td>
                <td nowrap bgcolor="#FFFFCC">���o����</td>
                <td nowrap bgcolor="#FFFFCC">�q�ɓ���</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">�f�o��<BR>����</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">��R��<BR>�ԋp����</td>
              </tr>
              <tr align="center" bgcolor="#FFFFCC">
                <td nowrap>�v��</td>
                <td nowrap>�\��^����</td>
                <td nowrap>��������</td>
                <td nowrap>���o����</td>
                <td nowrap>�\��^����</td>
                <td nowrap>�w���^����</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>����</td>
                <td align="center" rowspan="2" nowrap>
<% ' �d�o�` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(11),11)
%>
                </td>
end of comment by nics 2009.02.09 -->
                <td align="center" rowspan="2" nowrap>
<% ' �O�` - ���݊��� $�ǉ�
    Response.Write DispDateTimeCell(anyTmp(37),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�� �| ���݃X�P�W���[��
    If anyTmp(31)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(31),5)
    If anyTmp(31)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(2)<>"" Then
        bLate = false
        If anyTmp(3)<>"" Then
            If anyTmp(2)<anyTmp(3) Then
                bLate = true
            End If
        End If
        If anyTmp(31)<>"" Then
            If Left(anyTmp(31),10)<Left(anyTmp(2),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(2),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(2),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�� - ���[�h����(�m�F)����
    Response.Write DispDateTimeCell(anyTmp(12),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�� - ���[�h���o����
    Response.Write DispDateTimeCell(anyTmp(13),11)
%>
                </td>
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h - ���o�\�� $�ǉ�
    sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(34)<>"" Then
        If anyTmp(34)<anyTmp(14) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(34),11)
    If anyTmp(34)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ����A�� - �f�o������
    Response.Write DispDateTimeCell(anyTmp(15),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ����A�� - ��R���ԋp����
    Response.Write DispDateTimeCell(anyTmp(16),11)
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(3),11)
%>
                </td>
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h - ���o����
    Response.Write DispDateTimeCell(anyTmp(30),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ����A�� - �q�ɓ�������
    Response.Write DispDateTimeCell(anyTmp(14),11)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.09 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td><font color="#000000" size="-1">�i��1�j�d�o�`�E�O�`�̎����́A���n���Ԃł��B</font></td>-->
<!-- mod-s by mes 2009/05/12 -->
<!--
                <td><font color="#000000" size="-1">�i��3�j�{�^�����N���b�N����Ɠ��Y�`�ł̈ʒu��񓙂��\������܂��i���n���ԕ\���j�B&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�i��4�j�O�`�̎����́A���n���Ԃł��B</font></td>
-->
                <td><font color="#000000" size="-1">�i��3�j�O�`�̎����́A���n���Ԃł��B</font></td>
<!-- mod-e by mes 2009/05/12 -->
<!-- end of mod by nics 2009.02.09 -->
              </tr>
            </table>
<!-- commented by nics 2009.02.09
            <br>
end of comment by nics 2009.02.09 -->
<!---------------�葱���y�є����m�F--------------------------------- commented by nics 2009.02.09 -->
<!-----�葱���---------------->
<!-- commented by nics 2009.02.09
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�葱���y�у^�[�~�i�����o�ۏ��</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
            <br>
end of comment by nics 2009.02.09 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
				  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.02.09 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">�葱���y��<br>�^�[�~�i��<br>���o��<br>���</td>
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td rowspan="3" nowrap bgcolor="#FFCC33">����</td>
end of comment by nics 2009.02.09 -->
<!-- mod by nics 2009.02.09 -->
<!--                <td colspan="4" nowrap bgcolor="#FFCC33">�s���葱��</td>-->
                <td colspan="6" nowrap bgcolor="#FFCC33">�s���葱��</td>
<!-- end of mod by nics 2009.02.09 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">�����<br>
                  DO���s</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">�t���[<br>
                  �^�C��</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">�^�[�~�i��<br>
                  ���o��</td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
                <td rowspan="3" nowrap bgcolor="#FFCC33">�f�B�e���V����<br>�t���[�^�C��</td>
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">�����m�F����</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">���A��</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">�ʔ���</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">�ʊ� /<BR>�ېŗA��<font size="-1"><sup>(��2)</sup></font></td>-->
                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">�ʊ� /<BR>�ېŗA��<font size="-1"><sup>(��5)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td align="center" nowrap bgcolor="#FFFFCC" colspan="2">X������</td>
<!-- end of add by nics 2009.02.09 -->
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">�\��^����</td>
<!-- add by nics 2009.02.09 -->
                <td align="center" nowrap bgcolor="#FFFFCC">�L��</td>
                <td align="center" nowrap bgcolor="#FFFFCC">CY�ԋp</td>
<!-- end of add by nics 2009.02.09 -->
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>���</td>
end of comment by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' �����m�F�\�莞��
    If anyTmp(32)<>"" Then
        If anyTmp(18)<>"" Then
            If anyTmp(32)<anyTmp(18) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(32),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(32),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ���A��
    If anyTmp(17)="S" Then
        Response.Write "�~"
    ElseIf anyTmp(17)="C" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �ʔ���
    If anyTmp(33)<>"" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �ʊց^�ېŗA��
'' 2009/07/10 mod-s ImportContEx��CustStatus��C�A���邢�́ACY���o�������������ꍇ�Ɂ��A����ȊO�́~
'    If anyTmp(19)<>"" Then
'        If anyTmp(19)="O" Or anyTmp(19)="T" Then
'            Response.Write "<a href='#"
'            Response.Write iLineNo
'            Response.Write "' onClick=""winOpen('win1','impdetail-h.asp?line="
'            Response.Write iLineNo
'            Response.Write "',150,150)"">��</a>"
'        Else
'            Response.Write "��"
'        End If
'    Else
'        Response.Write "�~"
'    End If
    If anyTmp(20)="C" Or anyTmp(13)<>"" Then
        Response.Write "��"
    Else
        Response.Write "�~"
    End If
'' 2009/07/10 mod-e ImportContEx��CustStatus��C�A���邢�́ACY���o�������������ꍇ�Ɂ��A����ȊO�́~
%>
                </td>
<!-- add start by nics 2009.02.09  -->
                <td align="center" rowspan="2" nowrap>
<% ' X���L��
'' 2009/07/09 mod-s '-'�Œ��
'    If anyTmp(41)<>"" Then
'        Response.Write anyTmp(41)
'    Else
'        Response.Write "<br>"
'    End If
    Response.Write "�|"
'' 2009/07/09 mod-e '-'�Œ��
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' X��CY�ԋp
'' 2009/07/09 mod-s '-'�Œ��
'    If anyTmp(42)<>"" Then
'        Response.Write anyTmp(42)
'    Else
'        Response.Write "<br>"
'    End If
    Response.Write "�|"
'' 2009/07/09 mod-s '-'�Œ��
%>
                </td>
<!-- add end by nics 2009.02.09  -->
                <td align="center" rowspan="2" nowrap>
<% ' ������c�n���s
'' 2009/07/09 mod-s '-'�Œ��
'    If anyTmp(21)<>"Y" Then
'        Response.Write "�~"
'    Else
'        Response.Write "��"
'    End If
    Response.Write "�|"
'' 2009/07/09 mod-e '-'�Œ��
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �t���[�^�C��
'������ Mod_S  by nics 2009.02.09
'    If anyTmp(22)<>"" Then
'        If anyTmp(22)<DispDateTime(Now,10) Then
'            Response.Write "<font color='#FF0000'>"
'        Else
'            Response.Write "<font color='#000000'>"
'        End If
'        Response.Write DispDateTimeCell(anyTmp(22),5)
'        Response.Write "</font>"
'    Else
'        Response.Write DispDateTimeCell(anyTmp(22),5)
'    End If
'������
    ' anyTmp(13) �� CY���o����[yyyy/mm/dd hh:nn]
    ' anyTmp(22) �� �t���[�^�C��(�t���[�^�C���������t)[yyyy/mm/dd]
    strDisp = DispDateTimeCell(anyTmp(22),5)
    strColor = "#000000"    ' ��
    ' ���o�������ݒ肳��Ă���ꍇ
    If anyTmp(13) <> "" Then
        ' CY���o�������V�X�e�����t�̏ꍇ
        If Left(anyTmp(13),10) < DispDateTime(Now,10) Then
            strDisp = "�|"
        End If
    ' ���o�������ݒ肳��Ă��Ȃ��ꍇ
    Else
        ' �t���[�^�C�����ݒ肳��Ă���ꍇ
        If IsDate(anyTmp(22)) Then
            ' �t���[�^�C�����V�X�e�����t�̏ꍇ
            If anyTmp(22) <= DispDateTime(Now,10) Then
                strColor = "#FF0000"    ' ��
            ' (�t���[�^�C���|�Q��)���V�X�e�����t�̏ꍇ
            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(22))),10) <= DispDateTime(Now,10) Then
                strColor = "#FFA500"    ' ��
            End If
        End If
    End If
    Response.Write "<font color='" & strColor & "'>"
    Response.Write strDisp
    Response.Write "</font>"
'������ Mod_E  by nics 2009.02.09
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�����o��
'' 2009/07/09 mod-s CY���o���������������'��'�A����ȊO��'-'
'    If anyTmp(4)="Y" Then
'        Response.Write "��"
'    ElseIf anyTmp(4)="S" Then
'        Response.Write "��"
'    Else
'        Response.Write "�~"
'    End If
    ' ���o�������ݒ肳��Ă���ꍇ
    If anyTmp(13) <> "" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
'' 2009/07/09 mod-e CY���o���������������'��'�A����ȊO��'-'
%>
                </td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
                <td align="center" rowspan="2" nowrap>
<% ' �f�B�e���V�����t���[�^�C��
'' 2009/07/09 mod-s '-'�Œ��
''������ Mod_S  by nics 2009.02.09
''    Response.Write anyTmp(39)
''������
'    ' anyTmp(39) �� �f�B�e���V�����t���[�^�C��
'    ' anyTmp(16) �� ��o���ԋp����[yyyy/mm/dd hh:nn]
'    ' anyTmp(44) �� ��o���ԋp�\���[yyyy/mm/dd]
'    strDisp = anyTmp(39)
'    strColor = "#000000"    ' ��
'    ' ��o���ԋp�������ݒ肳��Ă���ꍇ
'    If anyTmp(16) <> "" Then
'        ' ��o���ԋp�������V�X�e�����t�̏ꍇ
'        If Left(anyTmp(16),10) < DispDateTime(Now,10) Then
'            strDisp = "�|"
'        End If
'    ' ��o���ԋp�������ݒ肳��Ă��Ȃ��ꍇ
'    Else
'        ' ��o���ԋp�\��������ݒ肳��Ă���ꍇ
'        If IsDate(anyTmp(44)) Then
'            ' ��o���ԋp�\������V�X�e�����t�̏ꍇ
'            If anyTmp(44) <= DispDateTime(Now,10) Then
'                strColor = "#FF0000"    ' ��
'            ' (��o���ԋp�\����|2��)���V�X�e�����t�̏ꍇ
'            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(44))),10) <= DispDateTime(Now,10) Then
'                strColor = "#FFA500"    ' ��
'            End If
'        End If
'    End If
'    Response.Write "<font color='" & strColor & "'>"
'    Response.Write strDisp
'    Response.Write "</font>"
''������ Mod_E  by nics 2009.02.09
    Response.Write DispDateTimeCell("",10)
'' 2009/07/09 mod-e '-'�Œ��
%>
              </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' �����m�F��������
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                </td>
              </tr>
            </table>
			</td>
<!-- commented by nics 2009.02.09
                <td>&nbsp;</td>
                <td valign="top"><table border="1" cellpadding=" 3" cellspacing="1" bgcolor="#FFFFFF">
                  <tr>
                    <td align="center" nowrap bgcolor="#FFCC33">�d�o�`���ʒu���<font size="-1"><sup>(��3)</sup></font></td>
                  </tr>
                  <tr>
                    <td align="center"><table border="0" cellspacing="5">
                        <tr>
                          <td nowrap><a href="javascript:Submit('Form1')" class="splink" onClick="javascript:winOpen('win1','./cct/index.html',560,500)">&nbsp;�Ԙp&nbsp;</a></td>
                          </tr>
                        <tr>
                          <td><a href="javascript:Submit('queryForm')" class="splink" onClick="javascript:winOpen('win1','./sct/index.html',560,500)">&nbsp;�֌�&nbsp;</a></td>
                          </tr>
                    </table></td>
                  </tr>
                </table></td>
end of comment by nics 2009.02.09 -->
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.09 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap>�i��2�j�����N���b�N����ƕېŗA�����Ԃ��\������܂��B<br>
                  �i��3�j�d�o�`���g���ɕ\������Ă���ꍇ�A�{�^�����N���b�N����Ɠ��Y�`�ł̈ʒu��񓙂��\������܂��B
                </td>-->
                <td nowrap><font color="#000000" size="-1">�i��5�j�ېŗA���̏ꍇ�A�����N���b�N����ƕېŗA�����Ԃ��\������܂��B</td>
<!-- end of mod by nics 2009.02.09 -->
              </tr>
            </table>
<!-- commented by nics 2009.02.09
            <br>
end of comment by nics 2009.02.09 -->
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='impreload.asp?request=impdetail.asp'">
</form>
<form name="queryForm" method="post" target="_blank" action="http://oi.sctcn.com/Default.aspx?Action=Nav&Content=CONTAINER%20INFO.%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&sm=CONTAINER%20INFO.">
    <input type="hidden" name="data" value="<%=anyTmp(1)%>">		
    <input type="hidden" name="OrgMenu" value="">
    <input type="hidden" name="targetPage" value="CONTAINER_INFO">
    <input type="hidden" name="nav" value="CONTAINER INFO.                         ">
</form>

<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_infoE.aspx" id="Form1" target="_blank">
    <input type="hidden" name="Image1.x" value="0" />
    <input type="hidden" name="Image1.y" value="0" />
    <input type="hidden" name="__EVENTTARGET" value="" />
    <input type="hidden" name="__EVENTARGUMENT" value="" /> 
    <input type="hidden" name="__VIEWSTATE" value="dDwtMzMwNTk0MTUxOztsPEltYWdlMTs+Po9koK7lFKyndTfCh4n1g7KjLvsH" />
    <input type="hidden" name="cont_no" id="cont_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="wyex" value="wyE" />
</form>

<%
    ' ������ʂ��璼�ڔ��ł����Ƃ��͕\������
    If bSingle Then
        Response.Write "<form action='impcsvout.asp'>"
        Response.Write "<center>"
        Response.Write "<input type='submit' name='submit' value='CSV�t�@�C���o��'>�@"
        Response.Write "<a href='../help05.asp'>CSV�t�@�C���o�͂Ƃ́H</a>"
        Response.Write "</center>"
        Response.Write "</form>"
    End If
%>
          </td>
        </tr>
      </table>
      </center>
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
<!-------------�ڍ׉�ʏI���--------------------------->
<%
    DispMenuBarBack "JavaScript:FancBack()"

'�s�v
'    ' ������ʂ��璼�ڔ��ł����Ƃ�
'    If bSingle Then
'        DispMenuBarBack "impentry.asp"
'    Else
'        If iReturn=1 Then
'            DispMenuBarBack "implist1.asp"
'        ElseIf iReturn=2 Then
'            DispMenuBarBack "implist2.asp"
'        ElseIf iReturn=3 Then
'            DispMenuBarBack "implist3.asp"
'        Else
'            DispMenuBarBack "implist.asp"
'        End If
'    End If
%>
</body>
</html>
