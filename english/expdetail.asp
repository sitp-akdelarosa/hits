<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "expentry.asp"

    ' �w������̎擾
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

    ' �\�����[�h�̎擾
    Dim bDispMode          ' true=�R���e�i���� / false=�u�b�L���O����
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
        Response.Redirect "expentry.asp"             '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="../temp/" & strFileName

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

    ' �A�o�R���e�i�Ɖ�ڍ�
    WriteLog fs, "1007","�A�o�R���e�i�Ɖ�-�P�ƃR���e�i","00", anyTmp(1) & ","

    Session.Contents("dispcntnr")=anyTmp(1)     ' �\���R���e�iNo.���L��
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="../index.css" rel="stylesheet" type="text/css">
<SCRIPT Language="JavaScript">
<!--
function Submit(formName){
    document.forms[formName].submit();
}
// -->
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#0000FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF">
<!-------------start--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/expdetailt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.17
'	DisplayCodeListButton
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
End of comment by seiko-denki 2003.07.17 -->
<!-- mod by nics 2009.03.04 -->
<!--		<table width=95% cellpadding=3>-->
		<table width=95% cellpadding=0>
<!-- end of mod by nics 2009.03.04 -->
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
<!-- mod by nics 2009.03.04 -->
<!--          <td>�@<br>-->
          <td>
<!-- end of mod by nics 2009.03.04 -->
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' Booking No
    If Not bDispMode Then
        Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>Booking No</b></font></td>"
        Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
    End If
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>Container No.</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �R���e�iNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--			<BR>-->
			<font size="-2"><BR></font>
<!-- end of mod by nics 2009.03.04 -->
<!---------------��{���------------------------------------------- commented by nics 2009.02.12 -->
<!-- commented by nics 2009.03.04
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Basis information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
end of comment by nics 2009.03.04 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<!-- add by nics 2009.03.04 -->
                <td rowspan="3" nowrap bgcolor="#6495ED">Basis<br>information</td>
<!-- end of add by nics 2009.03.04 -->
<!-- add by mes(2005/3/28) �e�A�E�F�C�g�ǉ� -->
<!--                <td colspan="4" nowrap>��R���e�i</td>-->
<!--                <td colspan="5" nowrap>��R���e�i</td>-->
<!-- Mod-S MES Aoyagi 2010.11.27 �R���e�i�^�C�v�ǉ� -->
		<td colspan="6" nowrap>��R���e�i</td>
<!-- Mod-E MES Aoyagi 2010.11.27 �R���e�i�^�C�v�ǉ� -->
<!-- end mes -->
<!-- mod by nics 2009.03.04 -->
<!--                <td colspan="5" nowrap bgcolor="#FFCC33">Full Container</td>-->
                <td colspan="4" nowrap bgcolor="#FFCC33">Full Container</td>
<!-- end of mod by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap colspan="2">Terminal open</td>
end of comment by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td rowspan="2" nowrap bgcolor="#FFCC33">Shipping Yard<br>(code)</td>
                <td rowspan="2" nowrap bgcolor="#FFCC33">Operater</td>
<!-- end of add by nics 2009.03.04 -->
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap><font color="#000000">Pickup place</font></td>
                <td nowrap><font color="#000000">size</font></td>
<!-- Add-S MES Aoyagi 2010.11.23 �R���e�i�^�C�v�ǉ� -->
		<td nowrap><font color="#000000">type</font></td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
                <td nowrap>height<font size="-1"><sup>(*1)</sup></font></td>
<!-- add by mes(2005/3/28) �e�A�E�F�C�g�ǉ� -->
                <td nowrap><font color="#000000">TW</font></td>
<!-- end mes -->
                <td nowrap><font color="#000000">Reefer</font></td>
                <td nowrap><font color="#000000">Seal No.</font></td>
                <td nowrap><font color="#000000">CW(t)</font></td>
                <td nowrap><font color="#000000">GW(t)</font></td>
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap><font color="#000000">Hazard</font></td>-->
                <td nowrap><font color="#000000">Hazard<font size="-1"><sup>(��2)</sup></font></font></td>
<!-- end of mod by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                <td nowrap><font color="#000000">Shipping
yard</font></td>
                <td nowrap><font color="#000000">open</font></td>
                <td nowrap>close</td>
end of comment by nics 2009.03.04 -->
              </tr>
              <tr> 
                <td nowrap align="center">
<% ' ��R�����ꏊ
    If anyTmp(2)<>"" Then
        Response.Write anyTmp(2)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �T�C�Y
    If anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>

<!-- Add-S MES Aoyagi 2010.11.23 -->
<% ' �^�C�v
    If anyTmp(39)<>"" Then
        Response.Write anyTmp(39)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<% ' ����
    If anyTmp(4)<>"" Then
        Response.Write anyTmp(4)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add by mes(2005/3/28) �e�A�E�F�C�g�ǉ� -->
                <td align="center" nowrap>
<% ' �e�A�E�F�C�g
    If anyTmp(32)<>"" And anyTmp(32)>0 Then
    	If anyTmp(32)<100 then
	        dWeight=anyTmp(32) * 100
	    Else
	        dWeight=anyTmp(32)
	    End If
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
<!-- end mes -->
                <td align="center" nowrap>
<% ' ���[�t�@�[
    If anyTmp(5)="R" Then
        Response.Write "��"
    ElseIf anyTmp(5)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �V�[��No.
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �ݕ��d�� $�ǉ�
    If anyTmp(27)<>"" And anyTmp(27)<>"0" Then
        dWeight=anyTmp(27) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ���d��
    If anyTmp(8)<>"" And anyTmp(8)<>"0" Then
        dWeight=anyTmp(8) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �댯�i
    If anyTmp(31)="H" Then
        Response.Write "��"
    ElseIf anyTmp(31)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td align="center" nowrap>
<% ' �����^�[�~�i����
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' CY�I�[�v��
    Response.Write DispDateTimeCell(anyTmp(9),5)
%>
                </td>
                <td align="center" nowrap>
<% ' CY�N���[�Y
    Response.Write DispDateTimeCell(anyTmp(10),5)
%>
                </td>
end of comment by nics 2009.03.04 -->
<!-- add by nics 2009.02.12 -->
                <td align="center" nowrap>
<% ' �����^�[�~�i��(���u�ꏊ�R�[�h)
    strDisp = "<br>"
    If anyTmp(6) <> "" Then
        strDisp = anyTmp(6)
        If anyTmp(36) <> "" Then
            strDisp = strDisp & "(" & anyTmp(36) & ")"
        End If
    End If
    Response.Write strDisp
%>
                </td>
                <td align="center" nowrap>
<% ' �S���I�y���[�^
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- end of add by nics 2009.02.12 -->
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.04 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td><font color="#000000" size="-1">(*1)96=HC</font></td>-->
                <td><font color="#000000" size="-1">(*1)96=HC &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (*2)Presence of dangerous articles related to Fire Defense Law</font></td>
<!-- end of mod by nics 2009.03.04 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.04
            <BR>
end of comment by nics 2009.03.04 -->
<!---------------�{�D���------------------------------------------- commented by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Vessel Information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
end of comment by nics 2009.03.04 -->
            <table border=1 cellpadding="3" cellspacing="1">
<!-- mod by nics 2009.03.04 -->
<!--              <tr> -->
              <tr align="center" bgcolor="#FFCC33"> 
<!-- end of mod by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td rowspan="2" nowrap bgcolor="#6495ED">Vessel<br>information</td>
<!-- end of add by nics 2009.03.04 -->
                <td bgcolor="#FFCC33" nowrap>Ship's Line</td>
<!-- add by nics 2009.03.04 -->
                <td bgcolor="#FFCC33" nowrap>Vessel Name</td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFCC33" nowrap>Discharge Port</td>
              </tr> 
              <tr align="center"> 
<!-- end of add by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF">
<% ' �D��
    If anyTmp(11)<>"" Then
        Response.Write anyTmp(11)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Vessel Name</td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' �D��
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' ���q
    If anyTmp(13)<>"" Then
        Response.Write anyTmp(13)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Discharge Port</td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' �d���`
    If anyTmp(14)<>"" Then
        Response.Write anyTmp(14)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--            <BR>-->
            <font size="-2"><BR></font>
<!-- end of mod by nics 2009.03.04 -->
<!---------------�ʒu���------------------------------------------- commented by nics 2009.02.12 -->
<!-- commented by nics 2009.03.04
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Position information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table> 
            <br>
end of comment by nics 2009.03.04 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr align="center" bgcolor="#FFCC33">
<!-- add by nics 2009.03.04 -->
                    <td rowspan="5" nowrap bgcolor="#6495ED">Position<br>information</td>
<!-- end of add by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                    <td nowrap>Place</td>
end of comment by nics 2009.03.04 -->
                    <td colspan="3" nowrap>Conveyance by land</td>
                    <td nowrap bgcolor="#FFCC33">Stock Yard</td>
                    <td colspan="4" nowrap bgcolor="#FFCC33">Terminal</td>
<!-- mod by nics 2009.03.04 -->
<!--                    <td bgcolor="#FFCC33" nowrap>Discharge Port</td>-->
                    <td bgcolor="#FFCC33" nowrap>Discharge Port<font size="-1"><sup>(��3)</sup></font></td>
<!-- end of mod by nics 2009.03.04 -->
                  </tr>
                  <tr align="center" bgcolor="#FFFF99">
<!-- commented by nics 2009.03.04
                    <td nowrap rowspan="2">Process</td>
end of comment by nics 2009.03.04 -->
<!-- mod by nics 2009.03.04 -->
<!--                    <td nowrap rowspan="2">Empty container pickup<br>
                      time</td>-->
                    <td nowrap rowspan="2">Empty container<br>pickup time</td>
<!-- end of mod by nics 2009.03.04 -->
                    <td nowrap>Warehouse arrival time</td>
<!-- mod by nics 2009.03.04 -->
<!--                    <td nowrap rowspan="2">Vanning time</td>-->
                    <td nowrap rowspan="2">Vanning<br>time</td>
<!-- end of mod by nics 2009.03.04 -->
                    <td nowrap>SY in</td>
                    <td nowrap>CY in time</td>
                    <td nowrap rowspan="2">Loading<br>
                      time</td>
                    <td nowrap colspan="2">Departure time</td>
<!-- commented by nics 2009.03.04
                    <td nowrap>arrival time<font size="-1"><sup>(*3)</sup></font></td>
end of comment by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                    <td rowspan="4" align="center" bgcolor="#FFFFFF"><table border="0" cellspacing="5">
                      <tr>
                        <td nowrap><a href="javascript:Submit('Form1')" class="splink">CHIWAN</a></td>
	                    <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                        </tr>
                      <tr>
                        <td><a href="javascript:Submit('queryForm')" class="splink">SHEKOU</a></td>
	                    <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                        </tr>
                      <tr>
	                    <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
	                    <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                        </tr>
                    </table></td>
<!-- end of add by nics 2009.03.04 -->
                  </tr>
                  <tr align="center" bgcolor="#FFFF99">
<!-- mod by nics 2009.03.04 -->
<!--                    <td nowrap><font color="#000000">Instruction<font size="-1"><sup>(*2)</sup></font>/Actual time</font></td>-->
                    <td nowrap><font color="#000000">Instruction/Actual time</font></td>
<!-- end of mod by nics 2009.03.04 -->
                    <td nowrap>Reservation<br>
                      /Actual</td>
                    <td nowrap>Instruction<br>
                      /Actual</td>
                    <td nowrap>Estimate</td>
                    <td nowrap>Intended<br>
                      /Actual</td>
<!-- commented by nics 2009.03.04
                    <td nowrap>Intended<br>
                      /Actual</td>
end of comment by nics 2009.03.04 -->
                  </tr>
                  <tr>
<!-- commented by nics 2009.03.04
                    <td nowrap rowspan="2" bgcolor="#FFFFCC" align="center">Time</td>
end of comment by nics 2009.03.04 -->
                    <td rowspan="2" align="center" nowrap><% ' ����^�� - ��R�����
    Response.Write DispDateTimeCell(anyTmp(16),11)
%>
                    </td>
                    <td align="center" nowrap><% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(26)<>"" Then
        If anyTmp(26)<anyTmp(17) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(26),11)
    If anyTmp(26)<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td rowspan="2" align="center" nowrap><% ' ����^�� - �o���j���O
    Response.Write DispDateTimeCell(anyTmp(18),11)
%>
                    </td>
                    <td align="center" nowrap><% ' �X�g�b�N���[�h - �����\�� $�ǉ�
    sTemp=DispReserveCell(anyTmp(28),anyTmp(29),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td align="center" nowrap><% ' �^�[�~�i�� - CY�����w�� $�ǉ�
    If anyTmp(30)<>"" Then
        If Left(anyTmp(30),10)<Left(anyTmp(19),10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(30),5)
    If anyTmp(30)<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td rowspan="2" align="center" nowrap><% ' �^�[�~�i�� - �D�ϊ���
    Response.Write DispDateTimeCell(anyTmp(20),11)
%>
                    </td>
                    <td rowspan="2" align="center" nowrap><% ' �^�[�~�i�� - ���݃X�P�W���[��
    If anyTmp(25)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(25),5)
    If anyTmp(25)<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td align="center" nowrap><% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(15)<>"" Then
        bLate = false
        If anyTmp(21)<>"" Then
            If anyTmp(15)<anyTmp(21) Then
                bLate = true
            End If
        End If
        If anyTmp(25)<>"" Then
            If Left(anyTmp(25),10)<Left(anyTmp(15),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(15),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(15),11)
    End If
%>
                    </td>
<!-- commented by nics 2009.03.04
                    <td align="center" nowrap><% ' �d���` - ���ݗ\��
    If anyTmp(23)<>"" Then
        If anyTmp(22)<>"" Then
            If anyTmp(23)<anyTmp(22) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(23),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(23),11)
    End If
%>
                    </td>
end of comment by nics 2009.03.04 -->
                  </tr>
                  <tr>
                    <td align="center" nowrap><% ' ����^�� - �q�ɓ���
    Response.Write DispDateTimeCell(anyTmp(17),11)
%>
                    </td>
                    <td align="center" nowrap><% ' �X�g�b�N���[�h - ��������
    Response.Write DispDateTimeCell(anyTmp(24),11)
%>
                    </td>
                    <td align="center" nowrap><% ' �^�[�~�i�� - CY��������
    Response.Write DispDateTimeCell(anyTmp(19),11)
%>
                    </td>
                    <td align="center" nowrap><% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(21),11)
%>
                    </td>
<!-- commented by nics 2009.03.04
                    <td align="center" nowrap><% ' �d���` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(22),11)
%>
                    </td>
end of comment by nics 2009.03.04 -->
                  </tr>
                </table></td>
                <td>&nbsp;</td>
<!-- commented by nics 2009.03.04
                <td valign="top"><table border="1" cellpadding=" 3" cellspacing="1" bgcolor="#FFFFFF">
                  <tr>
                    <td nowrap bgcolor="#FFCC33">Discharge Port<br> 
                      Information<font size="-1"><sup>(*4)</sup></font></td>
                  </tr>
                  <tr>
                    <td align="center"><table border="0" cellspacing="5">
                      <tr>
                        <td nowrap><a href="javascript:Submit('Form1')" class="splink">CHIWAN</a></td>
                        </tr>
                      <tr>
                        <td><a href="javascript:Submit('queryForm')" class="splink">SHEKOU</a></td>
                        </tr>

                    </table></td>
                  </tr>
                </table></td>
end of comment by nics 2009.03.04 -->
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.04 -->
              <tr> 
                <td width="15" rowspan="2">&nbsp;</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap><font color="#000000" size="-1">(*2) Displayed by red letters when the arrival was lated.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (*3) The time of the discharge port is local time.</font></td>-->
                <td><font color="#000000" size="-1">(*3) The location information in the port is displayed when clicking on a red button. </font></td>
<!-- end of mod by nics 2009.03.04 -->
              </tr>
<!-- commented by nics 2009.03.04
              <tr>
                <td><font color="#000000" size="-1">(*4) The location information in the port is displayed when clicking on a red button. </font></td>
              </tr>
end of comment by nics 2009.03.04 -->
            </table>
<!-- commented by nics 2009.03.04
            <BR>
end of comment by nics 2009.03.04 -->
<!---------------�葱���y�є����m�F--------------------------------- commented by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td rowspan="4" nowrap bgcolor="#6495ED">Procedure<br>and<br>Delivery<br>permission<br>information</td>
                <td bgcolor="#FFCC33" nowrap colspan="2">Terminal open</td>
<!-- mod by mes aoyagi 2010.5.13 -->
<!-- commented by nics 2010.02.22 -->
                <td colspan="2" nowrap bgcolor="#FFCC33">Administrative procedure</td>
<!-- end of comment by nics 2010.02.22 -->
<!-- mod by nics 2010.02.22 -->
<!--                <td colspan="3" nowrap bgcolor="#FFCC33">Administrative procedure</td> -->
<!-- end of mod by nics 2010.02.22 -->
<!-- end of mod by mes aoyagi 2010.5.13 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">Confirmation of arrival</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99">
                <td rowspan="2" nowrap><font color="#000000">open</font></td>
                <td rowspan="2" nowrap>close</td>
                <td colspan="2" nowrap>X-ray Inspection</td>
<!-- del by mes aoyagi 2010.5.14 -->
<!-- add by nics 2010.02.22 -->
<!--                <td rowspan="2" nowrap>Clearance</td> -->
<!-- end of add by nics 2010.02.22 -->
<!-- del by mes aoyagi 2010.5.14 -->
              </tr>
              <tr align="center" bgcolor="#FFFF99">
                <td colspan="1" nowrap>Need</td>
                <td colspan="1" nowrap>Return</td>
              </tr>

              <tr> 
                <td align="center" nowrap>
<% ' CY�I�[�v��
    Response.Write DispDateTimeCell(anyTmp(9),5)
%>
                </td>
                <td align="center" nowrap>
<% ' CY�N���[�Y
    Response.Write DispDateTimeCell(anyTmp(10),5)
%>
                </td>
                <td align="center" nowrap>
<% ' X���L��
        If anyTmp(33)<>"" Then
            Response.Write anyTmp(33)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td align="center" nowrap>
<% ' X��CY�ԋp
        If anyTmp(34)<>"" Then
            Response.Write anyTmp(34)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
<!-- del by mes aoyagi 2010.5.13 -->
<!-- add by nics 2010.02.22 -->
<!--                <td align="center" nowrap>
<% ' �ʊ�
        If anyTmp(38)<>"" Then
            Response.Write anyTmp(38)
        Else
            Response.Write "<br>"
        End If
%>
                </td> -->
<!-- end of add by nics 2010.02.22 -->
<!-- del by mes aoyagi 2010.5.13 -->
                <td align="center" nowrap>
<% ' �^�[�~�i�������m�F
        If anyTmp(35)<>"" Then
            Response.Write anyTmp(35)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
              </tr>
                  </table>
                </td>
              </tr>
            </table>
<!-- end of add by nics 2009.03.04 -->
<form>
      <input type=button value='Display Update' OnClick="JavaScript:window.location.href='expreload.asp?request=expdetail.asp'">
</form>
<form name="queryForm" method="post" target="_blank" action="http://oi.sctcn.com/Default.aspx?Action=Nav&Content=CONTAINER%20INFO.%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&sm=CONTAINER%20INFO.">
    <input type="hidden" name="data" value="<%=anyTmp(1)%>">		
    <input type="hidden" name="OrgMenu" value="">
    <input type="hidden" name="targetPage" value="CONTAINER_INFO">
    <input type="hidden" name="nav" value="CONTAINER INFO.                         ">
</form>

<!--
<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_infoE.aspx" id="Form1" target="_blank">
    <input type="hidden" name="Image1.x" value="0" />
    <input type="hidden" name="Image1.y" value="0" />
    <input type="hidden" name="__EVENTTARGET" value="" />
    <input type="hidden" name="__EVENTARGUMENT" value="" />
    <input type="hidden" name="__VIEWSTATE" value="dDwtMzMwNTk0MTUxOztsPEltYWdlMTs+Po9koK7lFKyndTfCh4n1g7KjLvsH" />
    <input type="hidden" name="cont_no" id="cont_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="wyex" value="wyE" />
</form>
-->

<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_info_show.aspx" id="Form1" target="_blank">
    <input type="hidden" name="Image1.x" value="0" />
    <input type="hidden" name="Image1.y" value="0" />
    <input type="hidden" name="cont_no" id="cont_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="wyex" value="wyE" />
</form>


<%
    ' ������ʂ��璼�ڔ��ł����Ƃ��͕\������
    If bSingle Then
        Response.Write "<form action='expcsvout.asp'>"
        Response.Write "<center>"
        Response.Write "<input type='submit' name='submit' value='CSV file output'>�@"
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
    ' ������ʂ��璼�ڔ��ł����Ƃ�
    If bSingle Then
        DispMenuBarBack "expentry.asp"
    Else
        If iReturn=1 Then
            DispMenuBarBack "explist1.asp"
        ElseIf iReturn=2 Then
            DispMenuBarBack "explist2.asp"
        ElseIf iReturn=3 Then
            DispMenuBarBack "explist3.asp"
        Else
            DispMenuBarBack "explist.asp"
        End If
    End If
%>
</body>
</html>
