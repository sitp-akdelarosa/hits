<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "vslentry.asp"

    ' �w������̎擾
    Dim strShipLine
    strShipLine = Request.QueryString("shipline")

    ' �{�D���Ï�񃌃R�[�h�̎擾
    ConnectSvr conn, rsd

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQL�𔭍s���ăJ�����g�|�[�g������
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    sql = "SELECT VslSchedule.VslCode, VslSchedule.VoyCtrl, VslSchedule.DsVoyage, VslSchedule.LdVoyage, mVessel.FullName FROM VslSchedule, VslPort, mVessel "
    sql = sql & " WHERE VslSchedule.ShipLine='" & strShipLine & "' And VslPort.VslCode=VslSchedule.VslCode And VslPort.VoyCtrl=VslSchedule.VoyCtrl And " & _
          "VslPort.PortCode='" & strPort & "' And VslPort.ETA>=DATEADD(day,-14,CONVERT(datetime,'" & DispDateTime(Now,0) & "')) And " & _
          "mVessel.VslCode=*VslSchedule.VslCode"
    'SQL�𔭍s���Ė{�D���Èꗗ������
    rsd.Open sql, conn, 0, 1, 1
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------����������--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/vslentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <br>�@
      <br>�@
      <center>
      <table width="100%">
        <tr>
          <td>  
            <br>
            <center>
            <table width="60%" border="1" cellspacing="1" cellpadding="3">
              <tr bgcolor="#FFCC33"> 
                <td nowrap width="60%" align=center valign=middle>�D��</td>
                <td nowrap width="20%" align=center valign=middle>���q</td>
                <td nowrap width="20%" align=center valign=middle>�R�[���T�C��</td>
              </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<%
    Do While Not rsd.EOF
        Response.Write "<tr bgcolor='#FFFFFF'>"
        Response.Write "<td nowrap align=left valign=middle>"
        Response.Write "<a href='vslschedule.asp?vslcode=" & Trim(rsd("VslCode")) & "&voyctrl=" & Trim(rsd("VoyCtrl")) & "'>" & rsd("FullName") & "</a>"
        Response.Write "</td>"
        strVoyage = Trim(rsd("DsVoyage"))
        strTemp = Trim(rsd("LdVoyage"))
        If strVoyage<>strTemp Then
           If strVoyage<>"" Then
               If strTemp<>"" Then
                   strVoyage = strVoyage & "/" & strTemp
               End If
           Else
               strVoyage = strTemp
           End If
        End If
        Response.Write "<td>" & strVoyage & "</td>"
        Response.Write "<td>" & rsd("VslCode") & "</td>"
        Response.Write "</tr>"

        rsd.MoveNext
    Loop
    rsd.Close
    conn.Close
%>
<!-- �����܂� -->
            </table>
            </center>
          </td>
        </tr>
      </table>
      <br>
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
<!-------------��ʏI���--------------------------->
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �{�D���Ɖ�
    WriteLog fs, "�{�D���Ɖ�", "�I��D��," & strShipLine
%>
