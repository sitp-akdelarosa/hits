<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	�y�R���e�i�����́z	�X�V���\���`�F�b�N�ATemp�t�@�C���쐬
	'
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

	' �����ꗗ�\���ő�l
	Dim sUser,sUserNo
    sUser    = UCase(Trim(Request.form("contuser")))
    sUserNo  = UCase(Trim(Request.form("contuserno")))
    sBooking = UCase(Trim(Request.form("contbooking")))

	' �����ꗗ�\���ő�l
	Dim iMaxCount
	iMaxCount = 100
    ' �G���[�t���O�̃N���A
    bError = false
	' �C�݃R�[�h�擾
	sSosin = Trim(Session.Contents("userid"))

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' ���R�[�h���̎擾
	Dim iRecCount,strErrMsg
    ConnectSvr conn, rsd
	sql = "SELECT count(*) FROM ExportCargoInfo WHERE Forwarder='" & sSosin & "'"
	If sUser<>"" Then
		sql = sql & " AND Shipper='" & sUser & "'"
	End If
	If sUserNo<>"" Then
		sql = sql & " AND ShipCtrl='" & sUserNo & "'"
	End If
	If sBooking<>"" Then
		sql = sql & " AND BookNo='" & sBooking & "'"
	End If

	rsd.Open sql, conn, 0, 1, 1
	If Not rsd.EOF Then
	    iRecCount = rsd(0)
	Else
	    bError = true
		strErrMsg = "DB�ڑ��G���["
	End If
	rsd.Close

	If iRecCount>iMaxCount Then
	    bError = true
		strErrMsg = "�����Ώی������ő�l�𒴂��Ă��܂��B<BR>�i�荞�݂����ĉ������B"
	Else If iRecCount=0 Then
	    bError = true
		strErrMsg = "�Ώۃf�[�^�����݂��܂���B"
	Else
		Dim strOut,bWrite
		' Temp�t�@�C�������o��
	    bWrite = 0        '�o�̓��R�[�h����

	    ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
	    strFileName="./temp/" & strFileName
	    ' �e���|�����t�@�C����Open
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

		sql = "SELECT ExportCargoInfo.Shipper,ExportCargoInfo.ShipCtrl,ExportCargoInfo.BookNo," & _
				"ExportCargoInfo.ContNo,ExportCargoInfo.VslCode,ExportCargoInfo.LdVoyage," & _
				"ExportCont.RFFlag, ExportCont.DGFlag," & _
				"Container.ContWeight,Container.SealNo,Container.CargoWeight " & _
			  "FROM ExportCargoInfo,Container,ExportCont " & _
			  "WHERE ExportCargoInfo.Forwarder='" & sSosin & "' And "

		If sUser<>"" Then
			sql = sql & "ExportCargoInfo.Shipper='" & sUser & "' And "
		End If
		If sUserNo<>"" Then
			sql = sql & "ExportCargoInfo.ShipCtrl='" & sUserNo & "' And "
		End If
		If sBooking<>"" Then
			sql = sql & "ExportCargoInfo.BookNo='" & sBooking & "' And "
		End If

		sql = sql & "ExportCont.VslCode=*ExportCargoInfo.VslCode And " & _
					"ExportCont.ContNo=*ExportCargoInfo.ContNo And " & _
					"ExportCont.BookNo=*ExportCargoInfo.BookNo And " & _
					"Container.VslCode=*ExportCargoInfo.VslCode And " & _
					"Container.ContNo=*ExportCargoInfo.ContNo"

	    rsd.Open sql, conn, 0, 1, 1

	    Do While Not rsd.EOF
	        strOut = Trim(rsd("VslCode")) & ","
     		strOut = strOut & Trim(rsd("LdVoyage")) & ","
     		strOut = strOut & Trim(rsd("Shipper")) & ","
     		strOut = strOut & Trim(rsd("ShipCtrl")) & ","
     		strOut = strOut & Trim(rsd("BookNo")) & ","
     		strOut = strOut & Trim(rsd("ContNo")) & ","
     		strOut = strOut & Trim(rsd("SealNo")) & ","
     		strOut = strOut & Trim(rsd("CargoWeight"))/10 & ","
     		strOut = strOut & Trim(rsd("ContWeight"))/10 & ","
     		strOut = strOut & Trim(rsd("RFFlag")) & ","
     		strOut = strOut & Trim(rsd("DGFlag"))

	        ti.WriteLine strOut
	        bWrite = bWrite + 1

	        rsd.MoveNext
	    Loop

  		rsd.Close

	End If
	End If

    If bError Then
        strOption = sUser & "/" & sUserNo & "/" & sBooking & "," & "���͓��e�̐���:1(���)"
    Else
        strOption = sUser & "/" & sUserNo & "/" & sBooking & "," & "���͓��e�̐���:0(������)"
    End If

    WriteLog fs, "4105","�C�ݓ��͗A�o�R���e�i���","10", strOption

	If Not bError Then
        Response.Redirect "ms-kaika-expcontinfo-list.asp"
	Else
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
		<BR>
		<BR>
		<BR>
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�X�V�Ώۈꗗ</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strErrMsg 
%>
			</center>
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
	End If
%>
