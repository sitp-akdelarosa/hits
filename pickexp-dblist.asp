<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	�y�C�ݓ��́z	�X�V���\���`�F�b�N�ATemp�t�@�C���쐬
	'
	'ExportCargoInfo ��R���󂯎����t�񖼁i�V�K�ǉ����j
	Dim receivedate
	receivedate = "PickDate"

%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "pickselect.asp"

	' �����ꗗ�\���ő�l
	Dim sUser,sUserNo
	sUser   = UCase(Trim(Request.form("suser")))
	sUserNo = UCase(Trim(Request.form("suserno")))
    strUserKind=Session.Contents("userkind")

	' �����ꗗ�\���ő�l
	Dim iMaxCount
	iMaxCount = 100
    ' �G���[�t���O�̃N���A
    bError = false
	' �C�݁i�׎�j�R�[�h�擾
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
	If strUserKind="�C��" Then
		sql = "SELECT count(*) FROM ExportCargoInfo WHERE Forwarder='" & sSosin & "'"
	Else
		sql = "SELECT count(*) FROM ExportCargoInfo WHERE Shipper='" & sSosin & "'"
	End If

	If sUser<>"" Then
		If strUserKind="�C��" Then
			sql = sql & " AND Shipper='" & sUser & "'"
		Else
			sql = sql & " AND Forwarder='" & sUser & "'"
		End If
	End If
	If sUserNo<>"" Then
		sql = sql & " AND ShipCtrl='" & sUserNo & "'"
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
	ElseIf iRecCount=0 Then
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

		sql = "SELECT Shipper,ShipCtrl,VslCode,LdVoyage,BookNo,Trucker,WHArTime,CYRecDate," & _
			  "ContSize,ContType,ContHeight,PickPlace,Remark," & receivedate & ",OpeCode,Forwarder " & _
			  "FROM ExportCargoInfo "
		If strUserKind="�C��" Then
			sql = sql &  "WHERE Forwarder='" & sSosin & "'"
		Else
			sql = sql &  "WHERE Shipper='" & sSosin & "'"
		End If

		If sUser<>"" Then
			If strUserKind="�C��" Then
				sql = sql & " AND Shipper='" & sUser & "'"
			Else
				sql = sql & " AND Forwarder='" & sUser & "'"
			End If
		End If
		If sUserNo<>"" Then
			sql = sql & " AND ShipCtrl='" & sUserNo & "'"
		End If

	    rsd.Open sql, conn, 0, 1, 1

	    Do While Not rsd.EOF
     		strOut = Trim(rsd("VslCode")) & ","
     		strOut = strOut & Trim(rsd("LdVoyage")) & ","
	        strOut = strOut & Trim(rsd("Shipper")) & ","
     		strOut = strOut & Trim(rsd("ShipCtrl")) & ","
     		strOut = strOut & Trim(rsd("BookNo")) & ","
     		strOut = strOut & Trim(rsd("Trucker")) & ","
     		strOut = strOut & DispDateTime(rsd("WHArTime"),0) & ","
     		strOut = strOut & DispDateTime(rsd("CYRecDate"),10) & ","
     		strOut = strOut & Trim(rsd("ContSize")) & ","
     		strOut = strOut & Trim(rsd("ContType")) & ","
     		strOut = strOut & Trim(rsd("ContHeight")) & ","
    		strOut = strOut & Trim(rsd("Remark")) & ","
    		strOut = strOut & Trim(rsd("PickPlace")) & ","
    		strOut = strOut & DispDateTime(rsd(receivedate),10) & ","
     		strOut = strOut & Trim(rsd("OpeCode")) & ","
     		strOut = strOut & Trim(rsd("Forwarder"))

	        ti.WriteLine strOut
	        bWrite = bWrite + 1

	        rsd.MoveNext
	    Loop

  		rsd.Close

	End If

	If Not bError Then
		strOption = sUser & "/" & sUserNo & "," & "���͓��e�̐���:0(������)"
	Else
		strOption = sUser & "/" & sUserNo & "," & "���͓��e�̐���:1(���)"
	End If

	If strUserKind="�C��" Then
		iWrkNum = "11"
	Else
		iWrkNum = "12"
	End If
    WriteLog fs, "a110","��R���s�b�N�A�b�v�V�X�e��-��R���s�b�N�A�b�v�˗�",iWrkNum, strOption


	If Not bError Then
        Response.Redirect "pickexp-list.asp"
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
<%
	If strUserKind="�C��" Then
		titlegif = "pickkat"
	Else
		titlegif = "picknit"
	End If
%>
          <td rowspan=2><img src="gif/<%=titlegif%>.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>�X�V�Ώۈꗗ</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
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
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

<%
	End If
%>
