<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
'    CheckLogin "expentry.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �w������̎擾
    Dim strBookingNo,strBookingNoLog
    strBookingNo = UCase(Trim(Request.Form("booking")))
	strBookingNoLog = strBookingNo
'2006/03/06 add-s h.matsuda
	dim ShipLine,ShoriMode
	ShoriMode = Trim(Request("ShoriMode"))
	ShipLine = Trim(Request("ShipLine"))
'2006/03/06 add-e h.matsuda
    If strBookingNo="" Then
        ' �����w��̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
        bError = true
        strError = "�Q�Ƃ�����Booking No.����͂��Ă��������B"
        strOption = "���͓��e�̐���:1(���)"
    Else
        ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
        Dim strFileName
        strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
        Session.Contents("tempfile")=strFileName

        ' �R���e�i��񃌃R�[�h�̎擾
        ConnectSvr conn, rsd

        iCanma = InStr(strBookingNo,",")
        Do While iCanma>0
            sTemp = Trim(Left(strBookingNo,iCanma-1))
            strBookingNo = Right(strBookingNo,Len(strBookingNo)-iCanma)
            If sWhere<>"" Then
                sWhere = sWhere & " Or Booking.BookNo='" & sTemp & "'"
				strBookingNoLog = strBookingNoLog & "/" & sTemp
            Else
                sWhere = "Booking.BookNo='" & sTemp & "'"
				strBookingNoLog = sTemp
            End If
            iCanma = InStr(strBookingNo,",")
        Loop
        If sWhere<>"" Then
            sWhere = sWhere & " Or Booking.BookNo='" & Trim(strBookingNo) & "'"
        Else
            sWhere = "Booking.BookNo='" & Trim(strBookingNo) & "'"
        End If

        ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
        strFileName="./temp/" & strFileName
        ' �e���|�����t�@�C����Open
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)


'        sql = "SELECT Booking.BookNo," &_
'					  "Booking.VslCode," &_
'					  "Booking.ShipLine," &_
'					  "mShipLine.FullName ShipLineName," &_
'					  "mVessel.FullName ShipName," &_
'					  "VslSchedule.LdVoyage," &_
'					  "Booking.DPort," &_
'					  "mPort.FullName PortName," &_
'					  "Pickup.PickPlace," &_
'					  "Pickup.ContSize," &_
'					  "Pickup.ContType," &_
'					  "Pickup.ContHeight," &_
'					  "Pickup.Qty " &_
'			  "FROM Booking,mShipLine,mVessel,VslSchedule,Pickup,mPort " &_
'			  "WHERE (" & sWhere & ") AND " &_
'					  "Pickup.VslCode=*Booking.VslCode AND " &_
'					  "Pickup.VoyCtrl=*Booking.VoyCtrl AND " &_
'					  "Pickup.BookNo=*Booking.BookNo AND " &_
'					  "mShipLine.ShipLine=*Booking.ShipLine AND " &_
'					  "mVessel.VslCode=*Booking.VslCode AND " &_
'					  "mPort.PortCode=*Booking.DPort AND " &_
'					  "VslSchedule.VslCode=*Booking.VslCode AND " &_
'					  "VslSchedule.VoyCtrl=*Booking.VoyCtrl"		'D20040223 

        sql = "SELECT Booking.BookNo," &_
					  "Booking.VslCode," &_
					  "Booking.ShipLine," &_
					  "mShipLine.FullName ShipLineName," &_
					  "mVessel.FullName ShipName," &_
					  "VslSchedule.LdVoyage," &_
					  "Booking.DPort," &_
					  "mPort.FullName PortName," &_
					  "Pickup.PickPlace," &_
					  "Pickup.ContSize," &_
					  "Pickup.ContType," &_
					  "Pickup.ContHeight," &_
					  "Pickup.Qty, " &_
					  "Pickup.Material " &_
			  "FROM Booking,mShipLine,mVessel,VslSchedule,Pickup,mPort " &_
			  "WHERE (" & sWhere & ") AND " &_
					  "Pickup.VslCode=*Booking.VslCode AND " &_
					  "Pickup.VoyCtrl=*Booking.VoyCtrl AND " &_
					  "Pickup.BookNo=*Booking.BookNo AND " &_
					  "mShipLine.ShipLine=*Booking.ShipLine AND " &_
					  "mVessel.VslCode=*Booking.VslCode AND " &_
					  "mPort.PortCode=*Booking.DPort AND " &_
					  "VslSchedule.VslCode=*Booking.VslCode AND " &_
					  "VslSchedule.VoyCtrl=*Booking.VoyCtrl"		'I20040223
'2006/03/06 add-s h.matsuda
        If ShipLine<>"" and ShoriMode<>"" Then
            sql = sql & " and Booking.ShipLine='" & ShipLine & "'"
        End If
'2006/03/06 add-e h.matsuda

		rsd.Open sql, conn, 0, 1, 1

		Dim sOutText()
		Dim strOut,bWrite
	    bWrite = 0        '�o�̓��R�[�h����

		Do While Not rsd.EOF
			strOut = Trim(rsd("VslCode")) & ","						' 0:VslCode
			strOut = strOut & Trim(rsd("BookNo")) & ","				' 1:Booking No.

			If IsNull(rsd("ShipLineName")) Then
				strOut = strOut & Trim(rsd("ShipLine")) & ","		' 2:�D��
			Else
				strOut = strOut & Trim(rsd("ShipLineName")) & ","	' 2:�D��
			End If

			If IsNull(rsd("ShipName")) Then
				strOut = strOut & Trim(rsd("VslCode")) & ","		' 3:�D��
			Else
				strOut = strOut & Trim(rsd("ShipName")) & ","		' 3:�D��
			End If

			strOut = strOut & Trim(rsd("LdVoyage")) & ","			' 4:Voyage No.

			If IsNull(rsd("PortName")) Then
				strOut = strOut & Trim(rsd("DPort")) & ","			' 5:�d���`
			Else
				strOut = strOut & Trim(rsd("PortName")) & ","		' 5:�d���`
			End If

			strOut = strOut & Trim(rsd("PickPlace")) & ","			' 6:�s�b�N�ꏊ
			strOut = strOut & Trim(rsd("ContSize")) & ","			' 7:�T�C�Y
			strOut = strOut & Trim(rsd("ContType")) & ","			' 8:�^�C�v
			strOut = strOut & Trim(rsd("ContHeight")) & ","			' 9:����
			strOut = strOut & Trim(rsd("Qty")) & ","				'10:�\��{��

			strOut = strOut & "," & Trim(rsd("Material"))			'12:�ގ�	'I20040223

			ReDim Preserve sOutText(bWrite)
			sOutText(bWrite) = strOut
			bWrite = bWrite + 1

			rsd.MoveNext
		Loop

	    rsd.Close

	    For i=0 To bWrite-1
	        strTmp=Split(sOutText(i),",")

'			sql = "SELECT ExportCont.ContNo FROM ExportCont,Container " &_
'				  "WHERE ExportCont.VslCode='" & strTmp(0) & "'" &_
'				   " AND ExportCont.BookNo='" & strTmp(1) & "'" &_
'				   " AND ExportCont.PickPlace='" & strTmp(6) & "'" &_
'				   " AND Container.VslCode='" & strTmp(0) & "'" &_
'				   " AND ExportCont.VoyCtrl=Container.VoyCtrl" &_
'				   " AND ExportCont.ContNo=Container.ContNo" &_
'				   " AND Container.ContSize='" & strTmp(7) & "'" &_
'				   " AND Container.ContType='" & strTmp(8) & "'" &_
'				   " AND Container.ContHeight='" & strTmp(9) & "'" 		'D20040223

'			sql = "SELECT ExportCont.ContNo FROM ExportCont,Container " &_		'Chenge 2005/03/28
			sql = "SELECT ExportCont.ContNo,Container.TareWeight FROM ExportCont,Container " &_
				  "WHERE ExportCont.VslCode='" & strTmp(0) & "'" &_
				   " AND ExportCont.BookNo='" & strTmp(1) & "'" &_
				   " AND ExportCont.PickPlace='" & strTmp(6) & "'" &_
				   " AND Container.VslCode='" & strTmp(0) & "'" &_
				   " AND ExportCont.VoyCtrl=Container.VoyCtrl" &_
				   " AND ExportCont.ContNo=Container.ContNo" &_
				   " AND Container.ContSize='" & strTmp(7) & "'" &_
				   " AND Container.ContType='" & strTmp(8) & "'" &_
				   " AND Container.ContHeight='" & strTmp(9) & "'" &_
				   " AND Container.Material='" & strTmp(12) & "'"		'I20040223

			rsd.Open sql, conn, 0, 1, 1

			Dim iContNum
			iContNum = 0
			sOutText(i) = sOutText(i) & ","

            Do While Not rsd.EOF
				If iContNum=0 Then
'					sOutText(i) = sOutText(i) & Trim(rsd("ContNo"))
					sOutText(i) = sOutText(i) & Trim(rsd("ContNo")) & "!" & Trim(rsd("TareWeight"))			' Chenge 2005/03/28
				Else
'					sOutText(i) = sOutText(i) & "/" & Trim(rsd("ContNo"))
					sOutText(i) = sOutText(i) & "/" & Trim(rsd("ContNo")) & "!" & Trim(rsd("TareWeight"))	' Chenge 2005/03/28
				End If
				iContNum = iContNum + 1
				rsd.MoveNext
			Loop

	        rsd.Close

     		strTmpIn=Split(sOutText(i),",")
			strTmpIn(11) = iContNum					'11:���o�ϖ{��
			sOutTextTmp = strTmpIn(0)
			For k=1 To UBound(strTmpIn)
				sOutTextTmp = sOutTextTmp & "," & strTmpIn(k)
			Next

'2006/03/06 add-s h.matsuda CSV���ԃt�@�C���̖����ɑD�ЃR�[�h�Ə������[�h��ǉ�
        If ShipLine<>"" and ShoriMode<>"" Then
				sOutTextTmp = sOutTextTmp & "," & ShipLine
				sOutTextTmp = sOutTextTmp & ",ShoriMode=" & ShoriMode
        End If
'2006/03/06 add-e h.matsuda

			ti.WriteLine sOutTextTmp
		Next

        ti.Close
        conn.Close

        ' Temp�t�@�C�������ݒ�
        SetTempFile "EXPORT"

        If bWrite = 0 Then
            ' �Y�����R�[�h�Ȃ��Ƃ�
            bError = true
            strError = "�w������ɊY������Booking No.�͂���܂���ł����B"
            strOption = "���͓��e�̐���:1(���)"
        Else
            strOption = "���͓��e�̐���:0(������)"
        End If

    End If


	Do While InStr(strBookingNoLog,",")>0
		strBookingNoLog = Left(strBookingNoLog,InStr(strBookingNoLog,",")-1) & _
						"/" & Right(strBookingNoLog,Len(strBookingNoLog)-InStr(strBookingNoLog,","))
	Loop

	strOption = strBookingNoLog & "," & strOption

    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "1009","�u�b�L���O���Ɖ�","10", strOption

    If bError Then
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
<!-------------��������Ɖ�G���[���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/shokait.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.07
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
end of comment by seiko-denki 2003.07.17 -->
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>�u�b�L���O���Ɖ�</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
		<BR>
      <table>
        <tr>
          <td>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError
%>
          </td></tr>
      </table>
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
<!-------------�Ɖ�G���[��ʏI���--------------------------->
<%
    DispMenuBarBack "bookentry.asp"
%>
</body>
</html>

<%
    Else
        ' �ꗗ��ʂփ��_�C���N�g
        Response.Redirect "booklist.asp"             '�u�b�L���O���ꗗ
    End If
%>
