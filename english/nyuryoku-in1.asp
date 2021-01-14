<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-in1.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' ���̓t���O�̃N���A
    bInput = true

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' �w������̎擾
    Dim strCallSign
    Dim strVoyage
    strCallSign = UCase(Trim(Request.QueryString("callsign")))
    strVoyage = UCase(Trim(Request.QueryString("voyage")))

	If InStr(strVoyage,",")<>0 Then
        ' �J���}���͎��G���[
        bError = true
        strError = "Voyage No.�ɔ��p�J���}�͎g�p���Ȃ��ŉ������B"
        strOption = strCallSign & "/" & strVoyage & "," & "���͓��e�̐���:1(���)"
    End If

    If strCallSign="" Or strVoyage="" Then
        If strCallSign<>"" Or strVoyage<>"" Then
            ' ���͂��Е������̂Ƃ� �G���[���b�Z�[�W��\��
            bError = true
            strError = "���͂��Ԉ���Ă��܂��B"
            strOption = strCallSign & "/" & strVoyage & "," & "���͓��e�̐���:1(���)"
        Else
            bInput = false
        End If
    End If

    If bInput And Not bError Then
        ' ���̓R�[���T�C���̃`�F�b�N
        ConnectSvr conn, rsd
        sql = "SELECT FullName, ShipLine FROM mVessel WHERE VslCode='" & strCallSign & "'"
        'SQL�𔭍s���đD���}�X�^�[������
        rsd.Open sql, conn, 0, 1, 1
        If Not rsd.EOF Then
            strVesselName = Trim(rsd("FullName"))
            strShipLine = Trim(rsd("ShipLine"))
            strOption = strCallSign & "/" & strVoyage & "," & "���͓��e�̐���:0(������)"
        Else
            ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
            bError = true
            strError = "�R�[���T�C�����Ԉ���Ă��܂��B"
            strOption = strCallSign & "/" & strVoyage & "," & "���͓��e�̐���:1(���)"
        End If
        rsd.Close
        If Not bError Then
            ' �D�Ж��̎擾
            sql = "SELECT FullName FROM mShipLine WHERE ShipLine='" & strShipLine & "'"
            'SQL�𔭍s���đD���}�X�^�[������
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                strShipLineName = Trim(rsd("FullName"))
            End If
            rsd.Close

            Dim strPortData()

            ' SQL�𔭍s���Ė{�D���Â�����
            sql = "SELECT VoyCtrl, DsVoyage, LdVoyage FROM VslSchedule " & _
                  "WHERE VslCode='" & strCallSign & "' And " & _
                  "(DsVoyage='" & strVoyage & "' Or LdVoyage='" & strVoyage & "')"
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                iVoyCtrl = rsd("VoyCtrl")
                strDsVoyage = Trim(rsd("DsVoyage"))
                strLdVoyage = Trim(rsd("LdVoyage"))
                rsd.Close
                ' �{�D���Ï�񃌃R�[�h�̍쐻(���ݗ\�莞���������Ă�����̂��ɓǂ�)
                strVslSchdule = strShipLine & "," & strShipLineName & "," & strCallSign & "," & strVesselName & "," & _
								iVoyCtrl & "," & strDsVoyage & "," & strLdVoyage
                ' SQL�𔭍s���Ė{�D��`�n������(��������̗v�]�ŁA��`�n���� 2002/02/27)
'               sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong, mPort.FullName " & _
'                     "FROM VslPort, mPort WHERE VslPort.VslCode='" & strCallSign & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
'                     " And mPort.PortCode=*VslPort.PortCode And VslPort.ETA is NOT Null ORDER BY VslPort.ETA "
                sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong, mPort.FullName " & _
                      "FROM VslPort, mPort WHERE VslPort.VslCode='" & strCallSign & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
                      " And mPort.PortCode=*VslPort.PortCode ORDER BY VslPort.CallSeq "
                rsd.Open sql, conn, 0, 1, 1
                iRecCount=0
				iSeq = 1
                Do While Not rsd.EOF
                    ' ��`�n��񃌃R�[�h�̍쐻
                    strRec = Trim(rsd("PortCode")) & "," & Trim(rsd("FullName")) & "," & _
							 DispDateTime(rsd("ETA"),0) & "," & DispDateTime(rsd("TA"),0) & ","  & _
							 DispDateTime(rsd("ETD"),0) & "," & DispDateTime(rsd("TD"),0) & ","  & _
							 DispDateTime(rsd("ETALong"),0) & "," & DispDateTime(rsd("ETDLong"),0)
                    ReDim Preserve strPortData(iRecCount)
                    strPortData(iRecCount) = strRec
                    iRecCount=iRecCount + 1
					iSeq = iSeq + 1
                    rsd.MoveNext
                Loop
' 				rsd.Close
'01/12/22 ADD (��������̗v�]�ŁA��`�n���ɂ������ߕs�v�� 2002/02/27)
'                ' �{�D���Ï�񃌃R�[�h�̍쐻(���ݗ\�莞���������Ă��Ȃ����̂�ǂ�)
'                ' SQL�𔭍s���Ė{�D��`�n������
'                sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong, mPort.FullName " & _
'                      "FROM VslPort, mPort WHERE VslPort.VslCode='" & strCallSign & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
'                      " And mPort.PortCode=*VslPort.PortCode And VslPort.ETA is Null"
'                rsd.Open sql, conn, 0, 1, 1
'                Do While Not rsd.EOF
'                    ' ��`�n��񃌃R�[�h�̍쐻
'                    strRec = Trim(rsd("PortCode")) & "," & Trim(rsd("FullName")) & "," & _
'							 DispDateTime(rsd("ETA"),0) & "," & DispDateTime(rsd("TA"),0) & ","  & _
'							 DispDateTime(rsd("ETD"),0) & "," & DispDateTime(rsd("TD"),0) & ","  & _
'							 DispDateTime(rsd("ETALong"),0) & "," & DispDateTime(rsd("ETDLong"),0)
'                    ReDim Preserve strPortData(iRecCount)
'                    strPortData(iRecCount) = strRec
'                    iRecCount=iRecCount + 1
'					iSeq = iSeq + 1
'                    rsd.MoveNext
'                Loop

            Else
                ' �{�D���Ï�񃌃R�[�h�̍쐻
                strVslSchdule = strShipLine & "," & strShipLineName & "," & strCallSign & "," & strVesselName & ",,"  & _
								strVoyage & "," & strVoyage
                iRecCount=0
            End If
            rsd.Close
            ' �����f�[�^���ꎞ�t�@�C���ɏo��
            strFileName="./temp/" & strFileName
            ' �e���|�����t�@�C����Open
            Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

            ti.WriteLine strVslSchdule & "," & iRecCount

            ti.WriteLine iRecCount
            For iCount=0 To iRecCount - 1
                ti.WriteLine strPortData(iCount)
            Next

            ti.Close
        End If
        conn.Close
    End If

    If bError Or Not bInput Then
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
<!-------------��������D�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/nyuryoku-s.gif" width="506" height="73"></td>
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

<table border=0><tr><td align=left>
  <table>
                  <tr>
                    
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                
                  <td nowrap><b>�{�D���Ó���</b></td>
                   <td><img src="gif/hr.gif"></td>
 </tr>
</table>
 <center>             
	  <table>
	   <tr>
	                <td nowrap>�ΏۂƂȂ�{�D�Ɋւ��鉺�L�̏�����͂̏�A<BR>���M�{�^�����N���b�N���ĉ������B</td>
          </tr>
		</table>
            	<FORM NAME="con" action="nyuryoku-in1.asp">
                  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                    <tr> 
                      <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF"> 
                        �R�[���T�C��</font></b></td>
                      <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=120>
								<input type=text name=callsign value="<%=strCallSign%>" size=10 maxlength=7>
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
								<font size=1 color="#2288ff">[ ���p�p�� ]</font>
							</td>
						  </tr>
						</table>
				                    	
                      </td>
                    </tr>
                    <tr> 
                      <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">Voyage 
                        No.</font></b></td>
                      <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=120>
								<input type=text name=voyage value="<%=strVoyage%>" size=12 maxlength=12>
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
								<font size=1 color="#2288ff">[ ���p�p�� ]</font>
							</td>
						  </tr>
						</table>
                      </td>
                    </tr>
                  </table>
                  <br>
			            <INPUT TYPE=submit VALUE=" ��  �M " name="���M">
			<BR>

<%
        ' �G���[���b�Z�[�W�̕\��
        If bError Then
            DispErrorMessage strError
       End If
%>
			<BR>
</center>
                  <table>
                    <tr> 
                      <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                      <td nowrap><b>CSV�t�@�C���]��</b></td>
                      <td><img src="gif/hr.gif"></td>
                    </tr>
                  </table>
<center>
<table border="0" cellspacing="1" cellpadding="2">


          <tr> 
              <td> 
                <p>�����t�@�C���]������ꍇ�͂������N���b�N</p>
              </td>
              <td>�c</td>
              <td><a href="nyuryoku-csv.asp">CSV�t�@�C���]��</a></td>

            </tr>
            <tr> 
              <td>CSV�t�@�C���]���ɂ��Ă̐����͂������N���b�N</td>
              <td>�c</td>
              <td><a href="help07.asp">�w���v</a></td>
            </tr>
          </table>
              </form>
				</center>
</td></tr></table>


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

<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
        If bError Then
		    ' �R�[���T�C���^���q����
		    WriteLog fs, "3001","�D�Ё^�^�[�~�i������","10", strOption
		Else
		    ' �R�[���T�C���^���q����
		    WriteLog fs, "3001","�D�Ё^�^�[�~�i������","00", ","
		End If
    Else
	    ' �R�[���T�C���^���q����
	    WriteLog fs, "3001","�D�Ё^�^�[�~�i������","10", strOption

        ' �{�D���Õ\����ʂփ��_�C���N�g
        Response.Redirect "nyuryoku-port.asp"    '�{�D���Õ\�����
    End If
%>
