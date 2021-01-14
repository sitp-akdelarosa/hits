<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-te.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' ���̓t���O�̃N���A
    bInput = true

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' �Z�b�V�����ϐ�����`�^�R�[�h���擾
    Dim strOpeCode
    strOpeCode = Trim(Session.Contents("userid"))
    strChoice = Trim(Session.Contents("choice"))

    ' �w������̎擾
    Dim strCallSign
    Dim strVoyage
    Dim strBLNo
    Dim strYear
    Dim strMonth
    Dim strDay
    Dim strHour
    Dim strMin
    strCallSign = UCase(Trim(Request.QueryString("callsign")))
    strVoyage = UCase(Trim(Request.QueryString("voyage")))
    strBLNo = UCase(Trim(Request.QueryString("blno")))
    strYear = Trim(Request.QueryString("year"))
    strMonth = Trim(Request.QueryString("month"))
    strDay = Trim(Request.QueryString("day"))
    strHour = Trim(Request.QueryString("hour"))
    strMin = Trim(Request.QueryString("min"))
    If strChoice="bl" Then
	    strInput = strCallSign & "/" & strVoyage & "/" & strBLNo & "/" & strYear & "/" & strMonth & "/" & strDay & " " & strHour & ":" & strMin
	Else
	    strInput = strCallSign & "/" & strVoyage & "/" & strYear & "/" & strMonth & "/" & strDay & " " & strHour & ":" & strMin
	End If

    If strCallSign="" Or strVoyage="" Or strYear="" Or strMonth="" Or strDay="" Then
        If strCallSign<>"" Or strVoyage<>"" Or strYear<>"" Or strMonth<>"" Or strDay<>"" Or strHour<>"" Or strMin<>"" Then
            ' ���͂��ꕔ�����̂Ƃ� �G���[���b�Z�[�W��\��
            bError = true
            strError = "���͂��Ԉ���Ă��܂��B"
            strOption = strInput & "," & "���͓��e�̐���:1(���)"
		ElseIf strChoice="bl" And strBLNo<>"" Then
            ' ���͂��ꕔ�����̂Ƃ� �G���[���b�Z�[�W��\��
            bError = true
            strError = "���͂��Ԉ���Ă��܂��B"
            strOption = strInput & "," & "���͓��e�̐���:1(���)"
        Else
            bInput = false
        End If
    End If

    If bInput And Not bError Then
        ' ���̓R�[���T�C���̃`�F�b�N
        ConnectSvr conn, rsd
        sql = "SELECT FullName FROM mVessel WHERE VslCode='" & strCallSign & "'"
        'SQL�𔭍s���đD���}�X�^�[������
        rsd.Open sql, conn, 0, 1, 1
        If Not rsd.EOF Then
            strVesselName = Trim(rsd("FullName"))
            strOption = strInput & "," & "���͓��e�̐���:0(������)"
        Else
            ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
            bError = true
            strError = "�R�[���T�C�����Ԉ���Ă��܂��B"
            strOption = strInput & "," & "���͓��e�̐���:1(���)"
        End If
        rsd.Close
        If Not bError Then
            ' SQL�𔭍s���Ė{�D���Â�����
            sql = "SELECT VoyCtrl FROM VslSchedule " & _
                  "WHERE VslCode='" & strCallSign & "' And DsVoyage='" & strVoyage & "'"
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                iVoyCtrl = rsd("VoyCtrl")
            Else
                ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
                bError = true
                strError = "Voyage No.���Ԉ���Ă��܂��B"
                strOption = strInput & "," & "���͓��e�̐���:1(���)"
            End If
            rsd.Close
        End If
        If Not bError And strChoice="bl" Then
            ' SQL�𔭍s���ėA��BL������
            sql = "SELECT ShipLine FROM BL " & _
                  "WHERE VslCode='" & strCallSign & "' And VoyCtrl=" & iVoyCtrl & " And BLNo='" & strBLNo & "'"
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
                strShipLine = Trim(rsd("ShipLine"))
            Else
                ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
                bError = true
                strError = "BL�ԍ����Ԉ���Ă��܂��B"
                strOption = strInput & "," & "���͓��e�̐���:1(���)"
            End If
            rsd.Close
        End If
        If Not bError Then
            ' ���͏����\������`�F�b�N

            ' ���̓f�[�^�𑗐M�t�@�C���ɏo��
            strTmp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
                     Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
            strSeqNo = GetDailyTransNo()
            strFileName = Mid(strTmp,5,4) & strSeqNo & ".snd"

            strFileName="./send/" & strFileName
            ' �e���|�����t�@�C����Open
            Set ti = fs.OpenTextFile(Server.MapPath(strFileName),2,True)

			strMonth = DateFormat(strMonth)
			strDay = DateFormat(strDay)
			strHour = DateFormat(strHour)
            If strHour="" Then
                strHour="23"
            End If
			strMin = DateFormat(strMin)
            If strMin="" Then
                strMin="59"
            End If
			If strChoice="bl" Then
	            ti.WriteLine strSeqNo & ",IM15,R," & strTmp & ",Web - " & Session.Contents("userid") & ",," & _
	                         strCallSign & "," & strVoyage & "," & strBLNo & "," & strYear & strMonth & strDay & strHour & strMin
			Else
	            ti.WriteLine strSeqNo & ",IM15,R," & strTmp & ",Web - " & Session.Contents("userid") & ",," & _
	                         strCallSign & "," & strVoyage & ",," & strYear & strMonth & strDay & strHour & strMin
			End If

            ti.Close

            bError = true
            strError = "����ɍX�V����܂����B"
        End If
        conn.Close
    End If

    ' �����m�F�\�莞������(BL�P��)
    If strChoice="bl" Then
		If bInput Then
	        WriteLog fs, "5002", "�^�[�~�i������-�����m�F�\�莞������(BL�P��)", "10", strOption
		Else
	        WriteLog fs, "5002", "�^�[�~�i������-�����m�F�\�莞������(BL�P��)", "00", ","
		End If
    Else
		If bInput Then
	        WriteLog fs, "5004", "�^�[�~�i������-�����m�F�\�莞������(�{�D�P��)", "10", strOption
		Else
	        WriteLog fs, "5004", "�^�[�~�i������-�����m�F�\�莞������(�{�D�P��)", "00", ","
		End If
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

	function ParamClear(){
		window.document.con.callsign.value  = "";
		window.document.con.voyage.value 	= "";
		window.document.con.year.value 		= "";
		window.document.con.month.value 	= "";
		window.document.con.day.value 		= "";
		window.document.con.hour.value 		= "";
		window.document.con.min.value 		= "";
<% If strChoice="bl" Then %>
		window.document.con.blno.value 		= "";
<%  End If %>
	}

	function DateCheck(){
		bErr  = true;
		mYear =	window.document.con.year.value;
		mMon  =	window.document.con.month.value;
		mDay  =	window.document.con.day.value;
		mHour = window.document.con.hour.value;
		mMin  =	window.document.con.min.value;

		if(!(mYear > 0 || mYear <= 0)|| mYear > 2100 || mYear < 1990){
            sName="�N";
            bErr = false;
        }
		if(!(mMon > 0 || mMon <= 0)|| mMon  > 12   || mMon  < 1){
            sName="��";
            bErr = false;
        }
		if(!(mDay > 0 || mDay <= 0)|| mDay  > 31   || mDay  < 1){
            sName="��";
            bErr = false;
        }
		if(!(mHour > 0 || mHour <= 0)|| mHour > 23   || mHour < 0){
            sName="��";
            bErr = false;
        }
		if(!(mMin > 0 || mMin <= 0)|| mMin  > 59   || mMin  < 0){
            sName="��";
            bErr = false;
        }

        if (mDay>30+((mMon==4||mMon==6||mMon==9||mMon==11)?0:1) || 
           (mMon==2 && mDay>28+(((mYear%4==0 && mYear%100!=0) || mYear%400==0)?1:0)) ){
            sName="��";
            bErr = false;
		}

		if(!bErr) window.alert("�����\�������" + sName + "�̓��͂��s���ł��B");
		return bErr;
	}

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------������������m�F�\�莞�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/terminal2t.gif" width="506" height="73"></td>
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
<table border=0><tr><td>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>
<% If strChoice="bl" Then %>
	�ʔ����m�F�\�莞������( BL�P�� )
<% Else %>
	�ꊇ�����m�F�\�莞������( �{�D�P�� )
<%  End If %>
			</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
<center>
      <table>
        <tr>
          <td nowrap align=left>���L�̍��ڂ���͂̏�A�w���M�x�{�^�����N���b�N���ĉ������B<BR>
			�e��ʊւ̎�������͂��܂��B
		  </td>
        </tr>
      </table>
      <FORM NAME="con" action="nyuryoku-te1.asp" onSubmit="return DateCheck()">
        <table border=0 cellpadding=0>
          <tr>
            <td align="center"> 
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF"> 
                    �R�[���T�C��</font></b></td>
                  <td>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=250>
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
						<td width=250>
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
                <tr>
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">�����m�F�\�����
                  </font></b></td>
                  <td>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=250>
<% 
	If strYear="" Then
		strYear = Year(Now)
	End If
	If strMonth="" Then
		strMonth = Month(Now)
	End If

If strChoice<>"bl" Then
	If strDay="" Then
		strDay = Day(Now)
	End If
End If

	If strHour="" Then
		strHour = "00"
	End If
	If strMin="" Then
		strMin = "00"
	End If
%>
		                    <input type=text name=year value="<%=strYear%>" size=4 maxlength=4>
		                    �N
		                    <input type=text name=month value="<%=strMonth%>" size=2 maxlength=2>
		                    ��
		                    <input type=text name=day value="<%=strDay%>" size=2 maxlength=2>
		                    ��
		                    <input type=text name=hour value="<%=strHour%>" size=2 maxlength=2>
		                    ��
		                    <input type=text name=min value="<%=strMin%>" size=2 maxlength=2>
		                    ��
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p���l ]</font>
						</td>
					  </tr>
					</table>
					&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N2��25�� 15��30��</font>
                  </td>
                </tr>
<% If strChoice="bl" Then %>
                <tr>
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">BL 
                    No.</font></b></td>
                  <td>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=250>
							<input type=text name=blno value="<%=strBLNo%>" size=20 maxlength=20>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>
<%  End If %>
              </table>
              <br>
              <INPUT TYPE=submit VALUE="�@���M�@">
              <INPUT TYPE=button VALUE="�@�N���A�@" onClick="ParamClear()">
            </td>
          </tr>
         </table>
<%
    ' �G���[���b�Z�[�W�̕\��
    If bError Then
        If strError="����ɍX�V����܂����B" Then
            DispInformationMessage strError
        Else
            DispErrorMessage strError
        End If
    End If
%>

</center>
         <BR><BR>
         <table>
           <tr> 
             <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
             <td nowrap><b>�t�@�C���]��</b></td>
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
              <td><a href="nyuryoku-tmnl-csv.asp">CSV�t�@�C���]��</a></td>
            </tr>
            <tr> 
              <td>CSV�t�@�C���]���ɂ��Ă̐����͂������N���b�N</td>
              <td>�c</td>
<% If strChoice="bl" Then %>
              <td><a href="help11.asp">�w���v</a></td>
<%  Else  %>
              <td><a href="help12.asp">�w���v</a></td>
<%  End If  %>
              
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
	DispMenuBarBack "nyuryoku-te.asp"
%>
</body>
</html>

<%
    Else
        ' �{�D���Õ\����ʂփ��_�C���N�g
        Response.Redirect "nyuryoku-te.asp"    '�{�D���Õ\�����
    End If
%>
