<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-in1.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' ���̓t���O�̃N���A
    bInput = true

    ' ����ʂ̈����ݒ�
	iLine = 0

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �����w��̂Ȃ��Ƃ�
        strFileName="test.csv"
    End If
    strFileName="./temp/" & strFileName

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)
    Dim sSensya, sSenmei, sJiko, sCallsign

    If Not ti.AtEndOfStream  Then
        anyTmp=Split(ti.ReadLine,",")
	'���̃Z�b�g
		iKensu 	= anyTmp(7)
        sSensya = anyTmp(1)	'�D��
        sSenmei = anyTmp(3)	'�D��
        If anyTmp(5) = anyTmp(6) Then	'���q
	    	sJiko = anyTmp(5)
		Else
	    	sJiko = anyTmp(5) & "/" & anyTmp(6)
		End If
    	sCallsign = anyTmp(2)	'�R�[���T�C��
	End If
    ti.Close

    ' �{�D���Ó��́i�V�K�o�^�j
    WriteLog fs, "3004","�D�Ё^�^�[�~�i������-�{�D���Ó���","01", ","
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {
	if (ChkSend("���ݗ\�莞��", document.con.port.value ,
				document.con.ayear.value, 
				document.con.amonth.value,
				document.con.aday.value,
				document.con.ahour.value,
				document.con.amin.value) && 
           	ChkSend("���݊�������", document.con.port.value ,
				document.con.tyear.value, 
				document.con.tmonth.value,
				document.con.tday.value,
				document.con.thour.value,
				document.con.tmin.value) &&
           	ChkSend("���݊�������", document.con.port.value ,
				document.con.dyear.value, 
				document.con.dmonth.value,
				document.con.dday.value,
				document.con.dhour.value,
				document.con.dmin.value) &&
           	ChkSend("���� Long Schedule", document.con.port.value ,
				document.con.cyear.value, 
				document.con.cmonth.value,
				document.con.cday.value,
//				document.con.chour.value,
//				document.con.cmin.value) &&
				"","") &&
	   		ChkSend("���� Long Schedule", document.con.port.value , 
				document.con.ryear.value, 
				document.con.rmonth.value,
				document.con.rday.value,
//				document.con.rhour.value,
//				document.con.rmin.value)) { 
				"","") ) {
		return true;
	}
	return false;

}

function ChkSend(Name, sPort, sYear, sMonth, sDay, sHour, sTime) {

	if (sPort == "") {	/* �`���̃`�F�b�N */
			window.alert("�`���͕K�{���͂ł��B");
			return false;
	}

	if (Name == "���ݗ\�莞��") {
		if (sYear == "" ||  sMonth == "" || sDay == "") {
			window.alert(Name + "�͕K�{���͂ł��B");
			return false;
		}
	}
	else {
		if (sYear == "" &&  sMonth == "" && sDay == "" &&  sHour == ""  && sTime == "") {
			return true;
		}
	}
	
	if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* �N�̃`�F�b�N */
			window.alert(Name + "�̔N�̓��͂��s���ł��B");
			return false;
	}
	if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* ���̃`�F�b�N */
			window.alert(Name + "�̌��̓��͂��s���ł��B");
			return false;
	}
	if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {			/* ���̃`�F�b�N */
			window.alert(Name + "�̓��̓��͂��s���ł��B");
			return false;
	}

	if (!(sHour > 0 || sHour <= 0)|| sHour < 0 || sHour > 24  ) {		/* ���̃`�F�b�N */
			window.alert(Name + "�̎��̓��͂��s���ł��B");
			return false;
	}

	if (!(sTime > 0 || sTime <= 0)|| sTime < 0 || sTime > 59  ) {		/* ���̃`�F�b�N */
			window.alert(Name + "�̕��̓��͂��s���ł��B");
			return false;
	}

	if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
	   (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
			window.alert(Name + "�̓��̓��͂��s���ł��B");
			return false;
	}
	return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="gif/nyuryoku-s.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
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
                    <td nowrap><b>�{�D���Ó��́@</b></td>
              <td><img src="gif/hr.gif"></td>
            </tr>
          </table>
<br>     

	���t�y�ю��Ԃ́A���p�����œ��͂��ĉ������B
	&nbsp;&nbsp;&nbsp;�i�� �j 2002�N2��25�� 15��30��
<BR><BR>

<table border=0><tr><td>
          <table border=1 cellpadding="3" cellspacing="1">
                <tr> 
                  <td bgcolor="#000099" backgrond="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�D��</b></font></td>
                  <td bgcolor="#FFFFFF" nowrap>
<%
    ' �D�Ж��̕\��
    Response.Write sSensya
%>
                  </td>
                  <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�D��</b></font></td>
                  <td bgcolor="#FFFFFF" nowrap>
<%
    ' �D���̕\��
    Response.Write sSenmei
%>
                  </td>
    </tr>
   </table>
              <table border=1 cellpadding="3" cellspacing="1">
                  <tr>
                    
                  <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>Voyage 
                    No. </b></font></td>
                    <td bgcolor="#FFFFFF" nowrap>
<%
    ' ���q�̕\��
    Response.Write sJiko
%>
                    </td>
                    <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�R�[���T�C��</b></font></td>
                    <td bgcolor="#FFFFFF" nowrap>
<%
    ' �R�[���T�C���̕\��
    Response.Write sCallsign
%>
                    </td>
                  </tr>
              </table>
<br>


<FORM NAME="con" METHOD="post" action="nyuryoku-new-ist.asp" onSubmit="return ClickSend()">
  <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
    <tr> 
      <td bgcolor="#000099" nowrap align=center valign=middle>
        <font color="#FFFFFF"><b>�`��</b></font>
      </td>
      <td nowrap>
		<table border=0 cellpadding=0 cellspacing=0>
		  <tr>
			<td width=250>
				<input type=text name=port value="<%=strPort%>" size=30 maxlength="40">
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
      <td bgcolor="#000099" nowrap align=center valign=middle>
        <font color="#FFFFFF"><b>���ݗ\�莞��</b></font>
      </td>
      <td nowrap>
		<table border=0 cellpadding=0 cellspacing=0>
		  <tr>
			<td width=250>
		        <input type=text name=ayear size=4 maxlength="4">�N
		        <input type=text name=amonth size=2 maxlength="2">��
		        <input type=text name=aday size=2 maxlength="2">���@
		        <input type=text name=ahour size=2 value="00" maxlength="2">��
		        <input type=text name=amin size=2 value="00" maxlength="2">��
			</td>
			<td align=left valign=middle nowrap>
				<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
			</td>
		  </tr>
		</table>
      </td>
    </tr>

    <tr> 
      <td bgcolor="#000099" nowrap align=center valign=middle>
        <font color="#FFFFFF"><b>���݊�������</b></font>
      </td>
      <td nowrap>
        <input type=text name=tyear size=4 maxlength="4">�N
        <input type=text name=tmonth size=2 maxlength="2">��
        <input type=text name=tday size=2 maxlength="2">���@
        <input type=text name=thour size=2 maxlength="2">��
        <input type=text name=tmin size=2 maxlength="2">��
      </td>
    </tr>

    <tr> 
      <td bgcolor="#000099" nowrap align=center valign=middle>
        <font color="#FFFFFF"><b>���݊�������</b></font>
      </td>
      <td nowrap>
        <input type=text name=dyear size=4 maxlength="4">�N
        <input type=text name=dmonth size=2 maxlength="2">��
        <input type=text name=dday size=2 maxlength="2">���@
        <input type=text name=dhour size=2 maxlength="2">��
        <input type=text name=dmin size=2 maxlength="2">��
      </td>
    </tr>

    <tr> 
      <td bgcolor="#000099" nowrap align=center valign=middle>
        <font color="#FFFFFF"><b>���� Long Schedule</b></font>
      </td>
      <td nowrap>
        <input type=text name=cyear size=4 maxlength="4">�N
        <input type=text name=cmonth size=2 maxlength="2">��
        <input type=text name=cday size=2 maxlength="2">���@
      </td>
    </tr>

    <tr> 
      <td bgcolor="#000099" nowrap align=center valign=middle>
        <font color="#FFFFFF"><b>���� Long Schedule</b></font>
      </td>
      <td nowrap>
        <input type=text name=ryear size=4 maxlength="4">�N
        <input type=text name=rmonth size=2 maxlength="2">��
        <input type=text name=rday size=2 maxlength="2">���@
      </td>
    </tr>

  </table>
  <br><br>
  <center>
    <input type=submit value=" ��  �� ">
    <input type="button" value=" �L�����Z��" onclick="history.back()">
  </center>
</form>
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
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-port.asp"
%>
</body>
</html>


