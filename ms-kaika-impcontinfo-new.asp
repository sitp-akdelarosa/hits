<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'
	'	�y�A���R���e�i�����́z	���͉��
	'
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

	' �C�݃R�[�h
	sSosin = Trim(Session.Contents("userid"))

	' �V�K�ǉ���(2) or �V�K(1) or �X�V(0)
    bKind = Request.QueryString("kind")
	If bKind=0 Then
		Session.Contents("kind") = 0
	ElseIf bKind=1 Then
		Session.Contents("kind") = 1
	End If
	Dim sVslCode,sVoyCtrl,sUser,sCont,sBL,sTraderCode,sArvTime,sSize,sType,sRemark
	If Not bKind=1 Then
		sVslCode 	= Request.form("vslcode")
		sVoyCtrl 	= Request.form("voyctrl")
		sUser 		= Request.form("user")
		sCont	 	= Request.form("cont")
		sBL		 	= Request.form("bl")
		sTraderCode = Request.form("tradercode")
		sArvTime 	= Request.form("arvtime")
		sSize 		= Request.form("size")
		sType 		= Request.form("type")
		sRemark		= Request.form("remark")
		iLineNo		= Request.form("lineno")
	End If

	' DB����^�C�v�ꗗ���擾	����ǉ��\�� 2002/1/30
    Dim strType()
	ConnectSvr conn, rsd

	sql = "SELECT ContType FROM mContType"
	rsd.Open sql, conn, 0, 1, 1
    TypeLineNo=0
    Do While Not rsd.EOF
        strTemp=Trim(rsd("ContType"))
        ReDim Preserve strType(TypeLineNo)
        strType(TypeLineNo) = strTemp
        TypeLineNo=TypeLineNo+1
        rsd.MoveNext
    Loop
	rsd.Close


%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

	function checkFormValue(){
		if(!checkBlank(getFormValue(0))){ return showAlert("�D��",true); }
		if(!checkBlank(getFormValue(1))){ return showAlert("Voyage No.",true); }
		if(!checkBlank(getFormValue(2))){ return showAlert("�׎�R�[�h",true); }
		if(!checkBlank(getFormValue(3))){ return showAlert("BL No.",true); }
		if(!checkBlank(getFormValue(4))){ return showAlert("�R���e�iNo.",true); }

		if(!checkDate(new getDateValue(getFormValue(6),getFormValue(7),getFormValue(8))) ||
		   !checkTime(new getTimeValue(getFormValue(9),getFormValue(10)))){
			return showAlert("������q�ɓ����w�����",false);
		}

		if(!checkNumberFormat(getFormValue(11))){ return showAlert("�T�C�Y",false); }

		return true;
	}

	function getFormValue(iNum){
		formvalue = window.document.input.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}

	function checkNumberFormat(formvalue){
		if(!((formvalue > 0) || (formvalue <= 0))){ return false; }
		return true;
	}

	function getDateValue(year,mon,day){
		this.year = year;
		this.mon  = mon;
		this.day  = day;
	}

	function getTimeValue(hour,min){
		this.hour = hour;
		this.min  = min;
	}

	function checkDate(gdv){
		if(gdv.year != "" || gdv.mon != "" || gdv.day != ""){
			if( !(gdv.year > 0 || gdv.year <= 0) || gdv.year < 2001 ) { return false; }
			if( !(gdv.mon > 0 || gdv.mon <= 0)   || (gdv.mon < 1 || gdv.mon > 12) ) { return false; }
			if( !(gdv.day > 0 || gdv.day <= 0)   || (gdv.day < 1 || gdv.day > 31) ) { return false; }
			if (gdv.day<=0 || gdv.day>30+((gdv.mon==4||gdv.mon==6||gdv.mon==9||gdv.mon==11)?0:1) || 
				(gdv.mon==2&&gdv.day>28+(((gdv.year%4==0&&gdv.year%100!=0)||gdv.year%400==0)?1:0)) ){ return false; }
		}
		return true;
	}

	function checkTime(gtv){
		if(gtv.hour != "" || gtv.min != ""){
			if( !(gtv.hour > 0 || gtv.hour <= 0) || (gtv.hour < 0 || gtv.hour > 23) ) { return false; }
			if( !(gtv.min > 0 || gtv.min <= 0)   || (gtv.min < 0 || gtv.min > 59) )   { return false; }
		}
		return true;
	}

	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "�������͂ł��B");
		} else {
			window.alert(strAlert + "���s���ł��B");
		}
		return false;
	}

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
          <td rowspan=2><img src="gif/kaika6t.gif" width="506" height="73"></td>
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
<% If bKind<>0 Then %>
          <td nowrap><b>�V�K������</b></td>
<% Else %>
          <td nowrap><b>�X�V������</b></td>
<% End If %>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>
      <table>
        <tr>
          <td nowrap align=center>
				�A���R���e�i�ɂ��āA�ȉ��̍��ڂ���͂��đ��M�������ĉ������B
            <form method=post name="input" action="ms-kaika-impcontinfo-exec.asp">
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�D��(�R�[���T�C��)</b></font>
                  </td>
                  <td nowrap>
<% If Not bKind=0 Then %>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=vslcode value="<%=sVslCode%>" size=9 maxlength=7>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                    
<%  Else %>
                    <%=sVslCode%>
					<input type=hidden name="vslcode" value="<%=sVslCode%>">
<%  End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Voyage No.</b></font>
                  </td>
                  <td nowrap>
<% If Not bKind=0 Then %>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=voyctrl value="<%=sVoyCtrl%>" size=14 maxlength=12>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                    
<%  Else %>
                    <%=sVoyCtrl%>
					<input type=hidden name="voyctrl" value="<%=sVoyCtrl%>">
<%  End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�C�݃R�[�h</b></font>
                  </td>
                  <td nowrap>
					<%=sSosin%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle> <font color="#FFFFFF"><b>�׎�R�[�h</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=user value="<%=sUser%>" size=7 maxlength=5>
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
                    <font color="#FFFFFF"><b>BL No.</b></font>
                  </td>
                  <td nowrap>
<% If Not bKind=0 Then %>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=bl value="<%=sBL%>" size=22 maxlength=20>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
<%  Else %>
                    <%=sBL%>
					<input type=hidden name="bl" value="<%=sBL%>">
<%  End If %>
                    
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�R���e�iNo.</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=cont value="<%=sCont%>" size=14 maxlength=12>
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
                    <font color="#FFFFFF"><b>�w�藤�^�Ǝ҃R�[�h</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=tradercode value="<%=sTraderCode%>" size=5 maxlength=3>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>������q�ɓ����w�����</b></font>
                  </td>
                  <td nowrap>
<%
	Dim eyear,emon,eday,ehour,emin
	If bKind=0 Then
		If sArvTime<>"" Then
			eyear= Left(sArvTime,4)
			emon = Mid(sArvTime,6,2) 
			eday = Mid(sArvTime,9,2) 
			ehour= Mid(sArvTime,12,2) 
			emin = Mid(sArvTime,15,2) 
		End If
	End If
%>

                    <input type=text name=emparvtime_year value="<%=eyear%>" size=5 maxlength=4>�N
                    <input type=text name=emparvtime_mon value="<%=emon%>" size=3 maxlength=2>��
                    <input type=text name=emparvtime_day value="<%=eday%>" size=3 maxlength=2>��&nbsp;&nbsp;
                    <input type=text name=emparvtime_hour value="<%=ehour%>" size=3 maxlength=2>��
                    <input type=text name=emparvtime_min value="<%=emin%>" size=3 maxlength=2>��
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N2��25�� 15��30��</font>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p���l ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�T�C�Y</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=size value="<%=sSize%>" size=4 maxlength=2>
							<font size=-1>�i��j 20 , 40 , 45</font>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p���l ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�^�C�v</b></font>
                  </td>
                  <td nowrap>
					<select name=type>
<% If sType="" Then %>
						<option value="" selected>
<% Else %>
						<option value="">
<% End If %>

<%
	For i = 1 to TypeLineNo
		If bKind<>1 And sType=strType(i-1) Then
%>
						<option value="<%=strType(i-1)%>" selected><%=strType(i-1)%>
<%
		Else 
%>
						<option value="<%=strType(i-1)%>"><%=strType(i-1)%>
<%
		End If
	Next 
%>
					</select>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�q�ɗ��́i������͂���j</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=remark value="<%=sRemark%>" size=7 maxlength=5>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���{����͉� ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>


              </table>
              <br>
				<input type=hidden name="kind" value="<%=bKind%>">
				<input type=hidden name="lineno" value="<%=iLineNo%>">
                <input type=submit name="send" value=" ��  �M " onClick="return checkFormValue()">

<% If bKind=1 Then %>
                <input type=submit name="stop" value=" �I  �� ">
<% ElseIf bKind=2 Then %>
                <input type=button value=" �I  �� " onClick="JavaScript:window.location.href='ms-kaika-impcontinfo-list.asp'">
<% Else %>
                <input type=button value=" �I  �� " onClick="JavaScript:window.history.back()">
<%  End If %>

<% If bKind=0 Then %>
                <input type=submit name="del" value=" ��  �� ">
<%  End If %>


				</center>
              </center>
            </form>
<%
            If bError Then
                ' �G���[���b�Z�[�W�̕\��
                DispErrorMessage strError
            End If
%>
          </td>
        </tr>
      </table>
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
    DispMenuBarBack "ms-kaika-impcontinfo.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log�쐬
	If bKind=0 Then
		'�X�V
	    WriteLog fs, "4110","�C�ݓ��͗A���R���e�i���-������", "02", sCont & ","
	Else
		'�V�K
	    WriteLog fs, "4110","�C�ݓ��͗A���R���e�i���-������", "01", ","
	End If
%>
