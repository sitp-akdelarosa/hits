<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'	�y�C�ݓ��́z	���͉��
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "pickselect.asp"

	' �C�݃R�[�h
	sSosin = Trim(Session.Contents("userid"))
    strUserKind=Session.Contents("userkind")

	' �V�K�ǉ���(2) or �V�K(1) or �X�V(0)
    bKind = Request.QueryString("kind")
	If bKind=0 Then
		Session.Contents("kind") = 0
	ElseIf bKind=1 Then
		Session.Contents("kind") = 1
	End If

	Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sTraderCode,sArvTime,sSize,sType,sHeight,sPick,sRemark,sRecDate,sOpeCode
	If Not bKind=1 Then
		sUser 		= Request.form("user")
		sUserNo 	= Request.form("userno")
		sVslCode 	= Request.form("vslcode")
		sVoyCtrl 	= Request.form("voyctrl")
		sBooking 	= Request.form("booking")
		sTraderCode = Request.form("tradercode")
		sArvTime 	= Request.form("arvtime")
		sSize 		= Request.form("size")
		sType 		= Request.form("type")
		sHeight 	= Request.form("height")
		sPick 		= Request.form("pickplace")
		sRemark		= Request.form("remark")
		sRecDate	= Request.form("receivedate")
		sOpeCode	= Request.form("opecode")
		sForwarder	= Request.form("forwarder")
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
		if(!checkBlank(getFormValue(3))){ return showAlert("�׎�R�[�h",true); }
		if(!checkBlank(getFormValue(4))){ return showAlert("�׎�Ǘ��ԍ�",true); }
		if(!checkBlank(getFormValue(5))){ return showAlert("Booking No.",true); }
<% If strUserKind="�C��" Then %>
		if(!checkBlank(getFormValue(6))){ return showAlert("�`�^�R�[�h",true); }
<% End If %>

<%
	If strUserKind="�C��" Then
		iJSNum1 = "8"
		iJSNum2 = "13"
	Else
		iJSNum1 = "6"
		iJSNum2 = "11"
	End If
%>

		if(!checkDate(new getDateValue(getFormValue(<%=iJSNum1%>),getFormValue(<%=iJSNum1+1%>),getFormValue(<%=iJSNum1+2%>))) ||
		   !checkTime(new getTimeValue(getFormValue(<%=iJSNum1+3%>),getFormValue(<%=iJSNum1+4%>)))){
			return showAlert("��R���q�ɓ����w�����",false);
		}

<% If strUserKind="�C��" Then %>
		if(!checkDate(new getDateValue(getFormValue(17),getFormValue(18),getFormValue(19)))){
			return showAlert("��R�����o�w���",false);
		}
<% End If %>

		if(!checkNumberFormat(getFormValue(<%=iJSNum2%>))){ return showAlert("�T�C�Y",false); }
		if(!checkNumberFormat(getFormValue(<%=iJSNum2+2%>))){ return showAlert("����",false); }

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
				�A�o�ݕ��ɂ��āA�ȉ��̍��ڂ���͂��đ��M�������ĉ������B
            <form method=post name="input" action="pickexp-exec.asp">
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�D��(�R�[���T�C��)</b></font>
                  </td>
                  <td nowrap>
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
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Voyage No.</b></font>
                  </td>
                  <td nowrap>
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
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�C�݃R�[�h</b></font>
                  </td>
                  <td nowrap>
<% If strUserKind="�C��" Then %>
					<%=sSosin%>
					<input type=hidden name="forwarder" value="<%=sSosin%>">
<% Else %>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
						  <input type=text name=forwarder value="<%=sForwarder%>" size=7 maxlength=5>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle> <font color="#FFFFFF"><b>�׎�R�[�h</b></font></td>
                  <td nowrap>
<% If Not bKind=0 And strUserKind="�C��" Then %>
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
<%  ElseIf strUserKind="�׎�" Then %>
                    <%=sSosin%>
					<input type=hidden name="user" value="<%=sSosin%>">
<%  Else %>
                    <%=sUser%>
					<input type=hidden name="user" value="<%=sUser%>">
<%  End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�׎�Ǘ��ԍ�</b></font>
                  </td>
                  <td nowrap>
<% If Not bKind=0 Then %>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=userno value="<%=sUserNo%>" size=12 maxlength=10>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
<%  Else %>
                    <%=sUserNo%>
					<input type=hidden name="userno" value="<%=sUserNo%>">
<%  End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Booking No.</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=booking value="<%=sBooking%>" size=22 maxlength=20>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

<% If strUserKind="�C��" Then %>
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�`�^�R�[�h</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=opecode value="<%=sOpeCode%>" size=5 maxlength=3>
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
<% End If %>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>��R���q�ɓ����w�����</b></font>
                  </td>
                  <td nowrap>
<%
	Dim eyear,emon,eday,ehour,emin,cyyear,cymon,cyday,recyear,recmon,recday
	If bKind=0 Then
		If sArvTime<>"" Then
			eyear= Left(sArvTime,4)
			emon = Mid(sArvTime,6,2) 
			eday = Mid(sArvTime,9,2) 
			ehour= Mid(sArvTime,12,2) 
			emin = Mid(sArvTime,15,2) 
		End If
		If sRecDate<>"" Then
			recyear= Left(sRecDate,4)
			recmon = Mid(sRecDate,6,2) 
			recday = Mid(sRecDate,9,2) 
		End If
	End If
%>

                    <input type=text name=emparvtime_year value="<%=eyear%>" size=5 maxlength=4>�N
                    <input type=text name=emparvtime_mon value="<%=emon%>" size=3 maxlength=2>��
                    <input type=text name=emparvtime_day value="<%=eday%>" size=3 maxlength=2>��&nbsp;&nbsp;
                    <input type=text name=emparvtime_hour value="<%=ehour%>" size=3 maxlength=2>��
                    <input type=text name=emparvtime_min value="<%=emin%>" size=3 maxlength=2>��<BR>
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
                    <font color="#FFFFFF"><b>����</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=height value="<%=sHeight%>" size=4 maxlength=2>
							<font size=-1>�i��j 86 , 96</font>
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
                    <font color="#FFFFFF"><b>��R���s�b�N�ꏊ</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							<input type=text name=pickplace value="<%=sPick%>" size=22 maxlength=20>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���{����͉� ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>

<% If strUserKind="�C��" Then %>
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>��R�����o�w���</b></font>
                  </td>
                  <td nowrap>
                    <input type=text name=recdate_year value="<%=recyear%>" size=5 maxlength=4>�N
                    <input type=text name=recdate_mon value="<%=recmon%>" size=3 maxlength=2>��
                    <input type=text name=recdate_day value="<%=recday%>" size=3 maxlength=2>��<BR>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=200>
							&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N2��25��</font>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p���l ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<% End If %>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�q�ɗ���</b></font>
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
                <input type=submit name="stop" value=" ��  �~ ">
<% ElseIf bKind=2 Then %>
                <input type=button value=" ��  �~ " onClick="JavaScript:window.location.href='pickexp-list.asp'">
<% Else %>
                <input type=button value=" ��  �~ " onClick="JavaScript:window.history.back()">
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
    DispMenuBarBack "pickexpinfo.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	If strUserKind="�C��" Then
		iNum = "a111"
	Else
		iNum = "a114"
	End If

	If bKind=0 Then
		'�X�V
	    WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-" & strUserKind & "�p�˗�����", "02", sBooking & ","
	Else
		'�V�K
	    WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-" & strUserKind & "�p�˗�����", "01", sBooking & ","
	End If
%>
