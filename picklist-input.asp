<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "index.asp"

	Dim iLoginKind,sLoginKind
	sLoginKind = Session.Contents("userkind")

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"             '���j���[��ʂ�
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	Dim bAryFlag
    ' �w������̎擾
    If Request.QueryString("line")<>"" Then
  	    Dim iLine(0)
		iLine(0) = CInt(Trim(Request.QueryString("line")))
		Session.Contents("lineary") = iLine(0)
		bAryFlag = 0
	Else
	    iLine = Split(Session.Contents("lines"),",")
		Session.Contents("lineary") = Session.Contents("lines") '�u���E�U��back�{�^���΍�
		Session.Contents("lines") = ""
		bAryFlag = 1
	End If

	Dim iNum
	iNum = "a109"

	If sLoginKind="�`�^" Then
    	strTitle="��R�����ꏊ�E���o��"
    	WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-��R�����ꏊ�E���o���ύX","02", ","
	Else
    	strTitle="��R�����o��"
    	WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-��R�����ꏊ�E���o���ύX","01", ","
	End If

  ' �ڍו\���s�̃f�[�^�̎擾
  If bAryFlag=0 Then
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLine(0)=LineNo Then
           Exit Do
        End If
    Loop
  End If
    ti.Close

%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function ClickSend() {
		if(!checkBlank(getFormValue(1)) && !checkBlank(getFormValue(2))){ return showAlert("<% If sLoginKind="�`�^" Then %>���ꏊ�y��<% End If %>���o��",true);}
		if(!checkDate(new getDateValue(getFormValue(2),getFormValue(3),getFormValue(4)))){
			return showAlert("��R�����o��",false);
		}
		return true;
	}

	function getFormValue(iNum){
		formvalue = window.document.con.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}

	function getDateValue(year,mon,day){
		this.year = year;
		this.mon  = mon;
		this.day  = day;
	}

	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "�������͂ł��B");
		} else {
			window.alert(strAlert + "���s���ł��B");
		}
		return false;
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

<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<% If sLoginKind="�`�^" Then %>
          <td rowspan=2><img src="gif/pickkot.gif" width="506" height="73"></td>
<% Else %>
          <td rowspan=2><img src="gif/pickrit.gif" width="506" height="73"></td>
<% End If %>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strRoute = Session.Contents("route")
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
<%' If sLoginKind="�`�^" Then %>
				<%'=strRoute%> &gt; ��R�����ꏊ�E���o���ύX
<%' Else %>
				<%'=strRoute%> &gt; ��R�����w����ύX
<%' End If %>
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
          <td nowrap><b>
<% If sLoginKind="�`�^" Then %>
			��R�����ꏊ�E���o���ύX
<% Else %>
			��R�����w����ύX
<% End If %>
			</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr>
<% If sLoginKind="�`�^" Then %>
          <td>�ύX���鍀�ڂ̂ݒl����͂��āA���M�{�^���������ĉ������B</td>
<% Else %>
          <td>�ύX������t����͂��āA���M�{�^���������ĉ������B</td>
<% End If %>
        </tr>
      </table>
      <FORM NAME="con" METHOD="post" action="picklist-input-syori.asp" onSubmit="return ClickSend()">
		<input type=hidden name=title value="<%=strTitle%>">
        <table border=0 cellpadding=0 bordercolor="#999999">
          <tr> 
            <td align="center"> 
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">

<% If sLoginKind="�`�^" Then %>
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                     ��R�����ꏊ</font></b></td>
                  <td bgcolor="#FFFFFF"> 
					<table border=0 cellpadding=0 cellspacing=0 width=100%>
					  <tr>
						<td nowrap>
<% ' ��R�����ꏊ
 	If bAryFlag=0 Then
	    Response.Write "<input type=text name='pickplace' value='" & anyTmp(20) & "' size=22 maxlength=20>"
	Else
	    Response.Write "<input type=text name='pickplace' size=22 maxlength=20>"
	End If
%>
						</td>
						<td nowrap align=right>
							<font size=1 color="#2288ff">[���{����͉�]</font><BR>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<% Else %>
				<input type=hidden name=dammy value="">
<% End If %>

                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
<% If sLoginKind="�`�^" Then %>
                    ��R�����o��</font></b></td>
<% Else %>
                    ��R�����w���</font></b></td>
<% End If %>
                  <td> 
					<table border=0 cellpadding=0 cellspacing=0 width=100%>
					  <tr>
						<td nowrap>
<%
 	If bAryFlag=0 Then
	    strTemp=anyTmp(24)
	    If strTemp="" Then
	        strTemp=DispDateTime(Now,0)
	    End If
	    Response.Write "<input type=text name='pickyear' value='" & Left(strTemp,4) & "' size=4 maxlength='4'>�N"
	    Response.Write "<input type=text name='pickmon' value='" & Mid(strTemp,6,2) & "' size=2 maxlength='2'>��"
	    Response.Write "<input type=text name='pickday' value='" & Mid(strTemp,9,2)  & "' size=2 maxlength='2'>��"
	Else
'	    Response.Write "<input type=text name='pickyear' value='" & Year(Now) & "' size=4 maxlength='4'>�N"
'	    Response.Write "<input type=text name='pickmon' value='" & Month(Now) & "' size=2 maxlength='2'>��"
	    Response.Write "<input type=text name='pickyear' size=4 maxlength='4'>�N"
	    Response.Write "<input type=text name='pickmon' size=2 maxlength='2'>��"
	    Response.Write "<input type=text name='pickday' size=2 maxlength='2'>��"
	End If
%>
					<BR>&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N 2�� 25��</font>
						</td>
						<td width=10></td>
						<td nowrap align=right>
<% If sLoginKind="���^" Then %>
							<font size=1 color="#ff0000">[�K�{����]</font> <BR>
<% End If %>
							<font size=1 color="#2288ff">[���p���l]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
              </table>
              <br>
              <input type=submit value="�@���M�@">
              <input type="button" value="�@���~�@" onclick="history.back()">
            </td>
          </tr>
        </table>
      </form>
      <br>
      <br>
      <br>
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
<!-------------�o�^��ʏI���--------------------------->
<%
	If sLoginKind="�`�^" Then
	    DispMenuBarBack "picklist.asp?kind=4"
	Else
	    DispMenuBarBack "picklist.asp?kind=2"
	End If
%>
</body>
</html>
