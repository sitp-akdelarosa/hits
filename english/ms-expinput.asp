<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' �w������̎擾
    Dim strKind       '���͎��(1=�͎���,2=��������)
    Dim iLine         '���͍s
    Dim strRequest    '�߂��
    strKind=Trim(Request.QueryString("kind"))
    iLine=CInt(Trim(Request.QueryString("line")))
    strRequest=Trim(Request.QueryString("request"))

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "index.asp"             '���j���[��ʂ�
        Response.End
    End If
    strFileName="./temp/" & strFileName

	Dim iNum
    ' �A�o���^������
    If strKind="1" Then
		iNum = ""
       strTitle="(�A�o)��R���e�i�q�ɓ�������"
    Else
		iNum = "1107"
       strTitle="(�A�o)�o���j���O��������"
    End If
    WriteLog fs, iNum,"�A�o�R���e�i�Ɖ�-�o���j���O������������","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �ڍו\���s�̃f�[�^�̎擾
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLine=LineNo Then
           Exit Do
        End If
    Loop
    ti.Close

    Session.Contents("editkind")=strKind         ' ���͎�ނ��L��
    Session.Contents("editline")=iLine           ' �ҏW�s���L��
    Session.Contents("request")=strRequest       ' �߂��ʂ��L��
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
    If strKind="1" Then
%>
function ClickSend() {
    return (ChkSend("�͂�����",
                    document.con.Year.value, 
                    document.con.Month.value, 
                    document.con.Day.value, 
                    document.con.Hour.value, 
                    document.con.Min.value));
}
<%
    Else
%>
function ClickSend() {
    return (ChkSend("��������",
                    document.con.Year.value, 
                    document.con.Month.value, 
                    document.con.Day.value, 
                    document.con.Hour.value, 
                    document.con.Min.value));
}
<%
    End If
%>
// ���̓`�F�b�N
function ChkSend(sMes, sYear, sMonth, sDay, sHour, sMin ) {
    if (sYear == "" ||  sMonth == "" || sDay == "" || sHour == "" || sMin == "") {
        window.alert(sMes+"�������͂ł��B");
        return false;
    }
    if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* �N�̃`�F�b�N */
        window.alert(sMes+"�̔N�̓��͂��s���ł��B");
        return false;
    }
    if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* ���̃`�F�b�N */
        window.alert(sMes+"�̌��̓��͂��s���ł��B");
        return false;
    }
    if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {		/* ���̃`�F�b�N */
        window.alert(sMes+"�̓��̓��͂��s���ł��B");
        return false;
    }
    if (!(sHour > 0 || sHour <= 0)|| sHour < 0 || sHour > 24  ) {		/* ���̃`�F�b�N */
        window.alert(sMes+"�̎��̓��͂��s���ł��B");
        return false;
    }
    if (!(sMin > 0 || sMin <= 0)|| sMin < 0 || sMin > 59  ) {		/* ���̃`�F�b�N */
        window.alert(sMes+"�̕��̓��͂��s���ł��B");
        return false;
    }
    if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
       (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
        window.alert(sMes+"�̓��̓��͂��s���ł��B");
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
          <td rowspan=2><img src="gif/exprikuun.gif" width="506" height="73"></td>
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
				<%=strRoute%> &gt; ��������
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
<%
    Response.Write strTitle
%>
            ����</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr>
          <td>���L�̍��ڂ���͂̏�A���M�{�^�����N���b�N���ĉ������B</td>
        </tr>
      </table>
      <FORM NAME="con" METHOD="post" action="ms-expinput-syori.asp" onSubmit="return ClickSend()">
		<input type=hidden name=title value="<%=strTitle%>">
        <table border=0 cellpadding=0 bordercolor="#999999">
          <tr> 
            <td align="center"> 
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                    �׎喼��</font></b></td>
                  <td bgcolor="#FFFFFF"> 
<% ' �׎��� - ����
    Response.Write anyTmp(7)
%>
                  </td>
                </tr>
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                     �׎�Ǘ��ԍ�</font></b></td>
                  <td bgcolor="#FFFFFF"> 
<% ' �׎��� - �Ǘ��ԍ�
    Response.Write anyTmp(14)
%>
                  </td>
                </tr>
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
<%
    If strKind="1" Then
        Response.Write "�͂�����"
    Else
        Response.Write "��������"
    End If
%>
                    </font></b></td>
                  <td> 
<%
    If strKind="1" Then
        strTemp=anyTmp(47)
    Else
        strTemp=anyTmp(48)
    End If
    If strTemp="" Then
        strTemp=DispDateTime(Now,0)
    End If
    Response.Write "<input type=text name='Year' value='" & Left(strTemp,4) & "' size=4 maxlength='4'>�N"
    Response.Write "<input type=text name='Month' value='" & Mid(strTemp,6,2) & "' size=2 maxlength='2'>��"
    Response.Write "<input type=text name='Day' value='" & Mid(strTemp,9,2)  & "' size=2 maxlength='2'>���@"
    Response.Write "<input type=text name='Hour' value='" & Mid(strTemp,12,2)  & "' size=2 maxlength='2'>��"
    Response.Write "<input type=text name='Min' value='" & Mid(strTemp,15,2)  & "' size=2 maxlength='2'>��"
%>
					<font size=1 color="#2288ff">[���p���l]</font><BR>
					&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N 2�� 25�� 15�� 30��</font>
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
    If strRequest="ms-expdetail.asp" Then
        strTemp=strRequest & "?line=" & iLine
    Else
        strTemp=strRequest
    End If
    DispMenuBarBack strTemp
%>
</body>
</html>
