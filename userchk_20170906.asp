<%@Language="VBScript" %>
<%
'for each name in session.contents
'	response.write(name &"===="& session(name) &"<br>")
'next
'response.end
%>

<!--#include file="Common.inc"-->

<%

'�����ʂւ̃����N���Ƀ��O���o�͂���
Sub CheckLinkLog
	Dim iNum,iWrkNum
    Select Case strLinkID
        Case "hits.asp"      strLinkNamne = "�X�g�b�N���[�h���p"
							iNum = "9002"
							iWrkNum = "00"
        Case "gate.asp"      strLinkNamne = "�Q�[�g�ʍs���ԗ\��"        
        Case Else            strLinkNamne = ""
    End Select
    If strLinkNamne<>"" Then
        ' File System Object �̐���
        Set fs=Server.CreateObject("Scripting.FileSystemobject")

        ' �����N�����o��
        WriteLog fs, iNum,strLinkNamne,iWrkNum, ","
    End If
End Sub

%>

<%


'��ʂ̕\��
Function DispLogIn(sError)
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
//2009/02/25 Add-S G.Ariola 
function finit(){
//document.usercheck.Screen_Mode.value="<%=Request.QueryString("Screen_Mode")%>";
document.usercheck.user.value="<%=Request.QueryString("user")%>";
document.usercheck.pass.value="<%=Request.QueryString("pass")%>";
document.usercheck.user.focus();

if ("<%=Request.QueryString("Screen_Mode")%>" == "R" && "<%=bOK%>" == "<%=false%>"){
	OpenReqUsercdWin();
}else if("<%=Request.QueryString("Screen_Mode")%>" == "R" && "<%=bOK%>" == "<%=true%>"){
	OpenUpdUserinfWin();
}else if("<%=bUpdated%>" == "0"){
	OpenConfirmWin();
}

}
function Check(){
  f=document.usercheck;
  userid = f.user.value;
  ret = CheckEisuji(userid);
  if(ret==false){
    alert("��ЃR�[�h�͔��p�p�����œ��͂��Ă��������B");
    return false;
  }
  return true;
}


function CheckEisuji(str){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}

function OpenWindow()
{
	document.usercheck.Screen_Mode.value = "R";
	document.usercheck.submit();
}

function OpenReqUsercdWin()
{
	var CodeWin;
	var w=420;
	var h=220;
	var l=0;
	var t=0;
	if(screen.width){
		l=(screen.width-w)/2;
	}
	if(screen.availWidth){
		l=(screen.availWidth-w)/2;
	}
	if(screen.height){
		t=(screen.height-h)/2;
	}
	if(screen.availHeight){
		t=(screen.availHeight-h)/2;
	}
	
  CodeWin = location.replace("./requsercd.asp","codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);

}

function OpenUpdUserinfWin()
{
	var CodeWin;
	var w=520;
	var h=580;
	var l=0;
	var t=0;
	if(screen.width){
		l=(screen.width-w)/2;
	}
	if(screen.availWidth){
		l=(screen.availWidth-w)/2;
	}
	if(screen.height){
		t=(screen.height-h)/2;
	}
	if(screen.availHeight){
		t=(screen.availHeight-h)/2;
	}
 	
  CodeWin = location.replace("./upduserinf.asp?user=<%=strInputUserID%>","codelist","scrollbars=yes,resizable=no,width="+w+",height="+h+",top="+t+",left="+l);

}

function OpenConfirmWin()
{
	var CodeWin;
	var w=420;
	var h=220;
	var l=0;
	var t=0;
	if(screen.width){
		l=(screen.width-w)/2;
	}
	if(screen.availWidth){
		l=(screen.availWidth-w)/2;
	}
	if(screen.height){
		t=(screen.height-h)/2;
	}
	if(screen.availHeight){
		t=(screen.availHeight-h)/2;
	}
	
  CodeWin = location.replace("./confirm.asp?user=<%=strInputUserID%>&pass=<%=strInputPassWord%>&link=predef/dmi000F.asp","codelist","scrollbars=yes,resizable=no,width="+w+",height="+h+",top="+t+",left="+l);
}
//2009/02/25 Add-E G.Ariola 
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/loginback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">

<!-------------�������烍�O�C�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
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

<!-- 2016/10/11 H.Yoshikawa Add Start -->
	<BR>
	<BR>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align=center><FONT color="red">
          �y�d�v�Ȃ��m�点�z<BR>
           <BR> 2016�N11��7���i���j�ɂj�`�b�b�r�E�gi�s�r�̉��C���s���܂��B<BR>
           ����ɔ������{�v���܂����C�݁E�����l�����y�ї��^���Ǝҗl�����V�V�X�e�����p������̎������g�s�b�N�X�ɃA�b�v�v���܂����B<BR>
           �ڂ����̓g�s�b�N�X�����m�F�������B
          </FONT></td>
        </tr>
	  </table>
<!-- 2016/10/11 H.Yoshikawa Add Start -->
		<table border=0><tr><td height=65></td></tr></table>

        <table border="0" cellspacing="4" cellpadding="0" bgcolor="#ff9933">
          <tr>
           <td>
		  <table border="0" cellspacing="3" cellpadding="4" bgcolor="#ffffff">
          <tr>
           <td>
        <table width="500" border="0" cellspacing="0" cellpadding="5" bgcolor="#FFFFCC">
          <tr>
           <td>
              <table width=100%>
                <tr>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>         <td><img src="gif/1.gif" width="1" height="1"></td>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
		</tr>
		<tr>
		 <td></td>		 
                  <td align="center">

      <table>
        <tr>
          <td nowrap align="Center"> 
            <form name="usercheck" action="userchk.asp" method="put"><INPUT name="Screen_Mode" size="9" readonly tabindex= -1 type= hidden>
			<INPUT name="Skip_Mode" size="2" readonly tabindex= -1 type= hidden>
<!-- 2010.5.19 Mod-S MES Aoyagi -->
            <dl>
                <dd> <!--��ЃR�[�h�ƃp�X���[�h����͂��A�w���M�x�{�^�����N���b�N���Ă������� -->
              </dl>

	  <%="��ЃR�[�h�ƃp�X���[�h����͂��A�w���M�x�{�^�����N���b�N���Ă�������"%>

<!-- 2010.5.19 Mod-E MES Aoyagi -->
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
                <tr> 
                  <td bgcolor="#ff8833" nowrap align=center valign=middle> <font color="#FFFFFF"><b>��ЃR�[�h</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=100>
							<input type=text name=user value="" size=7 maxlength=5>
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
                  <td bgcolor="#ff8833" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�p�X���[�h</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=100>
							<input type=password name=pass size=10 maxlength=8>
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
                <input type="submit" value=" ���@�M " onClick="return Check()"></center>
              </form>

          </td>
        </tr>
      </table>
<%
            If sError<>"" Then
                ' �G���[���b�Z�[�W�̕\��
                DispErrorMessage sError
            End If
%>
	  </td>
	  <td>
	  </td>
	 </tr><!--2009/02/25 Add-S G.Ariola  -->
        <tr>
	  <td  width="30"></td>
	  <!--2011/11/16 M.Marquez Upd-S-->
	  <!--td align="Left"><b><font color="Blue"><%= "�@���삵�Ȃ���Ԃ�" & Session.Timeout & "���ȏ㑱���Ǝ����I�Ƀ��O�A�E�g���܂��B<br> &nbsp;&nbsp;���̏ꍇ�A���萔�ł����ēx���O�C���̂��������p���������B"%></Font></b></td-->
	  <td align="Left"><b><font color="Blue"><%= "&nbsp;&nbsp;�g�b�v�y�[�W��'���O�C����'���\������Ă��Ă��^�C���A�E�g<br>&nbsp;&nbsp;(��" & Session.Timeout & "���j�ɂ��{��ʂ��\�������\��������܂��B"%></font></b></td>
	  <!--2011/11/16 M.Marquez Upd-E-->
	  <td></td>
	</tr>
        <tr>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
                  <td align="right"><img src="gif/1.gif" width=1 height=1><a href="javascript:OpenWindow()" ><font size=2 color="#0000ff">���p�ҏ��̍X�V</font></a></td>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
	          </td>
             </tr>
           </table>
	  	  </td>
        </tr><!--2009/02/25 Add-E G.Ariola  -->
      </table>
	  	  </td>
        </tr>
      </table>
	  	  </td>
        </tr>
      </table>
<br><br><br>
<a href="touroku/index.html" target="new">��ЃR�[�h�o�^���@</a>
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
<!-------------���O�C����ʏI���--------------------------->
<%
	If InStr(Request.QueryString("link"),"-expentry.asp")<>0 Then
		DispMenuBarBack "expentry.asp"
	ElseIf InStr(Request.QueryString("link"),"-impentry.asp")<>0 Then
		DispMenuBarBack "impentry.asp"
	Else
		DispMenuBarBack "index.asp"
	End If
%>
</body>
</html>

<%
End Function
%>

<%
'��ʂ̕\��
Function DispError
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
          <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>���O�C��</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>
      <table>
        <tr>
          <td nowrap align=center>
			<BR><BR>
            <dl>
				<img src="gif/error2.gif" width=210 height=63>
            </dl>
			<BR>
<%
            ' �G���[���b�Z�[�W�̕\��
            DispErrorMessage "���O�C���G���[�̂��߁A�g�p�ł��܂���B"
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
	If InStr(Request.QueryString("link"),"-expentry.asp")<>0 Then
		DispMenuBarBack "expentry.asp"
	ElseIf InStr(Request.QueryString("link"),"-impentry.asp")<>0 Then
		DispMenuBarBack "impentry.asp"
	Else
		DispMenuBarBack "index.asp"
	End If
%>
</body>
</html>

<%
End Function
%>

<%
' �����N��ʂ�\�����Ă悢���ǂ����̃`�F�b�N
Function CheckLinkKind(iNum,iWrkNum)
    ' �߂��ʏ����擾
    strLinkID = Session.Contents("linkid")

    strError=""
    Select Case strLinkID
        Case "nyuryoku-in1.asp"             ' �D�Ё^�^�[�~�i������
             If strUserKind<>"�D��" And strUserKind<>"�`�^" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�D�ЁA�`�^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
        Case "nyuryoku-kaika.asp", "nyuryoku-kaika2.asp"           ' �C�ݓ���  'Updated by seiko-denki 2003.07.21
             If strUserKind<>"�C��" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C��</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
        Case "nyuryoku-te.asp"              ' �^�[�~�i������
             If strUserKind<>"�`�^" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�`�^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
        Case "rikuun1.asp"                  ' ���^����
             If strUserKind<>"���^" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>���^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
        Case "ms-kaika.asp"                 ' �����d�l�C�ݓ���
             If strUserKind<>"�C��" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C��</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
' Commented by seiko-denki 2003.07.07
'        Case "ms-expentry.asp?kind=1"       ' �����d�l�A�o�R���e�i�Ɖ�
'             If strUserKind<>"�C��" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C��</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
'        Case "ms-expentry.asp?kind=2"       ' �����d�l�A�o�R���e�i�Ɖ�
'             If strUserKind<>"���^" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>���^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
'        Case "ms-expentry.asp?kind=3"       ' �����d�l�A�o�R���e�i�Ɖ�
'             If strUserKind<>"�׎�" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�׎�</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
'        Case "ms-expentry.asp?kind=4"       ' �����d�l�A�o�R���e�i�Ɖ�
'             If strUserKind<>"�`�^" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�`�^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
'        Case "ms-impentry.asp?kind=1"       ' �����d�l�A���R���e�i�Ɖ�
'             If strUserKind<>"�C��" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C��</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
'        Case "ms-impentry.asp?kind=2"       ' �����d�l�A���R���e�i�Ɖ�
'             If strUserKind<>"���^" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>���^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
'        Case "ms-impentry.asp?kind=3"       ' �����d�l�A���R���e�i�Ɖ�
'             If strUserKind<>"�׎�" Then
'                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�׎�</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
'             End If
' End of Comment by seiko-denki 2003.07.07
' Added by seiko-denki 2003.07.07
        Case "ms-expentry.asp"       ' �����d�l�A�o�R���e�i�Ɖ�
             If strUserKind<>"�C��" And strUserKind<>"���^" And strUserKind<>"�׎�" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C�݁A���^�A�׎�</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
        Case "ms-impentry.asp"       ' �����d�l�A�o�R���e�i�Ɖ�
             If strUserKind<>"�C��" And strUserKind<>"���^" And strUserKind<>"�׎�" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C�݁A���^�A�׎�</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
' End of Addition by seiko-denki 2003.07.07
        Case "pickselect.asp"             ' ��R���s�b�N�A�b�v�V�X�e��
             If strUserKind="�D��" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C�݁A���^�A�׎�A�`�^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If

        Case "hits.asp"                     ' �X�g�b�N���[�h���p
        Case "gate.asp"                     ' �Q�[�g�ʍs���ԗ\��

        Case "sokuji.asp"                   ' �������o�V�X�e��
             If strUserKind<>"�C��" And strUserKind<>"�`�^" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C�݁A�`�^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
' Added by seiko-denki 2003.12.25
        Case "SendStatus/sst000F.asp"             ' �X�e�[�^�X�z�M
             If strUserKind="�D��" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>�C�݁A���^�A�׎�A�`�^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
' End of Addition by seiko-denki 2003.12.15



'''''Added 20040131
        Case "Shuttle/SYWB013.asp"                  ' �V���g���\��
             If strUserKind<>"���^" Then
                 strError="</font><font color=#008800>�{�@�\��</font><font color=#ff0000>���^</font><font color=#008800>�ł̃��O�C�����݂̂��g�p�ɂȂ�܂��B"
             End If
''''Added 20040131 END



    End Select

    If strError<>"" Then
        DispLogIn(strError)

        ' File System Object �̐���
        Set fs=Server.CreateObject("Scripting.FileSystemobject")

        ' ���O�C���G���[
        WriteLog fs, iNum,"���O�C���G���[",iWrkNum, strUserKind & "," & "���͓��e�̐���:1(���)"
    End If
    CheckLinkKind = strError
End Function
%>

<%
    ' ���O�C���G���[�񐔂��`�F�b�N
    iError=CInt( Session.Contents("loginerror") )
    If iError>3 Then
        DispError
        Response.End
    End If

    ' �w������̎擾�i�߂��ʏ��j
    Dim strLinkID
    strLinkID = Request.QueryString("link")
    If strLinkID<>"" Then
        ' �߂��ʏ����Z�b�V�����ϐ��ɐݒ�
        Session.Contents("linkid") = strLinkID
        ' ���O�C���G���[�񐔂��N���A
        Session.Contents("loginerror") = 0
        iError=0
    End If

    ' �H�����̊ԁA���[�U�[�h�c�`�F�b�N�����Ȃ����
    Select Case strLinkID
        Case "hits.asp", "gate.asp"
            CheckLinkLog

            ' �߂��ʂփ��_�C���N�g
            Response.Redirect strLinkID
            Response.End
        Case Else
    End Select

    ' �Z�b�V�����̗L�������`�F�b�N
    Dim strSessionLink
    strSessionLink = Trim(Session.Contents("sessionlink"))
    ' �Z�b�V�����������ȂƂ�
    If strSessionLink="" Then
        ' �Z�b�V�����J�n���Z�b�V�����ϐ��ɐݒ�
        Session.Contents("sessionlink") = "on"

        ' �Z�b�V�����؂ꂪ�����ȉ�ʂ̂Ƃ��A���j���[�ɖ߂�

    End If


	Dim iNum,iWrkNum
' ���[�U�[ID���K�v�ȉ�ʂ��ǂ����`�F�b�N
' Select Case strLinkID
' strLinkID���ƃG���[���̃��O���擾�ł��Ȃ��̂ŃZ�b�V�����ɕύX���܂���	2002/2/21		
		Select Case Session.Contents("linkid")
		' ���[�U�[ID���K�v�ȉ��
		Case ""
		Case "hits.asp", "gate.asp"
		Case "nyuryoku-in1.asp"
				iNum = 3000
				iWrkNum = 10
		Case "nyuryoku-kaika.asp", "nyuryoku-kaika2.asp"  'Updated by seiko-denki 2003.07.21
				iNum = 4000
				iWrkNum = 10
		Case "nyuryoku-te.asp"
				iNum = 5000
				iWrkNum = 10
		Case "rikuun1.asp"
				iNum = 6000
				iWrkNum = 10
'		Case "ms-expentry.asp?kind=1"   ' Commented by seiko-denki 2003.07.07
'				iNum = 1100
'				iWrkNum = 11
'		Case "ms-expentry.asp?kind=2"
'				iNum = 1100
'				iWrkNum = 12
'		Case "ms-expentry.asp?kind=3"
'				iNum = 1100
'				iWrkNum = 13
'		Case "ms-expentry.asp?kind=4"
'				iNum = 1100
'				iWrkNum = 14
'		Case "ms-impentry.asp?kind=1"
'				iNum = 2100
'				iWrkNum = 11
'		Case "ms-impentry.asp?kind=2"
'				iNum = 2100
'				iWrkNum = 12
'		Case "ms-impentry.asp?kind=3"
'				iNum = 2100
'				iWrkNum = 13  ' End of Comment by seiko-denki 2003.07.07
		Case "ms-expentry.asp"
				iNum = 1100
				iWrkNum = 11
		Case "ms-impentry.asp"
				iNum = 2100
				iWrkNum = 11
		Case "sokuji.asp"
				iNum = 7000
				iWrkNum = 10
		Case "pickselect.asp"
				iNum = "a100"
				iWrkNum = 10
		Case "predef/dmi000F.asp","index.asp","index_en.asp","index_ch.asp" 'Y.TAKAKUWA Upd-S 2015-03-05�@index_ch.asp��ǉ��B
				iNum = "b000"
				iWrkNum = 10
		Case "SendStatus/sst000F.asp"  ' Added by seiko-denki 2003.12.25
				iNum = "c000"
				iWrkNum = 10             ' End of Addition by seiko-denki 2003.12.15
		Case "Shuttle/SYWB013.asp"		''''Added 20040131
				iNum = "d000"							''''Added 20040131
				iWrkNum = 10							''''Added 20040131
		' ���[�U�[ID���s�v�ȉ��
		Case "sokuji-kaika-list.asp", "sokuji-koun-list.asp"
		Case Else
				' �߂��ʂփ��_�C���N�g
 				CheckLinkLog
				Response.Redirect strLinkID
				Response.End
	End Select





    ' ���[�U�[ID�̗L�������`�F�b�N
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))
    
    '2010/05/17 Add-S C.Pestano
	if strLinkID = "index.asp" or strLinkID = "index_en.asp" or strLinkID = "index_ch.asp" then 'Y.TAKAKUWA Upd-S 2015-03-05�@index_ch.asp��ǉ��B
		strUserID = ""				
	end if
	'2010/05/17 Add-E C.Pestano
				
    ' �w������̎擾(���[�U�[�h�c)
    Dim strInputUserID, strInputPassWord
    strInputUserID = UCase(Trim(Request.QueryString("user")))
    strInputPassWord = UCase(Trim(Request.QueryString("pass")))
		
    ' ���[�U�[ID���L���ȂƂ�
    If strUserID<>"" And strInputUserID="" Then
        ' ���[�U��ނ��}�b�`���Ă��邩�`�F�b�N����
        strUserKind=Session.Contents("userkind")
        strError = CheckLinkKind(iNum,iWrkNum)
        If strError="" Then
            ' �߂��ʏ����擾
            strLinkID = Session.Contents("linkid")

            CheckLinkLog			
			
            ' �߂��ʂփ��_�C���N�g
            Response.Redirect strLinkID
        Else
            ' ���O�C���G���[�񐔂��J�E���g�A�b�v
            iError=iError+1
            Session.Contents("loginerror") = iError
        End If
    Else
        ' �G���[�t���O�̃N���A
        bOK = false
        bError = false		
		bUpdated = ""

        If strInputUserID<>"" Then
            ' ���̓��[�U�[�h�c�̃`�F�b�N
            ConnectSvr conn, rsd
            '2009/08/05 Add-S Tanaka
            '�����X�V������1�N�o�߂��Ă��邩�̃`�F�b�N
            sql=" select convert(char(10),dateadd(year,1,ForceDate),111) as ForceDate ,convert(char(10),getdate(),111)as Now from mParam Where Seq='1' "
            rsd.Open sql, conn, 0, 1, 1
            If Not rsd.EOF Then
            	'��N���o�߂��Ă���ꍇ�͌��ݔN�����ōX�V
            	if rsd("ForceDate")<=rsd("Now") Then
            		sql = "UPDATE mParam"
            		sql = sql & " SET "
            		sql = sql & "UpdtTime = '" & Now() & "' ,"		
            		sql = sql & "UpdtPgCd = 'Login', "
            		sql = sql & "UpdtTmnl = '" & Request.QueryString("user") & "' ,"
            		sql = sql & "ForceDate = '" & rsd("Now") & "'"
            		sql = sql & " Where Seq= '1'"
            		conn.execute sql
            	end if
            end if
            '2009/08/05 Add-E Tanaka

            ConnectSvr conn, rsd
            '2009/08/05 Add-E Tanaka
'=========== 03/07/17 �ύX =================================================================
			sql="select FullName,UserType,mParam.ForceDate,mUsers.UserUpdate, (CASE WHEN mUsers.UserUpdate > mParam.ForceDate and mParam.seq = '1' then '1' else '0' end) as 'Updated' from mUsers, mParam"
			sql=sql&" where UserCode='" & strInputUserID & "' and PassWord='" & strInputPassWord & "'"
			'SQL�𔭍s���ă��[�U�[�h�c������

			rsd.Open sql, conn, 0, 1, 1

			If Not rsd.EOF Then

				bOK = true				
				if request.querystring("Skip_Mode") = "1" then
					bUpdated = "1"
'2009/11/05 Upd-S Fujiyama
'				else
				elseif Now() < rsd("ForceDate") and IsNull(rsd("UserUpdate"))=false then
					bUpdated = "1"
'2009/11/05 Upd-E Fujiyama
				else
					bUpdated = Trim(rsd("Updated"))
				end if
				' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
				Session.Contents("userid") = strInputUserID
				' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
				Select Case Trim(rsd("UserType"))
					Case "1"
						Session.Contents("userkind") = "�׎�"
					Case "2"
						Session.Contents("userkind") = "�C��"
					Case "3"
						Session.Contents("userkind") = "�D��"
					Case "4"
						Session.Contents("userkind") = "�`�^"
					Case "5"
						Session.Contents("userkind") = "���^"
				End Select
				' ���O�C�������Z�b�V�����ϐ��ɐݒ�
				Session.Contents("username") = Trim(rsd("FullName"))
			End If
			rsd.Close
'=============================================================================================

'=========== 03/07/17 �R�����g�A�E�g =================================================================
            ' �׎�R�[�h�`�F�b�N
'             sql = "SELECT FullName FROM mShipper WHERE Shipper='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
            'SQL�𔭍s���ă��[�U�[�h�c������
'            rsd.Open sql, conn, 0, 1, 1
'            If Not rsd.EOF Then
'                bOK = true
                ' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
'                Session.Contents("userid") = strInputUserID
                ' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
'                Session.Contents("userkind") = "�׎�"
                ' ���O�C�������Z�b�V�����ϐ��ɐݒ�
'                Session.Contents("username") = Trim(rsd("FullName"))
'            End If
'            rsd.Close

'            If bOK=false Then
                ' �C�݃R�[�h�`�F�b�N
'                sql = "SELECT FullName FROM mForwarder WHERE Forwarder='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQL�𔭍s���ă��[�U�[�h�c������
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userid") = strInputUserID
                    ' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userkind") = "�C��"
                    ' ���O�C�������Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'            If bOK=false Then
                ' ���^�R�[�h�`�F�b�N
'                sql = "SELECT FullName FROM mTrucker WHERE Trucked='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQL�𔭍s���ă��[�U�[�h�c������
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userid") = strInputUserID
                    ' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userkind") = "���^"
                    ' ���O�C�������Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'            If bOK=false Then
                ' �D�ЃR�[�h�`�F�b�N
'                sql = "SELECT FullName FROM mShipLine WHERE ShipLine='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQL�𔭍s���ă��[�U�[�h�c������
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userid") = strInputUserID
                    ' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userkind") = "�D��"
                    ' ���O�C�������Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'            If bOK=false Then
                ' �`�^�R�[�h�`�F�b�N
'                sql = "SELECT FullName FROM mOperator WHERE OpeCode='" & strInputUserID & "' And sPassWord='" & strInputPassWord & "'"
                'SQL�𔭍s���č`�^�}�X�^�[������
'                rsd.Open sql, conn, 0, 1, 1
'                If Not rsd.EOF Then
'                    bOK = true
                    ' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userid") = strInputUserID
                    ' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("userkind") = "�`�^"
                    ' ���O�C�������Z�b�V�����ϐ��ɐݒ�
'                    Session.Contents("username") = Trim(rsd("FullName"))
'                End If
'                rsd.Close
'            End If

'=============================================================================================
if Request.QueryString("Screen_Mode") <> "R" then '2009/02/25 Add G.Ariola  
            If bOK=false Then
                ' ���[�U�[�h�c�G���[�̂Ƃ�
                bError=true
                strError = "���͂��ꂽ���e�ɊԈႢ������܂��B"
                ' ���O�C���G���[�񐔂��J�E���g�A�b�v
                iError=iError+1
                Session.Contents("loginerror") = iError
            End If
end if
            conn.Close
        End If

if Request.QueryString("Screen_Mode") <> "R" then '2009/02/25 Add G.Ariola 

        If Not bOK Then
            ' File System Object �̐���
            Set fs=Server.CreateObject("Scripting.FileSystemobject")

            ' ���O�C��
            If strInputUserID<>"" Then
                WriteLog fs, iNum,"���O�C��",iWrkNum, strInputUserID & "," & "���͓��e�̐���:1(���)" & iError
            Else
                WriteLog fs, iNum,"���O�C��", "00",","
            End If

            If iError>3 Then
                DispError
            Else
                If Not bError Then
                    strError=""
                    ' ���O�C���G���[�񐔂��J�E���g�A�b�v
                    iError=iError+1
                    Session.Contents("loginerror") = iError
                End If
                DispLogIn(strError)
            End If
        Else
		if bUpdated = "1" then '2009/02/25 Add G.Ariola  
            ' ���[�U��ނ��}�b�`���Ă��邩�`�F�b�N����
            strUserKind=Session.Contents("userkind")
            strError = CheckLinkKind(iNum,iWrkNum)
            If strError="" Then
                ' �߂��ʏ����擾
                strLinkID = Session.Contents("linkid")

                CheckLinkLog
				
                ' �߂��ʂփ��_�C���N�g
                Response.Redirect strLinkID
            Else
                ' ���[�U���N���A
                    Session.Contents("userid") = ""
                    Session.Contents("userkind") = ""
                    Session.Contents("username") = ""
                ' ���O�C���G���[�񐔂��J�E���g�A�b�v
                iError=iError+1
                Session.Contents("loginerror") = iError
            End If		 
		 '2009/02/25 Add-S G.Ariola  				
		 else
		    ' ���[�U���N���A
			Session.Contents("userid") = ""
			Session.Contents("userkind") = ""
			Session.Contents("username") = ""
			DispLogIn("")
		 end if
		 '2009/02/25 Add-E G.Ariola  
        End If
'2009/02/25 Add-S G.Ariola  		
else
	' ���[�U���N���A
	Session.Contents("userid") = ""
	Session.Contents("userkind") = ""
	Session.Contents("username") = ""
	DispLogIn("")
end if
'2009/02/25 Add-E G.Ariola  		
    End If

%>
