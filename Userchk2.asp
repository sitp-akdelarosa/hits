<!--#include file="Common.inc"-->
<%
'for each name in session.contents
'	response.write(name &"===="& session(name) &"<br>")
'next
'response.end
%>
<%
	' �߂���
	dim ReturnUrl
	ReturnUrl = request.querystring("ReturnUrl")
	' �w������̎擾(���[�U�[�h�c)
	Dim strInputUserID, strInputPassWord
	
	If UCase(Trim(Request.form("user")))<>"" and UCase(Trim(Request.form("pass")))<>"" then
		Session.Contents("userkind")=""
	End If

	If Session.Contents("userid")<>"" and Session.Contents("userkind") = "���^" Then
		strInputUserID = Session.Contents("userid")
		ReturnUrl = ReturnUrl & "?UserId=" & strInputUserID
		response.redirect ReturnUrl
	Else
	strInputUserID = UCase(Trim(Request.form("user")))
	strInputPassWord = UCase(Trim(Request.form("pass")))
	End If

	bOK = false
	bError = false

	If strInputUserID<>"" Then
		' ���̓��[�U�[�h�c�̃`�F�b�N
		ConnectSvr conn, rsd

		' ���^�R�[�h�`�F�b�N
			sql="select FullName from mUsers"
			sql=sql&" where UserCode='" & strInputUserID & "' and PassWord='" & strInputPassWord & "' and UserType='5'"
		'SQL�𔭍s���ă��[�U�[�h�c������
		rsd.Open sql, conn, 0, 1, 1
		If Not rsd.EOF Then
			bOK = true
			' ���O�C���h�c���Z�b�V�����ϐ��ɐݒ�
			Session.Contents("userid") = strInputUserID
			' ���O�C����ʂ��Z�b�V�����ϐ��ɐݒ�
			Session.Contents("userkind") = "���^"
			' ���O�C�������Z�b�V�����ϐ��ɐݒ�
			Session.Contents("username") = Trim(rsd("FullName"))
		End If
		rsd.Close

		If bOK=false Then
		    ' ���[�U�[�h�c�G���[�̂Ƃ�
		    bError=true
		    strError = "���͂��ꂽ���e�ɊԈႢ������܂��B"
		    ' ���O�C���G���[�񐔂��J�E���g�A�b�v
	'	    iError=iError+1
	'	    Session.Contents("loginerror") = iError
		End If

		conn.Close
	End If
	
	If bOK=true Then
		ReturnUrl = ReturnUrl & "?UserId=" & strInputUserID
		response.redirect ReturnUrl
	End If

%>


<html>
<head>
<title>���O�C��</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
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
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/loginback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
	DisplayCodeListButton
%>
          </td>
        </tr>
      </table>
      <center>

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
          <td nowrap align="center"> 
            <form action="userchk2.asp?ReturnUrl=<%=ReturnUrl%>" method="post" name="usercheck">
              <dl>
                <dd>��ЃR�[�h�ƃp�X���[�h����͂��A�w���M�x�{�^�����N���b�N���Ă��������B 
              </dl>
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
                <input type=submit value=" ���@�M " onClick="return Check()"></center>
              </form>
          </td>
        </tr>
      </table>
	<% If strError<>"" then %>
		<table border=1 cellpadding="2" cellspacing="1">
		<tr>
			<td bgcolor="#FFFFFF">
				<table border="0">
				<tr>
					<td valign="middle"><img src="gif/error.gif"></td>
					<td><b><font color="red"><% =strError %></font></b></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	<% End If %>
	  </td>
	  <td>
	  </td>
	 </tr>
        <tr>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
                  <td><img src="gif/1.gif" width=1 height=1></td>
                  <td><img src="gif/bo-yellow.gif" width="18" height="18"></td>
	          </td>
             </tr>
           </table>
	  	  </td>
        </tr>
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
</body>
</html>

