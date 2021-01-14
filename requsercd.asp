<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
Dim strInputUserID, strInputPassWoed, strError
dim Arr_MailSig
redim Arr_MailSig(0)

strInputUserID = UCase(Trim(Request.form("user")))
strInputPassWord = UCase(Trim(Request.form("pass")))
strError = ""

If strInputUserID <> "" Then

            ' 入力ユーザーＩＤのチェック
            ConnectSvr conn, rsd

			sql="select FullName,UserType from mUsers"
			sql=sql&" where UserCode='" & strInputUserID & "' and PassWord='" & strInputPassWord & "'"
			'SQLを発行してユーザーＩＤを検索
			rsd.Open sql, conn, 0, 1, 1
			If rsd.EOF Then
			 	strError = "入力された内容に間違いがあります。"
			End If
			rsd.Close
            conn.Close
End If			
	
%>

<html>
<head>
<title>利用者入力</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
function finit(){
document.frm.Gamen_Mode.value="<%=Request.form("Gamen_Mode")%>";
document.frm.user.focus();

if ("<%=Request.form("Gamen_Mode")%>" == "S")
{
	if ("<%=strError%>" != "")
	{
		alert("<%=strError%>");
		document.frm.user.focus();
	}
else
	{

		OpenUpdUserinfWin();		
	}	
}

}

function fIns(){
if (Check() == false){
			return;
}	
document.frm.Gamen_Mode.value = "S";
document.frm.submit();

}

function Check(){
  f=document.frm;
  userid = f.user;
  pass = f.pass;
  ret = CheckEisuji(userid.value);
  if(ret==false){
    alert("会社コードは半角英数字で入力してください。");
    return false;
  }
  
  if(userid.value==""){
    alert("必須入力項目です。");
	userid.focus();
    return false;
  }
  
  if(pass.value==""){
    alert("必須入力項目です。");
	pass.focus();
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
 	
  CodeWin = location.replace("./upduserinf.asp?user=<%=strInputUserID%>&flagwin='2'","codelist","scrollbars=yes,resizable=no,width="+w+",height="+h+",top="+t+",left="+l);
  //CodeWin.focus();

}
function fClose(){	
	CodeWin = location.replace("./userchk.asp?link=predef/dmi000F.asp","codelist");
}
</SCRIPT>

<style type="text/css">
	/* 検索項目 */
	td.kaisha{
		width:    70px;
		height:    23px;
		font-size: 14px;
		color:#ffffff;
		font-weight:bold;
		background-color:#000099;
		padding: 3px 5px 3px 5px;
	}
</style>

</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/loginback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();" tabindex= -1>
<!-------------ここから登録コード一覧画面--------------------------->
<form name="frm" method="post" tabindex= -1>
<SCRIPT src="/adminmenu/Common/KeyDown.js" type=text/javascript></SCRIPT>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
  <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
</tr>
<tr>
  <td align="right" width="100%" height="48"> 
  </td>
</tr>
</table>
<br>
<center>
<table border=0><INPUT name="Gamen_Mode" size="9" readonly tabindex= -1 type= hidden>
 <tr>
	<td align=left valign=middle colspan="3"  height="50">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会社コードとパスワードを入力してください。<BR></td>
	
  </tr>
  <tr>
  	<td width="40">&nbsp;</td>
	<td align=left valign=middle class="kaisha">会社コード</td>
	<td align=left valign=middle height="30"><INPUT type=text name="user" maxlength=5 size=15 tabindex=1></td>	
  </tr>
  <tr>
    <td>&nbsp;</td>
	<td align=left valign=middle class="kaisha">パスワード</td>
	<td align=left valign=middle height="30"><INPUT type=password name="pass" maxlength=8 size=15 tabindex=2></td>
  </tr>
  
</table>


<BR>
<table border=0 width=85%>
  <tr height="40">
	<td align=center valign=middle><img src="gif/1.gif" width=15 height=1 >
<input type=button value="   次へ   " onClick="fIns()" tabindex=3>&nbsp;&nbsp;&nbsp;&nbsp;
<input type=button value="   中止   " onClick="fClose();" tabindex=4>
	</td>
  </tr>
</table>
</form>

</center>
</body>
</html>

<%
%>
