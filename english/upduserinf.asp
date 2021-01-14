<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%

Dim v1,v2,v3,v4,v5,v6
const v_MailServer = "MAIL_SERVER"

if Request.form("Gamen_Mode") = "U" then
        call UpdatemUsers()
		
		if Request.form("checkMail") = "on" then
			gfSendMail(TRIM(Request.form("MailAddress")))
		end if
		call SelectmUsers()
else
		call SelectmUsers()
End If

function SelectmUsers()
	dim sql
	dim conn,rs
	'----------------------------------------
    ' ＤＢ接続
    '----------------------------------------        
    ConnectSvr conn, rs

		sql = "Select FullName,Address,TelNo,TTName,MailAddress,PassWord from mUsers"
		sql = sql & " WHERE UserCode = '" & Request.QueryString("user") & "'"		

	rs.Open sql, conn, 0, 1, 1
	
	If Not rs.EOF Then
		v1 = Trim(rs("FullName"))
		v2 = Trim(rs("Address"))
		v3 = Trim(rs("TelNo"))
		v4 = Trim(rs("TTName"))
		v5 = Trim(rs("MailAddress"))
		v6 = Trim(rs("PassWord"))
	end if
	
	rs.Close
	conn.close	
end function	

function UpdatemUsers()
	dim sql
	dim conn,rs
	'----------------------------------------
    ' ＤＢ接続
    '----------------------------------------        
    ConnectSvr conn, rs

		sql = "UPDATE mUsers"
		sql = sql & " SET "
		sql = sql & "UpdtTime = '" & Now() & "' ,"		
        sql = sql & "UpdtPgCd = 'USERINF', "
		sql = sql & "UpdtTmnl = '" & Request.QueryString("user") & "' ,"
		sql = sql & "FullName = '" & TRIM(Request.form("FullName")) & "', "
		if Request.form("password1") <> "" then
		sql = sql & "PassWord = '" & Request.form("password1") & "', "
		end if
		sql = sql & "TelNo = '" & TRIM(Request.form("TelNo")) & "', "
		sql = sql & "Address = '" & TRIM(Request.form("Address")) & "', "
		sql = sql & "TTName = '" & TRIM(Request.form("TTName")) & "', "
		sql = sql & "MailAddress = '" & TRIM(Request.form("MailAddress")) & "', "
		if Request.form("checkMail") = "on" then
			sql = sql & "MailSend = '1', "
		else
		    '2009/08/04 M.Marquez Upd-S
			'sql = sql & "MailSend = '', "
			sql = sql & "MailSend = NULL, "
			'2009/08/04 M.Marquez Upd-E
		end if
		sql = sql & "UserUpdate = '" & Now() & "' "
		sql = sql & "WHERE UserCode = '" & Request.QueryString("user") & "'"		
	
	conn.execute sql
	conn.close	
end function	

function gfSendMail(mailto)
    Dim objMail
	dim mailfrom, subject, body, mailserver
	dim param(2)
	
	call getUploadIni(param,v_MailServer)
	l_MailServer = param(0)	

	gfSendMail = ""
	mailfrom = "mrhits@hits-h.com"
	subject = "利用者情報更新"
	'body = "" & "HiTS／利用者情報更新完了" & "" & vbCrLf & vbCrLf
	body = body & "HiTSの利用者情報の更新が完了しました。"
	Set objMail = CreateObject("BASP21")
	if trim(mailto)<>"" and trim(mailfrom)<>"" then
		gfSendMail=objMail.Sendmail(l_MailServer, mailto, mailfrom, subject, body, "")
		if gfSendMail<>"" then
			if left(gfSendMail,3)="501" then
				Set objMail = Nothing				
				exit function
			end if								
		end if		
	end if
	Set objMail = Nothing
end function

function getUploadIni(param,strVariable)
	dim ObjFSO,ObjTS,tmpStr
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- ファイルを開く（読み取り専用） ---
	'Y.TAKAKUWA Upd-S 2015-03-17
	'Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("adminmenu/ini/admin.ini"),1,false)
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("/adminmenu/ini/admin.ini"),1,false)
	'Y.TAKAKUWA Upd-E 2015-03-17
	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)			
		Select Case tmpStr(0)							
			Case strVariable							
				param(0) = tmpStr(1)
		End Select
	Loop
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function	
	
%>

<html>
<head>
<title>Information update for users</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
function finit(){
document.frm.Gamen_Mode.value="<%=Request.form("Gamen_Mode")%>";

	/*var w=520;
	var h=550;
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
	window.resizeTo(w,h);
    window.moveTo(l,t);
	*/
	
document.frm.FullName.value="<%=v1%>";
document.frm.Address.value="<%=v2%>";
document.frm.TelNo.value="<%=v3%>";
document.frm.TTName.value="<%=v4%>";
document.frm.MailAddress.value="<%=v5%>";
//document.frm.password1.value="<%=v6%>";
//document.frm.password2.value="<%=Request.form("password2")%>";		
document.frm.Gamen_Mode.value="<%=Request.form("Gamen_Mode")%>";
document.frm.checkMail.checked ="<%=Request.form("checkMail")%>";
document.frm.FullName.focus();

	if ("<%=Request.form("Gamen_Mode")%>" == "U")
	{
			document.frm.FullName.value="<%=Request.form("FullName")%>";
			document.frm.Address.value="<%=Request.form("Address")%>";
			document.frm.TelNo.value="<%=Request.form("TelNo")%>";
			document.frm.TTName.value="<%=Request.form("TTName")%>";
			document.frm.MailAddress.value="<%=Request.form("MailAddress")%>";
			alert("登録処理が完了しました。");
			if ("<%=Request.QueryString("flagwin")%>" == 1)
			{
			var CodeWin;
			//opener.parent.document.usercheck.user.value = "<'%=Request.QueryString("user")%>";
			//opener.parent.document.usercheck.pass.value = "<'%=v6%>";
			//opener.parent.document.usercheck.submit();
			CodeWin = location.replace("./userchk.asp?user=<%=Request.QueryString("user")%>&pass=<%=v6%>&Skip_Mode=1&link=predef/dmi000F.asp","codelist");
			}
			fClose();
	}
	else
	{
	document.frm.checkMail.checked ="on";
	}

}

function fUpdate(){
f=document.frm;

	if (CheckText() == false){
	return;
	}
		
	document.frm.Gamen_Mode.value = "U";
	document.frm.submit();

}

function CheckText(){
  f=document.frm;
  
  if(f.FullName.value==""){
    alert("必須入力項目です。");
	f.FullName.focus();
    return false;
  }
  
  if(CheckByte(f.FullName.value) == false){
	f.FullName.select();
	return false;
  }	
  
  if(gfGetLength(f.FullName.value) > 20){
   alert("入力値が正しくありません。");
	f.FullName.select();
	return false;
  }	
  
  if(f.Address.value==""){
    alert("必須入力項目です。");
	f.Address.focus();
    return false;
  }
  
  if (CheckByte(f.Address.value) == false){
	f.Address.select();
	return false;
  }
  
  if(gfGetLength(f.Address.value) > 200){
   alert("入力値が正しくありません。");
	f.Address.select();
	return false;
  }	
  
  if(f.TelNo.value==""){
    alert("必須入力項目です。");
	f.TelNo.focus();
    return false;
  }

  if(CheckPhone(f.TelNo.value) == false){
     f.TelNo.select();
     return false;
  }	
  
  if(f.TTName.value==""){
    alert("必須入力項目です。");
	f.TTName.focus();
    return false;
  }
  
  if(CheckByte(f.TTName.value) == false){
	f.TTName.select();
	return false;
  }
  
  if(gfGetLength(f.TTName.value) > 16){
   alert("入力値が正しくありません。");
	f.TTName.select();
	return false;
  }	
  
  if(f.MailAddress.value==""){
    alert("必須入力項目です。");
	f.MailAddress.focus();
    return false;
  }
  
  if(CheckEmail(f.MailAddress.value) == false){
    f.MailAddress.select();
    return false;
  }
  
  // 2009/07/23 C.Pestano Add-S
  /*
  if(f.password1.value==""){
    alert("必須入力項目です。");
	f.password1.focus();
    return false;
  }
  
  if(f.password1.value!="" && f.password2.value==""){
    alert("必須入力項目です。");
	f.password2.focus();
    return false;
  }
  */
  // 2009/07/23 C.Pestano Add-E
  
  if(CheckEisuji(f.password1.value)==false){
    alert("新パスワードは半角英数字で入力してください。");
	f.password1.select();
    return false;
  }
  if(CheckEisuji(f.password2.value)==false){
    alert("新パスワード(再入力)は半角英数字で入力してください。");
	f.password2.select();
    return false;
  }
  if(f.password1.value!=f.password2.value){
    alert("入力値が正しくありません。");
	f.password1.select();
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

function CheckEmail(str){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_@.";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
	alert("入力値が正しくありません。");
      return false;
    }
  }
  return true;
}

function CheckPhone(str){
  checkstr="0123456789-";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
	alert("入力値が正しくありません。");
      return false;
    }
  }
  return true;
}


function CheckByte(ls_str)
{ 
    var li_count;

    for (li_count = 0; li_count < ls_str.length; li_count++) {
        //ラテン文字の変換コード使用
        if ((ls_str.charCodeAt(li_count) >= 65377 && ls_str.charCodeAt(li_count) <=65439)) {
			alert("入力値が正しくありません。");
            return false;
        }

    }
    return true;
}

function gfGetLength(ls_str)
{ 
    var li_count;
    var ll_len = 0;

    for (li_count = 0; li_count < ls_str.length; li_count++) {
        //ラテン文字の変換コード使用
        if ((ls_str.charCodeAt(li_count) >= 32 && ls_str.charCodeAt(li_count) <= 255)) {
            ll_len++;
        }
        else {
            ll_len += 2;
        }
    }
    return ll_len;
}
function fClose(){	
	CodeWin = location.replace("./userchk.asp?link=predef/dmi000F.asp","codelist");
}
</SCRIPT>

<style type="text/css">
	/* 検索項目 */
	td.kaisha{
		height: 20px;
		font-size: 14px;
		color:#ffffff;
		background-color:#000099;
		padding: 3px 5px 3px 5px;
	}
	
	td.kodo{
		height: 20;
		font-size: 14px;
		color:#000000;
		background-color:#ff8833;
		padding: 3px 5px 3px 5px;
	}
	
	td.kodo1{
		height:    20px;
		font-size: 14px;
		color:#000000;
		background-color:#ffff99;
		padding: 3px 5px 3px 5px;
		list-style:inside
	}
	
	TD.bordering
	{
    BORDER-BOTTOM: 1px dotted #000000;
    BORDER-LEFT: 1px dotted #000000;
    BORDER-RIGHT: 1px dotted #000000;
    BORDER-TOP: 1px dotted #000000;	
	}
</style>

</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/loginback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<!-------------ここから登録コード一覧画面--------------------------->

<form name="frm" method="post">
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


<!--<table border=0 align="center">
<tr>
	<td align=left valign=middle class="kodo" width="80">会社コード</td>
	<td align=left valign=middle class="kodo1" width="50"><%=Request.QueryString("user")%></td>
</tr>
</table> -->

<table border=0 align="center">
	<tr>
  		<td colspan="2">
			<table align="center" border=0><INPUT name="Gamen_Mode" size="9" readonly tabindex= -1 type= hidden>
			 <tr>
				<td align=left valign=middle height="40">Please  update the information in case of any changes. </td>
			</tr>
			</table>  	
  		</td>
  	</tr>
	<tr>
  		<td colspan="2">
			<table border=0 align="center">
			<tr>
				<td align=left valign=middle class="kodo" width="90">COMPANY ID</td>
				<td align=left valign=middle class="kodo1" width="50"><%=Request.QueryString("user")%></td>
			</tr>
			</table>	  	
  		</td>
  	</tr>
  <tr>
  <td colspan="2">&nbsp;</td>
  </tr>
  <!--<tr>
  	<td width="430">&nbsp;</td>
  	<td colspan="2">&nbsp;</td>
  </tr> -->
  <tr>
  	<!--<td width="430">&nbsp;</td> -->
	<td align=left colspan="2">●INPUT INFORMATION</td>
  </tr>
  <tr>
  	<!--<td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha" bgcolor="" width="300">COMPANY NAME</td>
	<td align=left valign=middle height="20"><INPUT type=text name="FullName" maxlength=20 size=40></td>	
  </tr> 
  <tr>
    <!--<td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha">ADDRESS</td>
	<td align=left valign=middle height="20"><INPUT type=text name="Address" maxlength=200 size=40></td>
  </tr>
  <tr>
  	<!--<td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha">TEL NO.<br>12 DIGITS （Including the hyphen)
</td>
	<td align=left valign=middle height="20"><INPUT type=text name="TelNo" maxlength=12 size=40></td>	
  </tr>
  <tr>
    <!--<td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha">PERSON IN CHARGE</td>
	<td align=left valign=middle height="20"><INPUT type=text name="TTName" maxlength=16 size=40></td>
  </tr>
  <tr>
   <!-- <td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha">E-MAIL</td>
	<td align=left valign=middle height="20"><INPUT type=text name="MailAddress" maxlength=100 size=40></td>
  </tr>
  <tr>
  <tr>
    <!--<td>&nbsp;</td> -->
	<td align=left valign=middle height="20" colspan="2"><INPUT type=checkbox name="checkMail"><font size=2>Request email notice from the terminal </font></td>
  </tr>
  <tr>
    <!--<td>&nbsp;</td> -->
	<td align=left valign=middle colspan="2">
		<table border=0>
		  <tr height="20">
			<td align=left valign=middle class="bordering">
			<font size=2>※With or without the check mark, e-mail will be <BR>sent to all users in case of emergency. </font>
			</td>
		  </tr>
		</table>
	</td>
  </tr> 
  <tr>
  <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
  	<!--<td width="430">&nbsp;</td> -->
	<td align=left colspan="2">●CHANGE PASSWORD (VOLUNTARY. 8 DIGITS)</td>
  </tr>
  <tr>
  	<!--<td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha" width="100">NEW PASSWORD</td>
	<td align=left valign=middle height="20"><INPUT type=password name="password1" maxlength=8 size=15></td>	
  </tr>
  <!--2009/07/23 C.Pestano Add-S-->
  <tr>
  	<!--<td width="70">&nbsp;</td> -->
	<td align=left valign=middle colspan="2">REINPUT PASSWORD</td>
  </tr>
  <!--2009/07/23 C.Pestano Add-E-->
  <tr>
    <!--<td>&nbsp;</td> -->
	<td align=left valign=middle class="kaisha">NEW PASSWORD</td>
	<td align=left valign=middle height="20"><INPUT type=password name="password2" maxlength=8 size=15></td>
  </tr>
  <tr>
  <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
  		<td colspan="2" align="center">
			<table border=0 width=40%>
			  <tr height="35">
				<td align=center valign=middle>
			<input type=button value="   REGISTER   " onClick="fUpdate()">&nbsp;&nbsp;&nbsp;&nbsp;
			<input type=button value="   CLOSE   " onClick="fClose();">
				</td>
			</table>
  		</td>
  </tr>
  <tr>
  		<td colspan="2" align="center">
			<table border=0>
			  <tr height="20">
				<td align=center valign=middle class="bordering">
				Email will be sent when the registeration is complete.
				</td>
			  </tr>
			</table>
  		</td>
  </tr>
</table>




<br/>
<br/>
</center>
</form>
</body>
</html>

<%
%>
