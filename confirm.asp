<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%

%>

<html>
<head>
<title>利用者情報更新確認</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
function finit(){

document.frm.Update.focus();
//alert("<'%=Request.QueryString("pass")%>");
}

function fUpdate(){
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
 	
  CodeWin = location.replace("./upduserinf.asp?user=<%=Request.QueryString("user")%>&flagwin=1&link=predef/dmi000F.asp","codelist");

}

function fClose(){		
	//opener.parent.document.usercheck.Skip_Mode.value="1";
	//opener.parent.document.usercheck.submit();
	//window.close();	
	
	//2009/12/04 Upd-S C.Pestano
	CodeWin = location.replace("./userchk.asp?link=predef/dmi000F.asp","codelist");
	//2009/12/04 Upd-E C.Pestano
}
</SCRIPT>

</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/loginback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<!-------------ここからログイン入力画面--------------------------->
<form name="frm" method="post">
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
<BR>
<CENTER>
<table border=0><INPUT name="Gamen_Mode" size="9" readonly tabindex= -1 type= hidden>
 <tr>
	<td align=left valign=middle colspan="2"  height="50">利用者情報を更新してください。<BR></td>
 </tr>
</table>

<BR>
<table border=0 width=85%>
  <tr height="40">
	<td align=center valign=middle>
<input type=button name="Update" value="   更新   " onClick="fUpdate()">&nbsp;&nbsp;&nbsp;&nbsp;
<!--<input type=button value="   閉じる   " onClick="JavaScript:window.close()">-->
<input type=button value="   閉じる   " onClick="fClose();">
	</td>
  </tr>
</table>
</form>

</body>
</html>

<%
%>
