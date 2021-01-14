<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	dim strUser, str_show_column, str_Title
	dim FieldKey
	dim v_loop
	dim v_ItemName

	call LfGetRequestItem()

	Select Case  str_show_column
		Case "1"
			ReDim FieldKey(13)
			'FieldKey=Array("入力日","指示元","指示元へ回答"	,"ブッキング番号","ピック済","SZ","タイプ","高さ","材質","船社","船名","CYカット日","コン搬出先","指示先","指示先回答","備考１","備考２","担当")			
			FieldKey=Array("入力日","指示元","指示元回答","指示元回答選択","ブッキング番号","ピック済み本数","サイズ","タイプ","高さ","材質","船社","船名","指示先","指示先回答")	
		Case "2"
			ReDim FieldKey(14)
			FieldKey=Array("搬入票出力","搬出入予定日","指示元","指示元へ回答","指示元回答選択","作業番号","コンテナ番号/BL番号","船社","船名","サイズ","搬入元/搬出先","CY","搬出可否","フリータイム","CYカット日")		
	end select

function LfGetRequestItem()
	strUser = Request.QueryString("user")
	str_show_column = Request.QueryString("show_column")
	str_Title = Request.QueryString("pagetitle")
end function
%>

<html>
<head>
<title><%=str_Title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
function fInit(){
	//str = opener.document.cookie;
	
	var str;
	
	if("<%=str_show_column%>" == "1"){ 
		str = readCookie('HitsTbl1')		
	}else{
		str = readCookie('HitsTbl2')
	}		
			
	if(str == null) return false;
	
	for(i=0;i<15;i++){
		fCheck(str,i,"Chk_Sel"+i);
	}

}

function fCheck(str,colNo,id){
	if(str.charAt(colNo) == "0"){
		document.getElementById(id).innerHTML = "";
		document.getElementById(id).innerHTML = "<input type=checkbox></td>"
    }	      
}

function fFormatPage(){
	chk = document.getElementsByTagName('input');
	str = "";
	for(i=0; i<chk.length; i++){
		if (chk[i].type == "checkbox"){
			if(chk[i].checked == true){
				str = str + "1";
			}else{
				str = str + "0";
			}	
		}
	}
	
	if("<%=str_show_column%>" == "1"){  
		createCookie('HitsTbl1', str, 15)
	}else{
		createCookie('HitsTbl2', str, 14)
	}
	opener.finit();
	window.close();
}

function createCookie(name,value,days) {
	if (days) {
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000));
		var expires = "; expires="+date.toGMTString();
	}
	else var expires = "";
	opener.document.cookie = name+"="+value+expires+"; path=/";
}

function readCookie(name) {
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++) {
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length);
		if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
	}
	return null;
}
</SCRIPT>
</head>
<body onLoad="fInit();" bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録コード一覧画面--------------------------->
<center>
<BR>
<% If strUser="" Then %>
	<table border=1 cellpadding=3 cellspacing=1 bgcolor="#ffffff">
		<tr>
			<td align=center nowrap>
				<font color="#ff3300"><b>ログインしていない時は表示できません。</b></font>
			</td>
		</tr>
	</table>
	<BR>
<% Else %>
<form name="frm" method="post">
<table border=0>
	<% for v_loop = 0 to ubound(FieldKey) %>
	<tr>
		<% v_ItemName = "Chk_Sel" + cstr(v_loop) %>
		<td id="<%=v_ItemName%>"><input type="checkbox" checked></td>		
		<td><%=FieldKey(v_loop)%></td>		
	</tr>
	<% next %>
	<tr>
		<td height="10">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2">
			<table border="0" cellpadding="2" cellspacing="0" width="100%">
			<tr>			
			<td align=center><input name="btn1" type="button" value="   OK   " onClick="fFormatPage();"></td>
			<td align=center><input type="button" value="CANCEL" onClick="JavaScript:window.close()"></td>
			</tr>
			</table>
		</td>
	</tr>
</table>  
</form>
<% End If %>
</center>
</body>
</html>