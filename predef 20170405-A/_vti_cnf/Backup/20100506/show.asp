<%
'**********************************************
'  �y�v���O�����h�c�z�@: 
'  �y�v���O�������́z�@: 
'
'  �i�ύX�����j
'	2010/04/26	C.Pestano	���ށ�SZ,������H�̏C��
'**********************************************
Option Explicit
Response.Expires = 0
'HTTP�R���e���c�^�C�v�ݒ�
Response.ContentType = "text/html; charset=Shift_JIS"
Response.AddHeader "Pragma", "no-cache" 
%>
<%	'**********************************************
  	'���ʂ̑O�񏈗�
  	'���ʊ֐�  (Commonfunc.inc)
%>
<!--#include file="Common.inc"-->
<%
	'**********************************************
	
	dim strUser, str_show_column, str_Title
	dim FieldKey
	dim v_loop
	dim v_ItemName
	dim v_ItemName2
	
	call LfGetRequestItem()

	Select Case  str_show_column
		Case "1"
			'2010/04/26 Upd C.Pestano
			ReDim FieldKey(13)			
			FieldKey=Array("���͓�","�w����","�w������","�w�����񓚑I��","�u�b�L���O�ԍ�","�s�b�N�ςݖ{��","SZ","�^�C�v","H","�ގ�","�D��","�D��","�w����","�w�����")	
		Case "2"
			ReDim FieldKey(14)
			FieldKey=Array("�����[�o��","���o���\���","�w����","�w�����։�","�w�����񓚑I��","��Ɣԍ�","�R���e�i�ԍ�/BL�ԍ�","�D��","�D��","SZ","������/���o��","CY","���o����","�t���[�^�C��","CY�J�b�g��")		
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
		fCheck(str,i,"Chk_Sel"+i,"Chk_SelDis"+i);
	}

}

function fCheck(str,colNo,id,name){
	if(str.charAt(colNo) == "0"){
		document.getElementById(id).innerHTML = "";		
		document.getElementById(id).innerHTML = "<input type=checkbox name=" + name + "></td>"
    }	      
}

function fFormatPage(){
	chk = document.getElementsByTagName('input');
	str = "";
	
	if(fChkDisplay() == false){
		return false;
	}
	
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
	opener.document.frm.submit();
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

function fChkDisplay(){
	var cnt;
	cnt = 0;
		
	for (i = 0; i <= <%=UBound(FieldKey)%>; i++) {
        if (chkobj("Chk_SelDis" + i)) {  //�`�F�b�N�{�b�N�X���`�F�b�N����Ă���ꍇ
            cnt++;
        }
    }
    if(cnt == 0) {
        window.alert("�S�Ă̍��ڂ��\���ɂ��鎖�͂ł��܂���B");
        return false;
    }
}
function fClear(){
	var obj;

	//�`�F�b�N�{�b�N�X�����[�v
	for (i = 0; i <= <%=UBound(FieldKey)%>; i++) {
		document.frm.elements["Chk_SelDis"+i].checked = true;
	}

	if("<%=str_show_column%>" == "1"){  
		createCookie('HitsTbl1', "", 15)
	}else{
		createCookie('HitsTbl2', "", 14)
	}
	opener.document.frm.submit();
	window.close();
}

//�����������`�F�b�N����Ă��邩�ǂ������m�F
//�߂�l�F�P �`�F�b�N����Ă���
//�@�@�@�@�O �`�F�b�N����Ă��Ȃ�
function chkobj(id)
{
    var obj;
    obj = eval("document.frm." + id);	
	if(obj != null){			
    	return (obj.checked) ? 1 : 0;
	}
}
</SCRIPT>
</head>
<body onLoad="fInit();" bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^�R�[�h�ꗗ���--------------------------->
<center>
<BR>
<% If strUser="" Then %>
	<table border=1 cellpadding=3 cellspacing=1 bgcolor="#ffffff">
		<tr>
			<td align=center nowrap>
				<font color="#ff3300"><b>���O�C�����Ă��Ȃ����͕\���ł��܂���B</b></font>
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
		<% v_ItemName2 = "Chk_SelDis" + cstr(v_loop) %>
		<td id="<%=v_ItemName%>"><input type="checkbox" name="<%=v_ItemName2%>" checked></td>		
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
	<tr>
		<td colspan="2" align=center valign=middle>
		<input type=button value=" �S�ĕ\�� " onClick="fClear()">
		</td>
	</tr>
</table>  
</form>
<% End If %>
</center>
</body>
</html>