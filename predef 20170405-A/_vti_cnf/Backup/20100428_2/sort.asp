<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
Dim strUser, str_left_menu

Dim v_Key1,v_Key2,v_Key3
Dim v_KeySort1,v_KeySort2,v_KeySort3

Dim FieldKey
dim v_loop

call LfGetRequestItem

Select Case  str_left_menu
	Case "1"
		ReDim FieldKey(18)
		FieldKey=Array("���o�\���","�w���� �| �R�[�h","��Ɣԍ�","�w����","�R���e�i�ԍ�/BL�ԍ�","�D��","�D��","SZ","CY","�t���[�^�C��","�[����P","��������","�ԋp�\��","�ԋp","�w����","�w�����","���l�P","���l�Q","�S��")
	Case "2"
		ReDim FieldKey(12)
		FieldKey=Array("�����\���","�w���� �| �R�[�h","�R���e�i�ԍ�","�D��","�D��","SZ","�ԋp��","�f�B�e���V�����t���[�^�C��","�w����","�w�����","���l","�S��")	
	Case "3"
		'2010/04/26 Upd-S C.Pestano
		ReDim FieldKey(17)	
		FieldKey=Array("���͓�","�w���� �| �R�[�h","�u�b�L���O�ԍ�","�s�b�N��","SZ","�^�C�v","H","�ގ�","�D��","�D��","CY�J�b�g��","��R�����o��","�w����","�w�����","���l�P","���l�Q","�S��")	
		'2010/04/26 Upd-E C.Pestano
 	Case "4"
		ReDim FieldKey(20)
		FieldKey=Array("�����\���","�w���� �| �R�[�h","��Ɣԍ�","�u�b�L���O�ԍ�","�R���e�i�ԍ�","�D��","�D��","SZ","H","TW","������","CY","CY�J�b�g��","��������","�w����","�w�����","���l�P","���l�Q","���l�R","�S��")
	'2010/02/20 C.Pestano Add-S
	Case "5"		
		'2010/04/26 Upd-S C.Pestano
		ReDim FieldKey(12)	'2010/04/27-3 Upd-E C.Pestano
		FieldKey=Array("���͓�","�w����","�w�����S����","�u�b�L���O�ԍ�","�s�b�N�ςݖ{��","SZ","�^�C�v","H","�ގ�","�D��","�D��","�w����","�w�����")	
		'2010/04/26 Upd-E C.Pestano
	Case "6"
		ReDim FieldKey(12) '2010/04/27-3 Upd-E C.Pestano
		FieldKey=Array("���o���\���","�w����","�w�����։�","��Ɣԍ�","�R���e�i�ԍ�/BL�ԍ�","�D��","�D��","SZ","������/���o��","CY","���o����","�t���[�^�C��","CY�J�b�g��")
	'2010/02/20 C.Pestano Add-E
end select

'FieldKey=Array("�����\���","��Ɣԍ�","�u�b�L���O�ԍ�","�R���e�i�ԍ�","�D��","�D��","SZ","H","TW","������","CY","CY�J�b�g��","��������")

'FieldName=Array("ITC.WorkDate","ITC.WkNo","CYV.BookNo","ITC.ContNo","CYV.ShipLine","CYV.VslName","CYV.ContSize","CYV.ContHeight","CYV.TareWeight","CYV.ReceiveFrom","BOK.RecTerminal","VSLS.CYCut","ITC.WorkCompleteDate")

if Request.form("Gamen_Mode") = "S" then
	if str_left_menu = "6" then
		Session("TB2Key1") = Request.form("Key1")
		Session("TB2Key2") = Request.form("Key2")
		Session("TB2Key3") = Request.form("Key3")	
		Session("TB2KeySort1") = Request.form("KeySort1")
		Session("TB2KeySort2") = Request.form("KeySort2")
		Session("TB2KeySort3") = Request.form("KeySort3")		
	else
		Session("Key1") = Request.form("Key1")
		Session("Key2") = Request.form("Key2")
		Session("Key3") = Request.form("Key3")	
		Session("KeySort1") = Request.form("KeySort1")
		Session("KeySort2") = Request.form("KeySort2")
		Session("KeySort3") = Request.form("KeySort3")
	end if
end if

function LfGetRequestItem()
	strUser = Request.QueryString("user")
	str_left_menu = Request.QueryString("left_menu")

	if str_left_menu = "6" then		
		v_Key1 = Session("TB2Key1")
		v_Key2 = Session("TB2Key2")
		v_Key3 = Session("TB2Key3")
		v_KeySort1 = Session("TB2KeySort1")
		v_KeySort2 = Session("TB2KeySort2")
		v_KeySort3 = Session("TB2KeySort3")	
	else
		v_Key1 = Session("Key1")
		v_Key2 = Session("Key2")
		v_Key3 = Session("Key3")
		v_KeySort1 = Session("KeySort1")
		v_KeySort2 = Session("KeySort2")
		v_KeySort3 = Session("KeySort3")
	end if
end function
%>

<html>
<head>
<title>���בւ��̎w��</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
function finit(){
	document.frm.Gamen_Mode.value="<%=Request.form("Gamen_Mode")%>";
	
	if (document.frm.Gamen_Mode.value == "S"){	
		document.frm.Key1.value="<%=Request.form("Key1")%>";
		document.frm.Key2.value="<%=Request.form("Key2")%>";
		document.frm.Key3.value="<%=Request.form("Key3")%>";
		Sorting();
		window.close();	
	}else{	
		//2010/04/27-3 Upd-S C.Pestano
		if ("<%=str_left_menu%>" == "5" || "<%=str_left_menu%>" == "6"){
			if("<%=v_Key1%>" == ""){
				document.frm.Key1.value="<%=v_Key1%>";
			}else{
				if("<%=Mid(v_Key1,1,1)%>" == "0"){ 
					document.frm.Key1.value="<%=Mid(v_Key1,2,1)%>";
				}else{
					document.frm.Key1.value="<%=v_Key1%>";
				}				
			} 
	
			if("<%=v_Key2%>" == ""){
				document.frm.Key2.value="<%=v_Key2%>";
			}else{
				if("<%=Mid(v_Key2,1,1)%>" == "0"){ 
					document.frm.Key2.value="<%=Mid(v_Key2,2,1)%>";
				}else{
					document.frm.Key2.value="<%=v_Key2%>";
				}	
			} 
	
			if("<%=v_Key3%>" == ""){
				document.frm.Key3.value="<%=v_Key3%>";
			}else{
				if("<%=Mid(v_Key3,1,1)%>" == "0"){ 
					document.frm.Key3.value="<%=Mid(v_Key3,2,1)%>";
				}else{
					document.frm.Key3.value="<%=v_Key3%>";
				}	

			} 			
		}else{
			document.frm.Key1.value="<%=v_Key1%>";
			document.frm.Key2.value="<%=v_Key2%>";
			document.frm.Key3.value="<%=v_Key3%>";
		}
		//2010/04/27-3 Upd-E C.Pestano
		document.frm.KeySort1.value="<%=v_KeySort1%>";
		document.frm.KeySort2.value="<%=v_KeySort2%>";
		document.frm.KeySort3.value="<%=v_KeySort3%>";

		if ("<%=v_KeySort1%>" == "DESC"){
			document.frm.KeySort1[1].checked=true;
		}else{
			document.frm.KeySort1[0].checked=true;
		}
		
		if ("<%=v_KeySort2%>" == "DESC"){
			document.frm.KeySort2[1].checked=true;
		}else{
			document.frm.KeySort2[0].checked=true;
		}
		
		if ("<%=v_KeySort3%>" == "DESC"){
			document.frm.KeySort3[1].checked=true;
		}else{
			document.frm.KeySort3[0].checked=true;
		}
	
	
		if (document.frm.Gamen_Mode.value == "S"){	
			Sorting();		
			window.close();	
		}
		
		document.frm.Key1.focus();
	}
}

function fSort(){
	str = "";
	
	if(document.frm.Key1.value == ""){
		str = str + "XX"
	}else{
		if(document.frm.Key1.value.length == 1){ 
			str = str + "0" + document.frm.Key1.value
		}else{
			str = str + document.frm.Key1.value
		}		
	}

	if (document.frm.KeySort1[0].checked){
		str = str + "0"
	}else{
		str = str + "1"
	}		
		
	if(document.frm.Key2.value == ""){
		str = str + "XX"
	}else{
		//2010/04/27-3 Upd-S C.Pestano
		if(document.frm.Key2.value.length == 1){ 
			str = str + "0" + document.frm.Key2.value
		}else{
			str = str + document.frm.Key2.value
		}	
		//2010/04/27-3 Upd-E C.Pestano
	}

	if (document.frm.KeySort2[0].checked){
		str = str + "0"
	}else{
		str = str + "1"
	}		
	
	if(document.frm.Key3.value == ""){
		str = str + "XX"
	}else{
		//2010/04/27-3 Upd-S C.Pestano
		if(document.frm.Key3.value.length == 1){ 
			str = str + "0" + document.frm.Key3.value
		}else{
			str = str + document.frm.Key3.value
		}	
		//2010/04/27-3 Upd-E C.Pestano
	}

	if (document.frm.KeySort3[0].checked){
		str = str + "0"
	}else{
		str = str + "1"
	}		

	if (fCHKDuplicate() == false){
			return;
	}
	
	if("<%=str_left_menu%>" == "5"){  
		createCookie('SortTbl1', str, 3)
	}else if("<%=str_left_menu%>" == "6"){
		createCookie('SortTbl2', str, 3)
	}
	
	document.frm.Gamen_Mode.value = "S";
	document.frm.submit();
}

function fClear(){
	document.frm.KeySort1[0].checked=true;
	document.frm.KeySort2[0].checked=true;
	document.frm.KeySort3[0].checked=true;
	document.frm.Key1.value="";
	document.frm.Key2.value="";
	document.frm.Key3.value="";
	
	if("<%=str_left_menu%>" == "5"){  
		createCookie('SortTbl1', "", 3)
	}else if("<%=str_left_menu%>" == "6"){
		createCookie('SortTbl2', "", 3)
	}
	
	document.frm.Gamen_Mode.value = "S";
	document.frm.submit();
}

function Sorting(){
	target=opener.document.serch;
	if ("<%=str_left_menu%>" == "4")  
	{
	  if(target.way[0].checked){
		opener.parent.DList.SerchC("4",target.SortKye.value);
	  } else if(target.way[1].checked){
		opener.parent.DList.SerchC("5",target.SortKye.value);
	  } else if(target.way[2].checked) {
		opener.parent.DList.SerchC("11",target.SortKye.value);
	  } else {
		opener.parent.DList.SerchC("","");
	  }
	}else if ("<%=str_left_menu%>" == "1")  
	{
	  if(target.way[0].checked){
		opener.parent.DList.SerchC("3",target.SortKye.value);
	  } else if(target.way[1].checked){
		opener.parent.DList.SerchC("4",target.SortKye.value);
	  } else if(target.way[2].checked) {
		opener.parent.DList.SerchC("11",target.SortKye.value);
	  } else {
		opener.parent.DList.SerchC("","");
	  }
	}else if("<%=str_left_menu%>" == "5"){
		opener.SerchC1("4",opener.document.frm.SortKey1.value);
		return;	
	}else if("<%=str_left_menu%>" == "6"){
		opener.SerchC2("4",opener.document.frm.SortKey2.value);
		return;
	}else if("<%=str_left_menu%>" == "3"){
	    opener.parent.DList.SerchC("0",target.SortKye.value);
	    //return;
	}else{
		opener.parent.DList.SerchC("4",target.SortKye.value);
	}

	opener.parent.Top.sort();	
	/*target2 = opener.parent.window.frames(0);
	target3 = opener.parent.window.frames(1)
	target1 = target2.window.document.getElementById("loading");
	target1.style.display='block';
	target3.window.document.getElementById("content").style.display='none';*/
}

function fCHKDuplicate(){
	var v_objectKey_x;
	var v_objectKey_y;
	
	var l_item ;
	var l_loop;

	for(l_item =1; l_item <= 3; l_item++){
		v_objectKey_x = eval("document.frm.Key" + l_item);
		for(l_loop =(l_item + 1); l_loop <= 3; l_loop++){
			v_objectKey_y = eval("document.frm.Key" + l_loop);

			if (v_objectKey_y.value!="")
			{
				if(v_objectKey_x.value==v_objectKey_y.value)
				{
				alert("���͒l������������܂���B")
				v_objectKey_y.focus();
				return false;
				}
			}
		}

	}

    return true;
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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
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
  <tr>
	<td align=left valign=middle width="200" height="30" colspan="2"><INPUT name="Gamen_Mode" size="9" readonly tabindex= -1 type= hidden>&nbsp;</td>
	<td align=center valign=middle width="55" height="30">����</td>
	<td align=center valign=middle width="55" height="30">�~��</td>
  </tr>
  <tr>
	<td align=left valign=middle width="80" height="30">���L�[</td>
	<td align=left valign=middle width="120" height="30">
	<select name="Key1">
	<option value=""><%="�i�w�薳���j" %>
	<%
		for v_loop = 0 to ubound(FieldKey)
			Response.Write "<OPTION VALUE ='" & v_loop & "'"
			Response.write ">" & FieldKey(v_loop)
		next
		
	%>
	</select>
	</td>
	<td align=center valign=middle height="30"><INPUT type=radio name="KeySort1" value="ASC"></td>
	<td align=center valign=middle height="30"><INPUT type=radio name="KeySort1" value="DESC"></td>
  </tr>
  <tr>
	<td align=left valign=middle height="30">���L�[</td>
	<td align=left valign=middle height="30">
	<select name="Key2">
	<option value=""><%="�i�w�薳���j" %>
	<%
		for v_loop = 0 to ubound(FieldKey)
			Response.Write "<OPTION VALUE ='" & v_loop & "'"
			Response.write ">" & FieldKey(v_loop)
		next
		
	%>
	</select>
	</td>
	<td align=center valign=middle height="30"><INPUT type=radio name="KeySort2" value="ASC"></td>
	<td align=center valign=middle height="30"><INPUT type=radio name="KeySort2" value="DESC"></td>
  </tr>
  <tr>
	<td align=left valign=middle height="30">��O�L�[</td>
	<td align=left valign=middle height="30">
	<select name="Key3">
	<option value=""><%="�i�w�薳���j" %>
	<%
		for v_loop = 0 to ubound(FieldKey)
			Response.Write "<OPTION VALUE ='" & v_loop & "'"
			Response.write ">" & FieldKey(v_loop)
		next
		
	%>
	</select>
	</td>
	<td align=center valign=middle height="30"><INPUT type=radio name="KeySort3" value="ASC"></td>
	<td align=center valign=middle height="30"><INPUT type=radio name="KeySort3" value="DESC"></td>
  </tr>

</table>

<BR>

<table border=0 width=85%>
  <tr>
	<td align=center valign=middle>
		����̃L�[���珇�Ɏw�肵�Ă�������
	</td>
  </tr>
</table>

<% End If %>

<BR>
<table border=0 width=85%>
  <tr height="40">
	<td align=center valign=middle>
<input type=button value="   ���s   " onClick="fSort()">&nbsp;&nbsp;&nbsp;&nbsp;
<input type=button value="   ���~   " onClick="JavaScript:window.close()">
	</td>
  </tr>
  <tr>
	<td align=center valign=middle>
<input type=button value=" ���בւ����� " onClick="fClear()">
	</td>
  </tr>
</table>
</form>

</center>
</body>
</html>

<%
%>
