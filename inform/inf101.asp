<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf101.asp                                      _/
'_/	Function	:お知らせメールアドレス登録画面                  _/
'_/	Date			:2005/03/03                                      _/
'_/	Code By		:aspLand HARA                                    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="Common.inc"-->
<%
	'''HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"

	Dim EMAIL
	EMAIL=Request.Form("email")

	'''エラートラップ開始
	on error resume next
	'''DB接続
	Dim cn, rs, sql, cnt
	ConnDBH cn, rs

	sql="select * from send_information where email='" & EMAIL & "'"
	rs.open sql, cn, 3, 1
	if err <> 0 then
		DisConnDBH cn, rs	'DB切断
		response.write("inf101.asp:send_informationテーブルアクセスエラー!")
		response.end
	end if

	Dim GROUP_CODE, COMPANY_NAME, NAME, TEL, ADDRESS, exist_flag
	exist_flag = 0
	if rs.RecordCount > 0 then
		GROUP_CODE = rs("group_code")
		COMPANY_NAME = rs("company_name")
		NAME = rs("user_name")
		TEL = rs("tel")
		ADDRESS = rs("address")
		exist_flag = 1
	end if
	rs.close

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>メールアドレス登録</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./js/common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(500,350);
window.focus();

function GoEntry(){
	f=document.inf101;

	if(f.dantai.options[f.dantai.selectedIndex].value == ""){
		alert("業種を選択してください。");
		f.dantai.focus();
		return false;
	}
	if(f.company_name.value == ""){
		alert("会社名を入力してください。");
		f.company_name.focus();
		return false;
	}
	if(f.user_name.value == ""){
		alert("氏名を入力してください。");
		f.user_name.focus();
		return false;
	}
	if(f.tel.value == ""){
		alert("電話番号を入力してください。");
		f.tel.focus();
		return false;
	}else{
		if(!checkPhoneNumber(f.tel.value)){
			alert("電話番号は半角数字で入力してください。");
			f.tel.focus();
			return false;
		}
	}
	if(f.address.value == ""){
		alert("住所を入力してください。");
		f.address.focus();
		return false;
	}

	if(<%=exist_flag%>){
		if(confirm("更新します。よろしいですか？")){
			f.action="inf103.asp";
			f.submit();
		}else{
			return false;
		}
	}else{
		if(confirm("登録します。よろしいですか？")){
			f.action="inf102.asp";
			f.submit();
		}else{
			return false;
		}
	}
}
function checkPhoneNumber(a){
	if(a==""){
		return(true);
	}
	var b=a.replace(/[0-9\-]/g,'');
	if(b.length!=0){
		return(false);
	}
	return(true);
}
function GoDelete(){
	f=document.inf101;
	if(!<%=exist_flag%>){
		alert("まだ登録されていませんので削除は無効です！");
		return false;
	}
	if(confirm("削除します。よろしいですか？")){
		f.action="inf104.asp";
			f.submit();
	}else{
		return false;
	}
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------メールアドレス等登録画面--------------------------->
<% Session.Contents("InsertSubmitted")="False"  %>
<% Session.Contents("UpdateSubmitted")="False"  %>
<% Session.Contents("DeleteSubmitted")="False"  %>
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="inf101" method="POST">
<input type="hidden" name="email" value="<%=EMAIL%>">
	<TR>
		<TD colspan="2">
			<b>
				<font color="navy">
					●新規登録、または、更新の場合<br>
					　　　下記をすべて入力して「登録」ボタンを押してください。<br>
					●削除の場合<br>
					　　　「削除」ボタンを押してください。
				</font>
			</b>
		</TD>
	</TR>
<% if exist_flag=0 then %>
	<tr><td colspan="2" align="center"><font color="red">新規登録依頼です</font></td></tr>
<% end if %>
	<TR>
		<TD align="right">メールアドレス：</TD>
		<TD>
			<%=EMAIL%>
		</TD>
	</TR>
	<TR>
		<TD width="25%" align="right">業種：</TD>
		<TD width="75%">

			<select name="dantai">
				<option value="">--選択してください--</option>
<%
					sql = "select * from group_name order by group_code"
					rs.open sql,cn,3,1
					if err <> 0 then
						'''DB切断
						DisConnDBH cn, rs
						response.write("inf101.asp:group_nameテーブルアクセスエラー!")
						response.end
					end if
					while not rs.EOF
						if GROUP_CODE = rs("group_code") then
%>
							<option value="<%=rs("group_code")%>" selected><%=rs("group_name")%></option>
<%					else	%>
							<option value="<%=rs("group_code")%>"><%=rs("group_name")%></option>
<%					end if
						rs.movenext
					wend
					rs.close
					'''DB接続解除
					DisConnDBH cn, rs
					'''エラートラップ解除
					on error goto 0
%>
			</select>
		</TD>
	</TR>
	<TR>
		<TD align="right">会社名：</TD>
		<TD>
			<INPUT type="text" name="company_name" value="<%=COMPANY_NAME%>" size="45" maxlength="25">
		</TD>
	</TR>
	<TR>
		<TD align="right">氏名：</TD>
		<TD>
			<INPUT type="text" name="user_name" value="<%=NAME%>" size="17" maxlength="10">
		</TD>
	</TR>
	<TR>
		<TD align="right">連絡先(電話番号)：</TD>
		<TD>
			<INPUT type="text" name="tel" value="<%=TEL%>" size="17" maxlength="13">&nbsp;(記入例：092-123-4567)
		</TD>
	</TR>
	<TR>
		<TD align="right">住所：</TD>
		<TD>
			<INPUT type="text" name="address" value="<%=ADDRESS%>" size="60" maxlength="50">
		</TD>
	</TR>
	<TR>
		<TD colspan="2" align="center">
			<INPUT type="button" value="戻る" onClick="javascript:history.back();">　　
			<INPUT type="button" value="登録" onClick="GoEntry()">　
			<INPUT type="button" value="削除" onClick="GoDelete()">　　
			<INPUT type="button" value="中止" onClick="window.close()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
