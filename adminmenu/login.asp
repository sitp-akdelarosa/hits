<%
'**********************************************
'  【プログラムＩＤ】　: login.asp
'  【プログラム名称】　: ログイン
'
'  （変更履歴）
'
'**********************************************
Option Explicit
Response.Expires = 0
%>

<!--#include File="./Common/common.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<%
    '--- 変数宣言 ---
    dim wArgc                   ' パラメータ最大数
    dim wRtnB                   ' 関数戻り値判定用
    dim msg
	dim conn, rs
	dim sql
	
	msg = ""
    
	session("username") = empty
	
	' 指定引数の取得(ユーザーＩＤ)
    Dim strInputUserID, strInputPassWord
    strInputUserID = UCase(Trim(Request.Form("txtUserID")))
    strInputPassWord = UCase(Trim(Request.Form("txtPass")))
	
	If strInputUserID <> "" and strInputPassWord <> "" then
       'session("Loginid") = strInputUserID
        '----------------------------------------
        ' ＤＢ接続
        '----------------------------------------        
        ConnectSvr conn, rs

        '----------------------------------------
        ' ユーザ情報取得
        '----------------------------------------
        session("user_id")   = empty
        session("username") = empty

        sql="SELECT FullName,UserType FROM mUsers WHERE UserCode = '" & gfSQLEncode(strInputUserID) & "' And Password = '" & gfSQLEncode(strInputPassWord) & "' AND UserType = '0'"
        rs.Open sql, conn, 0, 1, 1
		
		on error resume next
		
        If rs.eof or err.number<>0 then
            msg="入力された内容に間違いがあります。"
        Else
			' ログイン名をセッション変数に設定
			session("username") = Trim(rs("FullName"))            
			session("user_id") = strInputUserID            '2016/07/28 H.Yoshikawa Add
        End If
		
        rs.close
		conn.Close
    End If
	
    If session("username") <> "" then
        response.redirect "menu.asp"
    End If    
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-管理者用画面</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<SCRIPT Language="JavaScript">

function finit(){
    document.frm.txtUserID.focus();
}


function Check(){
	var obj = document.frm;  

	ret = CheckEisuji(obj.txtUserID.value);
  
	if(ret == false){
    	alert("管理者IDは半角英数字で入力してください。");
		obj.txtUserID.focus();
	    return false;
	}
	
	if(obj.txtUserID.value == ""){
    	alert("必須入力項目です。");
		obj.txtUserID.focus();
	    return false;	
	}
	
	if(obj.txtPass.value == ""){
    	alert("必須入力項目です。");
		obj.txtPass.focus();
	    return false;	
	}
	
    return true;
}
</SCRIPT>
</HEAD>

<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<form name="frm" action="login.asp" method="post">
<SCRIPT src="./Common/KeyDown.js" type=text/javascript></SCRIPT>
<!-------------ここからログイン入力画面--------------------------->
<table class="main" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			DisplayHeader
		%>
      </table>
      <center>
	  	<BR><BR>
		<table border=0><tr><td height=50></td></tr></table>
        <table class="square" cellspacing="4" cellpadding="0">
          <tr>
           <td>
		  	<table border="0" cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
	        <table width="500" border="0" cellspacing="0" cellpadding="5">
	          <tr>
	           <td>
	              <table width="100%">	                
					<tr>
			   		<td></td>		 
	                <td align="center">
					<table width="100%">
						<tr>
							<td align="center"><B>管理者用画面へのログイン。</B></td>
						</tr>
						<tr>
						  <td nowrap align="center"> 
						  	  <BR>															  
							  <table border="0" cellspacing="2" cellpadding="3">
								<tr> 
								  <td nowrap align=left valign=middle><B>管理者ID</B></td>
								  <td nowrap>
									<table border=0 cellpadding=0 cellspacing=0>
									  <tr>
										<td width=100>
											<input type=text name="txtUserID" value="" size=10 maxlength=5>
										</td>										
									  </tr>
									</table>
								  </td>
								</tr>
								<tr> 
								  <td nowrap align=left valign=middle><B>パスワード</B></td>
								  <td nowrap>
									<table border=0 cellpadding=0 cellspacing=0>
									  <tr>
										<td width=100>
											<input type=password name="txtPass" size=10 maxlength=8>
										</td>									
									  </tr>
									</table>
								  </td>
								</tr>
								<tr>
									<td colspan="2">
									<%  if msg<>"" then%>
			   						  <font color="red"><%=msg%></font>
		    					    <%  end if%>
	  							    </td>
								</tr>
							  </table>
						  	  <br>
						  	  <input type="submit" value=" ログイン " onClick="return Check();">
						 </td>
						</tr>
					</table>
		 			</td>
		  			<td></td>
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
	  </center>
    </td>
 </tr>
 	<%
		DisplayFooter
	%>
</table>
</form>
</body>
</HTML>
