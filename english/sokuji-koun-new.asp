<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	即時搬出システム【港運用】	変更,削除用画面

%>

<%
	' セッションのチェック
	CheckLogin "sokuji.asp"

	' 港運コード取得
	sOpe = Trim(Session.Contents("userid"))

	Dim sTimeset,sCorrfail
	sTimeset = Trim(Request.form("timeset"))
	sCorrfail = Trim(Request.form("corrfail"))

	If sTimeset<>"" Then
		Session.Contents("send") = 0
	Else
		Session.Contents("send") = 1
	End If

	Dim sShipper(),sShipLine(),sVslCode(),sBL(),sCont(),sForwarder(),sLineNo(),iChkCount

	' セッションのチェック
	CheckLogin "sokuji-koun-list.asp"
	iChkCount=Session.Contents("ChkCount")
	Session.Contents("ChkCount")=iChkCount

	For i=1 to iChkCount
		Session.Contents("shipper" & i)=Request.form("shipper" & i)
		Session.Contents("shipline" & i)=Request.form("shipline" & i)
		Session.Contents("vslcode" & i)=Request.form("vslcode" & i)
		Session.Contents("bl" & i)=Request.form("bl" & i)
		Session.Contents("cont" & i)=Request.form("cont" & i)
		Session.Contents("forwarder" & i)=Request.form("forwarder" & i)
		Session.Contents("lineno" & i)=Request.form("chk" & i)
'		ReDim Preserve sShipper(i)
'		ReDim Preserve sShipLine(i)
'		ReDim Preserve sVslCode(i)
'		ReDim Preserve sBL(i)
'		ReDim Preserve sCont(i)
'		ReDim Preserve sForwarder(i)
'		ReDim Preserve sLineNo(i)
'		sShipper(i) 	= Request.form("shipper" & i)
'		sShipLine(i)	= Request.form("shipline" & i)
'		sVslCode(i) 	= Request.form("vslcode" & i)
'		sBL(i) 				= Request.form("bl" & i)
'		sCont(i) 			= Request.form("cont" & i)
'		sForwarder(i)	= Request.form("forwarder" & i)
'		sLineNo(i)		= Request.form("chk" & i)
	Next


	If sCorrfail<>"" Then
		Response.Redirect "sokuji-koun-exec.asp"
		Response.End
	End If

%>
<html>
<head>
<title>即時搬出申込み（港運）</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function checkFormValue(){
		if(!checkBlank(getFormValue(0))){ return showAlert("日付",true); }
		if(!checkBlank(getFormValue(1))){ return showAlert("日付",true); }
		if(!checkBlank(getFormValue(2))){ return showAlert("日付",true); }
		if(!checkBlank(getFormValue(3))){ return showAlert("時間",true); }
		if(!checkBlank(getFormValue(4))){ return showAlert("時間",true); }
		if((Number(getFormValue(1))<1)||(Number(getFormValue(1))>12)) { return showAlert("月は1～12",false);}
		if((Number(getFormValue(2))<1)||(Number(getFormValue(2))>31)) { return showAlert("日は1～31",false);}
		if((Number(getFormValue(3))<0)||(Number(getFormValue(3))>23)) { return showAlert("時は0～23",false);}
		if((Number(getFormValue(4))<0)||(Number(getFormValue(4))>59)) { return showAlert("分は0～59",false);}
		return true;
	}
	function getFormValue(iNum){
		formvalue = window.document.input.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}
	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "が未入力です。");
		} else {
			window.alert(strAlert + "の範囲で入力してください。");
		}
		return false;
	}
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここからログイン入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/sokuji2t.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'	DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>

<!-- commented by seiko-denki 2003.07.18
	  <table width=95% cellpadding="0" cellspacing="0" border="0">
		<tr>
		  <td align="right">
			<font color="#333333" size="-1">
				<%=strRoute%>
			</font>
		  </td>
		</tr>
	  </table>
End of comment by seiko-denki 2003.07.18 -->
	  <BR>
	  <BR>
	  <BR>
      <table>
        <tr>
          <td nowrap align=center> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>（港運用）即時搬出申込み</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>

            <br>
			<table border=0 cellpadding=0 cellspacing=0>
			<tr><td nowrap>
			容器通関の時刻を入力します。<BR>
			搬入確認予定時刻を入力して、送信ボタンを押して下さい。
			</td></tr>
			</table>

            <form method=post name="input" action="sokuji-koun-exec.asp">
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>搬入確認<BR>予定時刻</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td nowrap>
							<input type=text name=year value="<%=Year(Now)%>" size=5 maxlength=4>年
							<input type=text name=month value="<%=Month(Now)%>" size=3 maxlength=2>月
							<input type=text name=day value="<%=Day(Now)%>" size=3 maxlength=2>日
							&nbsp;&nbsp;
							<input type=text name=hour size=3 maxlength=2>時
							<input type=text name=min size=3 maxlength=2>分
							<BR>
							&nbsp;&nbsp;&nbsp;
							（例） 2002年2月25日 15時30分
							&nbsp;&nbsp;&nbsp;
							<font size=1 color="#2288ff">[ 半角数値 ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

              </table>
              <br>

              <input type=submit name="send" value=" 送  信 " onClick="return checkFormValue()">
              <input type=submit name="stop" value=" 中  止 ">

              </center>
            </form>
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
<!-------------ログイン画面終わり--------------------------->
<%
    DispMenuBarBack "sokuji-koun-list.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    WriteLog fs, "7004", "即時搬出システム-港運用予定時刻入力", "00", ""
%>
