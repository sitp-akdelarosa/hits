<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "index.asp"

	Dim iLoginKind,sLoginKind
	sLoginKind = Session.Contents("userkind")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"             'メニュー画面へ
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	Dim bAryFlag
    ' 指定引数の取得
    If Request.QueryString("line")<>"" Then
  	    Dim iLine(0)
		iLine(0) = CInt(Trim(Request.QueryString("line")))
		Session.Contents("lineary") = iLine(0)
		bAryFlag = 0
	Else
	    iLine = Split(Session.Contents("lines"),",")
		Session.Contents("lineary") = Session.Contents("lines") 'ブラウザのbackボタン対策
		Session.Contents("lines") = ""
		bAryFlag = 1
	End If

	Dim iNum
	iNum = "a109"

	If sLoginKind="港運" Then
    	strTitle="空コン受取場所・搬出日"
    	WriteLog fs, iNum,"空コンピックアップシステム-空コン受取場所・搬出日変更","02", ","
	Else
    	strTitle="空コン搬出日"
    	WriteLog fs, iNum,"空コンピックアップシステム-空コン受取場所・搬出日変更","01", ","
	End If

  ' 詳細表示行のデータの取得
  If bAryFlag=0 Then
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLine(0)=LineNo Then
           Exit Do
        End If
    Loop
  End If
    ti.Close

%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function ClickSend() {
		if(!checkBlank(getFormValue(1)) && !checkBlank(getFormValue(2))){ return showAlert("<% If sLoginKind="港運" Then %>受取場所及び<% End If %>搬出日",true);}
		if(!checkDate(new getDateValue(getFormValue(2),getFormValue(3),getFormValue(4)))){
			return showAlert("空コン搬出日",false);
		}
		return true;
	}

	function getFormValue(iNum){
		formvalue = window.document.con.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}

	function getDateValue(year,mon,day){
		this.year = year;
		this.mon  = mon;
		this.day  = day;
	}

	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "が未入力です。");
		} else {
			window.alert(strAlert + "が不正です。");
		}
		return false;
	}

	function checkDate(gdv){
		if(gdv.year != "" || gdv.mon != "" || gdv.day != ""){
			if( !(gdv.year > 0 || gdv.year <= 0) || gdv.year < 2001 ) { return false; }
			if( !(gdv.mon > 0 || gdv.mon <= 0)   || (gdv.mon < 1 || gdv.mon > 12) ) { return false; }
			if( !(gdv.day > 0 || gdv.day <= 0)   || (gdv.day < 1 || gdv.day > 31) ) { return false; }
			if (gdv.day<=0 || gdv.day>30+((gdv.mon==4||gdv.mon==6||gdv.mon==9||gdv.mon==11)?0:1) || 
				(gdv.mon==2&&gdv.day>28+(((gdv.year%4==0&&gdv.year%100!=0)||gdv.year%400==0)?1:0)) ){ return false; }
		}
		return true;
	}

<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<% If sLoginKind="港運" Then %>
          <td rowspan=2><img src="gif/pickkot.gif" width="506" height="73"></td>
<% Else %>
          <td rowspan=2><img src="gif/pickrit.gif" width="506" height="73"></td>
<% End If %>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strRoute = Session.Contents("route")
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
<%' If sLoginKind="港運" Then %>
				<%'=strRoute%> &gt; 空コン受取場所・搬出日変更
<%' Else %>
				<%'=strRoute%> &gt; 空コン受取指定日変更
<%' End If %>
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
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>
<% If sLoginKind="港運" Then %>
			空コン受取場所・搬出日変更
<% Else %>
			空コン受取指定日変更
<% End If %>
			</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr>
<% If sLoginKind="港運" Then %>
          <td>変更する項目のみ値を入力して、送信ボタンを押して下さい。</td>
<% Else %>
          <td>変更する日付を入力して、送信ボタンを押して下さい。</td>
<% End If %>
        </tr>
      </table>
      <FORM NAME="con" METHOD="post" action="picklist-input-syori.asp" onSubmit="return ClickSend()">
		<input type=hidden name=title value="<%=strTitle%>">
        <table border=0 cellpadding=0 bordercolor="#999999">
          <tr> 
            <td align="center"> 
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">

<% If sLoginKind="港運" Then %>
                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                     空コン受取場所</font></b></td>
                  <td bgcolor="#FFFFFF"> 
					<table border=0 cellpadding=0 cellspacing=0 width=100%>
					  <tr>
						<td nowrap>
<% ' 空コン受取場所
 	If bAryFlag=0 Then
	    Response.Write "<input type=text name='pickplace' value='" & anyTmp(20) & "' size=22 maxlength=20>"
	Else
	    Response.Write "<input type=text name='pickplace' size=22 maxlength=20>"
	End If
%>
						</td>
						<td nowrap align=right>
							<font size=1 color="#2288ff">[日本語入力可]</font><BR>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<% Else %>
				<input type=hidden name=dammy value="">
<% End If %>

                <tr> 
                  <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
<% If sLoginKind="港運" Then %>
                    空コン搬出日</font></b></td>
<% Else %>
                    空コン受取指定日</font></b></td>
<% End If %>
                  <td> 
					<table border=0 cellpadding=0 cellspacing=0 width=100%>
					  <tr>
						<td nowrap>
<%
 	If bAryFlag=0 Then
	    strTemp=anyTmp(24)
	    If strTemp="" Then
	        strTemp=DispDateTime(Now,0)
	    End If
	    Response.Write "<input type=text name='pickyear' value='" & Left(strTemp,4) & "' size=4 maxlength='4'>年"
	    Response.Write "<input type=text name='pickmon' value='" & Mid(strTemp,6,2) & "' size=2 maxlength='2'>月"
	    Response.Write "<input type=text name='pickday' value='" & Mid(strTemp,9,2)  & "' size=2 maxlength='2'>日"
	Else
'	    Response.Write "<input type=text name='pickyear' value='" & Year(Now) & "' size=4 maxlength='4'>年"
'	    Response.Write "<input type=text name='pickmon' value='" & Month(Now) & "' size=2 maxlength='2'>月"
	    Response.Write "<input type=text name='pickyear' size=4 maxlength='4'>年"
	    Response.Write "<input type=text name='pickmon' size=2 maxlength='2'>月"
	    Response.Write "<input type=text name='pickday' size=2 maxlength='2'>日"
	End If
%>
					<BR>&nbsp;&nbsp;&nbsp;<font size=-1>（例） 2002年 2月 25日</font>
						</td>
						<td width=10></td>
						<td nowrap align=right>
<% If sLoginKind="陸運" Then %>
							<font size=1 color="#ff0000">[必須入力]</font> <BR>
<% End If %>
							<font size=1 color="#2288ff">[半角数値]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
              </table>
              <br>
              <input type=submit value="　送信　">
              <input type="button" value="　中止　" onclick="history.back()">
            </td>
          </tr>
        </table>
      </form>
      <br>
      <br>
      <br>
      <br>
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
<!-------------登録画面終わり--------------------------->
<%
	If sLoginKind="港運" Then
	    DispMenuBarBack "picklist.asp?kind=4"
	Else
	    DispMenuBarBack "picklist.asp?kind=2"
	End If
%>
</body>
</html>
