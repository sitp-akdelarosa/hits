<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ＴＯＰ画面
    WriteLog fs, "ＴＯＰ画面", ""
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
<!--
if(navigator.appVersion.charAt(0) >= "3") {
    IMG01on = new Image(177,19);
    IMG01on.src = "gif/i-b1-22.gif";
    IMG01off = new Image(177,19);
    IMG01off.src = "gif/i-b1-2.gif";

    IMG02on = new Image(178,19);
    IMG02on.src = "gif/i-b1-32.gif";
    IMG02off = new Image(178,19);
    IMG02off.src = "gif/i-b1-3.gif";

    IMG03on = new Image(236,23);
    IMG03on.src = "gif/i-b22.gif";
    IMG03off = new Image(236,23);
    IMG03off.src = "gif/i-b2.gif";

    IMG04on = new Image(215,18);
    IMG04on.src = "gif/i-b3-12.gif";
    IMG04off = new Image(215,18);
    IMG04off.src = "gif/i-b3-1.gif";

    IMG05on = new Image(224,21);
    IMG05on.src = "gif/i-b3-22.gif";
    IMG05off = new Image(224,21);
    IMG05off.src = "gif/i-b3-2.gif";

    IMG06on = new Image(146,20);
    IMG06on.src = "gif/i-b3-32.gif";
    IMG06off = new Image(146,20);
    IMG06off.src = "gif/i-b3-3.gif";

    IMG07on = new Image(46,20);
    IMG07on.src = "gif/i-b4-12.gif";
    IMG07off = new Image(46,20);
    IMG07off.src = "gif/i-b4-1.gif";

    IMG08on = new Image(93,19);
    IMG08on.src = "gif/i-b4-22.gif";
    IMG08off = new Image(93,19);
    IMG08off.src = "gif/i-b4-2.gif";

    IMG09on = new Image(88,19);
    IMG09on.src = "gif/i-b4-32.gif";
    IMG09off = new Image(88,19);
    IMG09off.src = "gif/i-b4-3.gif";

    IMG10on = new Image(166,23);
    IMG10on.src = "gif/i-b52.gif";
    IMG10off = new Image(166,23);
    IMG10off.src = "gif/i-b5.gif";
}
function slideonchg(imgname) {
    if(navigator.appVersion.charAt(0) >= "3") {
        document[imgname].src = eval(imgname + "on.src");
    } else { }
}
function slideoffchg(imgname) {
    if(navigator.appVersion.charAt(0) >= "3") {
        document[imgname].src = eval(imgname + "off.src");
    } else { }
}
//-->
</script>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/indexback.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/indextitle.gif" width="570" height="110"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/1.gif"></td>
        </tr>
          <td align="right" width="100%" height="85">
<%
    DispMenu
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td valign="top" align="center"> 
      <table  border="0" cellspacing="0" cellpadding="0" height="100%" width=100%>
        <tr>    
          <td align="right" valign="middle"> 
            <!----------------------------topボタン-------------------------------------------->
            <table border="0" cellspacing="0" cellpadding="2">
              <tr> 
                <td colspan="2"><img src="gif/i-b1.gif" width="164" height="23"></td>
              </tr>
              <tr> 
                <td width="40" align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=expentry.asp" onmouseover="slideonchg('IMG01')" onmouseout="slideoffchg('IMG01')"><img src="gif/i-b1-2.gif" width="177" height="19"name="IMG01" naturalsizeflag="0" border="0"></a></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=impentry.asp" onmouseover="slideonchg('IMG02')" onmouseout="slideoffchg('IMG02')"><img src="gif/i-b1-3.gif" width="178" height="19" name="IMG02" naturalsizeflag="0" border="0"></a></td>
              </tr>
              <tr> 
                <td colspan="2"><a href="userchk.asp?link=terminal.asp" onmouseover="slideonchg('IMG03')" onmouseout="slideoffchg('IMG03')"><img src="gif/i-b2.gif" width="236" height="23" vspace="4" border="0" name="IMG03" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
                <td colspan="2"><img src="gif/i-b3.gif" width="254" height="23"></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=hits.asp" onmouseover="slideonchg('IMG04')" onmouseout="slideoffchg('IMG04')"><img src="gif/i-b3-1.gif" width="215" height="18" border="0" name="IMG04" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=gate.asp" onmouseover="slideonchg('IMG05')" onmouseout="slideoffchg('IMG05')"><img src="gif/i-b3-2.gif" width="224" height="21" border="0" name="IMG05" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=sokuji.asp" onmouseover="slideonchg('IMG06')" onmouseout="slideoffchg('IMG06')"><img src="gif/i-b3-3.gif" width="146" height="20" border="0" name="IMG06" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
                <td colspan="2"><img src="gif/i-b4.gif" width="128" height="23" vspace="4"></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=nyuryoku-in1.asp" onMouseOver="slideonchg('IMG07')" onMouseOut="slideoffchg('IMG07')"><img src="gif/i-b4-1.gif" width="46" height="20" border="0" name="IMG07" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=nyuryoku-kaika.asp" onMouseOver="slideonchg('IMG08')" onMouseOut="slideoffchg('IMG08')"><img src="gif/i-b4-2.gif" width="93" height="19" border="0" name="IMG08" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
                <td align="right"><img src="gif/i-b.gif" width="11" height="15"></td>
                <td><a href="userchk.asp?link=nyuryoku-te.asp" onMouseOver="slideonchg('IMG09')" onMouseOut="slideoffchg('IMG09')"><img src="gif/i-b4-3.gif" width="88" height="19" border="0" name="IMG09" naturalsizeflag="0"></a></td>
              </tr>
              <tr> 
<!--
                <td colspan="2"><a href="userchk.asp?link=request.htm" onmouseover="slideonchg('IMG10')" onmouseout="slideoffchg('IMG10')"><img src="gif/i-b5.gif" width="166" height="23" vspace="4" name="IMG10" naturalsizeflag="0" border="0"></a></td>
-->
                <td colspan="2"><a href="userchk.asp?link=request.asp" onmouseover="slideonchg('IMG10')" onmouseout="slideoffchg('IMG10')"><img src="gif/i-b5.gif" width="166" height="23" vspace="4" name="IMG10" naturalsizeflag="0" border="0"></a></td>
              </tr>
            </table>
            <!----------------------------topボタン終わり-------------------------------------------->
          </td>
          <td valign="bottom" align="center"><br>
            　<br>　<br>
            <img src="gif/index.gif" width="370" height="370"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td valign="bottom" height="20"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="000099" height="15"><img src="gif/1.gif"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!-------------登録画面終わり--------------------------->
</body>
</html>
