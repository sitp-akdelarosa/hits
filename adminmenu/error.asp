<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:error.asp				_/
'_/	Function	:エラー画面				_/
'_/	Date		:2003/06/18				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'エラー情報取得
  dim ObjConn, ObjRS,WinFlag,dispId,wkID,wkName,errerCd,etc
  WinFlag= Session.Contents("WinFlag")
  dispId = Session.Contents("dispId")
  wkID   =  Session.Contents("wkID")
  wkName =  Session.Contents("wkName")
  errerCd=  Session.Contents("errerCd")
  etc    =  Session.Contents("etc")
'セッションクリア
  Session.Contents.Remove("WinFlag")
  Session.Contents.Remove("dispId")
  Session.Contents.Remove("wkID")
  Session.Contents.Remove("wkName")
  Session.Contents.Remove("errerCd")
  Session.Contents.Remove("etc")

'エラーメッセージ取得
  dim ErrerM1,ErrerM2
  dim ObjFSO,ObjTS,tmpStr,tmp
  ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
  ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/ADMINERROR.ini"),1,false)
  '--- ファイルデータの読込み ---
  Do Until ObjTS.AtEndofStream
    tmpStr = ObjTS.ReadLine
    If Left(tmpStr,3) = errerCd Then
      tmp=Split(tmpStr,":",3,1)
      ErrerM1 = tmp(1)
      ErrerM2 = tmp(2)
      Exit Do
    End If
  Loop
  ObjTS.Close
  ObjTS = Nothing
  ObjFSO = Nothing

'ボタン表示制御
  dim Button
  If WinFlag = 0 Then
    Button="'ログイン画面に戻る' onClick='submit()'"
  ElseIf WinFlag = 1 Then
    Button="'閉じる' onClick='window.close()'"
  Else
    Button="'戻る' onClick='window.history.back()'"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>エラー</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
// -->
</SCRIPT>
<!--#include File="./Common/common.inc"-->
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY class="bckcolor" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------エラー画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			DisplayHeader2("お知らせメンテナンス")
		%>
		<INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
    	<INPUT type="hidden" name="Data_Cnt" size="9" readonly tabindex= -1>
      </table>
      <center>				
        <table class="square" border="0" cellspacing="4" cellpadding="0">
          <tr>
           <td>
		  	<table border="0" cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
				<table width="500" border="0" cellspacing="0" cellpadding="5">
				  <tr>
				   <td>
					  <table width=100%>	               
						<tr>						  		 
						  <td align="center">
							<table width="100%" border=0>
							<tr>
							  <td align="left" width="100%"> 
								  <table width="100%" border="0" cellspacing="2" cellpadding="3">									
									<tr><td colspan=2 align="center" class="menu">エラー</td></tr>
									<tr><td>エラー画面ID：作業ID</td><td>：<%=dispId%>：<%=wkId%></td></tr>
									<tr><td>作業名</TD><TD>：<%=wkName%></td></tr>
									<tr><td>エラーコード</TD><TD>：<%=errerCd%></td></tr>
									<tr><td>メッセージ</TD><TD>：<%=ErrerM1%><BR></td></tr>
									<tr><td>対処</td><td>：<%=ErrerM2%><BR></td></tr>
									<tr><td colspan=2><%=etc%></td></tr>
									<tr><td colspan=2 height="20"></td></tr>
									<tr><td colspan=2 align=center>
									<form action="./login.asp" target="_top">										
										<input type=button value=<%=Button%>>
									</form>
									</td></tr>
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
	  	    </td>
    	   </tr>
      	  </table>
	  	 </td>
        </tr>
      </table>
	  </center>	  
	  <table border=0><tr><td height=20></td></tr></table>
    </td>	
 </tr> 
 	<%
		DisplayFooter
	%> 
</table>
<!-------------画面終わり--------------------------->
</BODY></HTML>
