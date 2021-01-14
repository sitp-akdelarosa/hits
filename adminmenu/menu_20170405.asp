<%
'***********************************************************
'  【プログラムＩＤ】　: 
'  【プログラム名称】　: 
'
'  （変更履歴）
'2017/01/19 T.Okui メニュー(10)「承認ドライバ一覧・削除」追加
'***********************************************************
Option Explicit
Response.Expires = 0

call CheckLoginH()

%>
<!--#include File="./Common/Common.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-管理者用メニュー </TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="menu" action="menu.asp" method="post">
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader
		%>
      </table>
      <center>		
		<table border=0><tr><td height=50></td></tr></table>
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
					  <td></td>		 
					  <td align="center">
					  <table>
					<tr>
					  <td nowrap align="center" class="menu">
					  <dl>
					  <B>管理者用メニュー</B>
					  </dl>
					  <center>
					  <table border="0" cellspacing="2" cellpadding="3">
						<tr> 
						  <td nowrap align=left valign=middle><a href="upload.asp">（１）様式アップロード</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="maintenance.asp">（２）お知らせメンテナンス</a></td>
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="update.asp">（３）テロップ更新</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="agreement_update.asp">（４）利用規約の更新</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="accesstotal.asp">（５）利用件数表示</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="place.asp">（６）蔵置場所コードメンテナンス</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="lockonservice.asp">（７）ロックオンサービス制限</a></td>
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="settings.asp">（８）各種パラメータ設定</a></td>
						</tr>
						<!-- 2016/07/27 H.Yoshikawa add start -->
						<tr> 
						  <td nowrap align=left valign=middle><a href="driver.asp">（９）ドライバ承認</a></td>
						</tr>
						<!-- 2016/07/27 H.Yoshikawa add end   -->
						<!-- 2017/01/20 T.Okui add start -->
						<tr> 
						  <td nowrap align=left valign=middle><a href="driverlist.asp">（10）承認ドライバ一覧・削除</a></td>
						</tr>
						<!-- 2017/01/20 T.Okui add end -->
					  </table>	
					  </center>				  					  
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
    </td>
 </tr>
    <%
		DisplayFooter
	%>  
</table>
</form>
</body>
</HTML>

