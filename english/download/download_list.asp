<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: download_list.asp
'  【プログラム名称】　: ダウンロード
'
'  （変更履歴）
'
'**********************************************
	Option Explicit
	Response.Expires = 0	

	dim i
	dim v_Data_Cnt1	
	dim v_Data_Cnt2
	dim Arr_GuideFiles
	dim Arr_FormFiles
	
	redim Arr_GuideFiles(0)	
	redim Arr_FormFiles(0)		
	
	call LfGetGuideFiles()
	call LfGetFormFiles()

'------------------------------
'   ガイドブックファイルを取得
'------------------------------
function LfGetGuideFiles()
	dim ObjFSO,ObjTS,myfile
	dim cnt
	dim param(2)
	
	cnt = 0
	
	call getGuideIni(param,v_Guide)			
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	Set ObjTS = ObjFSO.GetFolder(param(0) + "\en\")
	
	for each myfile in ObjTS.Files
		cnt = cnt + 1			
		redim preserve Arr_GuideFiles(cnt)  	
		Arr_GuideFiles(cnt) = myfile.Name		
	next
	
	v_Data_Cnt1 = cnt
		
	Set ObjTS = Nothing
	Set ObjFSO = Nothing	
end function

'------------------------------
'   各種様式等を取得
'------------------------------
function LfGetFormFiles()
	dim ObjFSO,ObjTS,myfile
	dim cnt
	dim param(2)
	
	cnt = 0
	
	call getGuideIni(param,v_Form)			
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set ObjTS = ObjFSO.GetFolder(param(0) + "\en\")

	for each myfile in ObjTS.Files
		cnt = cnt + 1			
		redim preserve Arr_FormFiles(cnt)  	
		Arr_FormFiles(cnt) = myfile.Name		
	next
	
	v_Data_Cnt2 = cnt
		
	Set ObjTS = Nothing
	Set ObjFSO = Nothing	
end function


'-------------------------------------
'   INIファイルからパラメータを読込む
'	Input :Array(1), Variable Name
'	Output:Array(0) = ディレクトリパス
'-------------------------------------
function getGuideIni(param,strVariable)
	dim ObjFSO,ObjTS,tmpStr
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- ファイルを開く（読み取り専用） ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("../../adminmenu/ini/admin.ini"),1,false)
	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 11, 1)
		Select Case tmpStr(0)			
			Case strVariable
				param(0) = tmpStr(1)
		End Select
	Loop
	
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function	

%>
<!--#include File="../../adminmenu/Common/common.inc"-->
<!--#include File="download.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>Download</TITLE>
<!-------------2009/07/17 Del-S Tanaka--------------------------->
<!--link href="../adminmenu/Common/style.css" rel="stylesheet" type="text/css">-->
<!-------------2009/07/17 Del-E Tanaka--------------------------->
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</HEAD>
<body bgcolor="DEE1FF" text="#000000" background="/gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" scroll="auto">
<form name="frm" action="download_list.asp" method="post">
<!-------------ここからログイン入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="../gif/download.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="/gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
          </td>
        </tr>
      </table>
      <center>
		<table border=0><tr><td height=20></td></tr></table>
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
							<table width="100%">
							<tr>
							  <td align="left" width="100%"> 												  
								  <table width="100%" cellspacing="2" cellpadding="3">
									<tr>	    							  
									<!-------------2009/07/17 Upd-S Tanaka-------------------------->
									<!--<td colspan="2" nowrap align="left" valign="middle">■ ガイドブック</td>-->
									  <td colspan="2" nowrap align="left" valign="middle"><font Color="#990000">■ GUIDANCE MANUAL</Font></td>	
									<!-------------2009/07/17 Upd-E Tanaka-------------------------->

									 </tr>
									 <tr>
									  <td width="30">&nbsp;</td>
									  <td>			  
									  <!--<div class="listbox">-->
										<table align="left" cellspacing=0 cellPadding=0>
										<% for i = 1 to UBOUND(Arr_GuideFiles) %>						
											<tr>												
											  <td>
								<a href="download.asp?guide=<%=Arr_GuideFiles(i)%>"><%= Mid(Arr_GuideFiles(i),1,Len(Arr_GuideFiles(i)) - 4)%></a>
											  </td>												
											</tr>	
										<% next %>																			
									    </table>
									  <!--</div> -->
									  </td>
									  <!--<td nowrap valign="top" width="100">(ファイル名称)</td>-->
									</tr>									
									<tr> 
									  <td border="1"></td>
									<tr> 
									<tr> 
									  <td border="1"></td>
									<tr> 
									<tr> 
									  <td border="1"></td>
									<tr> 
									<tr> 
									  <td border="1"></td>
									<tr> 
									<!-------------2009/07/17 Upd-S Tanaka-------------------------->
									<!--<td colspan="2" nowrap align="left" valign="middle">■ 各種様式　等</td>-->				  
									  <td colspan="2" nowrap align="left" valign="middle"><font Color="#990000">■ REGISTRATION FORMS</Font></td>				  
									<!-------------2009/07/17 Upd-E Tanaka-------------------------->
									</tr>
									 <tr>
									  <td width="30">&nbsp;</td>
									  <td>			  
									  <!--<div class="listbox">-->
										<table align="left" cellspacing=0 cellPadding=0>
										<% for i = 1 to UBOUND(Arr_FormFiles) %>						
											<tr>
												<td>
									<a href="download.asp?form=<%=Arr_FormFiles(i)%>"><%= Mid(Arr_FormFiles(i),1,Len(Arr_FormFiles(i)) - 4)%></a>
												</td>
											</tr>	
										<% next %>																			
									    </table>
									  <!--</div> -->
									  </td>
									  <!--<td nowrap valign="top" width="100">(ファイル名称)</td>-->
									</tr>									
									<tr> 
									  <td border="1"></td>
									<tr> 
									<tr> 
									  <td border="1"></td>
									<tr> 
									<tr> 
									  <td border="1"></td>
									<tr> 
									<tr> 
									  <td border="1"></td>
									<tr> 
								  </table>
								  <br>
								  <center>						  
								  <br>
								  <a href="javascript:window.close();">CLOSE</a>			
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
</table>
</form>
</body>
</HTML>
