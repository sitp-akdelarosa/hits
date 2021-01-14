<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: settings.asp 
'  【プログラム名称】　: 各種パラメータ設定 
'
'  （変更履歴）
'
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
%>
<!--#include File="./Common/Common.inc"-->
<!--#include File="./Common/upload.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<SCRIPT src="./Common/calpopup.js" type=text/javascript></SCRIPT>
<SCRIPT src="./Common/dateparse.js" type=text/javascript></SCRIPT>
<%
	'画面項目変数
	dim obj	
	dim buf
	dim totalByte
	dim i
	dim v_Filename
	dim v_Mode
	dim v_MailMinute
	dim v_InformDate
	dim v_Msg
	
	const lProgramID = "settings"
	
	Set obj=server.createobject("basp21")
	on error resume next
	totalByte = Request.TotalBytes
	buf	= Request.BinaryRead(totalByte)
	
	v_Mode = ""
	v_MailTime = ""
	v_InformDate = ""
	v_Msg = ""
	'----------------------------------------
    ' 再描画前の項目取得
    '----------------------------------------	
	call LfRequestItem()
	
	if v_Mode = "I" and v_Filename <> "" then		
		if not gfUploadFile(v_Filename,"txtFileUpload",v_Terminal) then
			v_Msg = "アップロードは失敗しました。"			
		end if	
	end if
	
	if v_Mode = "U" then
		call LfsetInfo()
		call LfUpdateDB()
	end if
	
	call LfgetInfo()
	
function LfRequestItem()
	v_Mode = obj.Form(buf,"Gamen_Mode")
	v_MailMinute = obj.Form(buf,"txtMailTime")
	v_InformDate = obj.Form(buf,"txtInformDate")	
	v_Filename = obj.FormFileName(buf,"txtFileUpload")		
end function 

'---------------------------------
'   mail通信間隔と強制更新日を変更
'---------------------------------
function LfsetInfo()
	dim ObjFSO,ObjTS
	dim Arr_Temp
	dim cnt
	dim strTemp
	
	cnt = 0
	redim Arr_Temp(0)	 	
   
	'--- ファイルを開く（読み取り専用） ---
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")	
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/admin.ini"),1,false)
	
	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream		
		cnt = cnt + 1
		strTemp = ObjTS.ReadLine
		redim preserve Arr_Temp(cnt)				
		if Mid(Trim(strTemp),1,18) = v_MailTime then		
			strTemp = v_MailTime & "=" & v_MailMinute
		elseif Mid(Trim(strTemp),1,23) = v_InformUser then
			strTemp = v_InformUser & "=" & v_InformDate
		end if				
		Arr_Temp(cnt) = strTemp
	Loop
	
	ObjTS.Close	
	
	'--- ファイルを開く（読み取り専用） ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/admin.ini"),2,false)
	
	for i = 1 to UBOUND(Arr_Temp)	
		ObjTS.WriteLine(Arr_Temp(i))
	next
	
	ObjTS.Close	
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

'----------------------
'   mail通信間隔を取得
'----------------------
function LfgetInfo()
	dim sql
	dim conn,rs
	'----------------------------------------
    ' ＤＢ接続
    '----------------------------------------        
    ConnectSvr conn, rs
	
	sql = "SELECT * FROM mParam WHERE Seq='1'"	
	
	rs.Open sql, conn, 0, 1, 1	

	if not rs.eof then
		v_InformDate = gfTrim(rs("ForceDate"))
		v_MailMinute = gfTrim(rs("ComInterval"))
	end if	
	
	rs.close
	conn.close	
end function

''---------------------------
''   mail通信間隔を取得
''---------------------------
'function LfgetMailTime()
'	dim param(2)			
'	call getUploadIni(param,v_MailTime)	
'	v_MailMinute = param(0)	
'end function
'
''---------------------------
''   mail通信間隔を取得
''---------------------------
'function LfgetInformDate()
'	dim param(2)		
'	call getUploadIni(param,v_InformUser)	
'	v_InformDate = param(0)	
'end function

function LfUpdateDB()
	dim sql
	dim conn,rs
	'----------------------------------------
    ' ＤＢ接続
    '----------------------------------------        
    ConnectSvr conn, rs

	sql = "SELECT Seq As Seq FROM mParam"	
	
	rs.Open sql, conn, 0, 1, 1	

	if not rs.eof then	
		sql = "UPDATE mParam"
		sql = sql & " SET "
		sql = sql & "ForceDate = '" & v_InformDate & "',"
		sql = sql & "ComInterval = '" & v_MailMinute & "'"
		sql = sql & "WHERE Seq = '1'"		
	else
		sql = "INSERT INTO mParam(Seq,UpdtPgCd,UpdtTmnl,ForceDate,ComInterval)"
		sql = sql & " VALUES ("
		sql = sql & "'1'"
		sql = sql & ",'" & lProgramID & "'"
		sql = sql & ",'manual'"
		sql = sql & ",'" & v_InformDate & "'"
		sql = sql & ",'" & v_MailMinute & "')"				
	end if
	
	conn.execute sql
	
	if err.number<>0 then				'--- エラー
		conn.rollbacktrans
		v_Msg = "変更できません。"
	end if
	
	rs.close
	conn.close	
end function

	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-各種パラメータ設定</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<link type="text/css" href="./Common/calpopup.css" media="screen" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<script type="text/javascript">
// Optionally change the date format.
	g_Calendar.setDateFormat('yyyymmdd');
</script>

<SCRIPT Language="JavaScript">
//画面項目に設定
function finit(){	
	document.frm.Gamen_Mode.value = "<%=v_Mode%>";
	document.frm.txtMailTime.value = "<%=v_MailMinute%>";
	document.frm.txtInformDate.value = "<%=v_InformDate%>";
	document.frm.txtInformDate.focus();	
}

// アップロードボタンを押下時
function fUpload(){
	if (gfCHKNull(document.frm.txtFileUpload) == false){
		document.frm.txtFileUpload.focus();
        return false;
    } 	
	
	if (gfCHKImage(document.frm.txtFileUpload) == false){
		document.frm.txtFileUpload.focus();
        return false;
    }
	
	document.frm.Gamen_Mode.value = "I";	
	document.frm.submit();
}

// 登録ボタンを押下時
function fIns(){
	if (gfCHKDate(document.frm.txtInformDate) == false){
		document.frm.txtInformDate.focus();
        return false;
    } 

	if (gfCHKNumber(document.frm.txtMailTime) == false){
		document.frm.txtMailTime.focus();
		return false;
	}
    
	document.frm.Gamen_Mode.value = "U";	
	document.frm.submit();
}

function gfCHKImage(obj){	
	var img = new Image();
	var ext = obj.value;
	img.src = obj.value;

	ext = ext.substring(ext.length-15,ext.length);
    ext = ext.toLowerCase();
	
    if (ext == "terminalmap.gif"){
		if(img.width > 431 || img.height > 272){
			//確認メッセージ
			if(confirm("画像サイズは現在のサイズと違いますがこのサイズで登録します。よろしいですか？") == false){
				return false; 
			}else{
				return true;
			}	
		}else{
			return true;
		}    	
    }
	
	
	var ext2 = obj.value;	
	ext2 = ext2.substring(ext2.length-20,ext2.length);
    ext2 = ext2.toLowerCase();
		
    if (ext2 == "terminalmap.icct.gif"){
		if(img.width > 440 || img.height > 252){
			//確認メッセージ
			if(confirm("画像サイズは現在のサイズと違いますがこのサイズで登録します。よろしいですか？") == false){
				return false; 
			}else{
				return true;
			}	
		}else{
			return true;
		}    	
    }	
	
	alert("入力値が正しくありません。");
	return false;
}
</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<form name="frm" action="settings.asp" method="post" enctype="multipart/form-data">
<SCRIPT src="./Common/KeyDown.js" type=text/javascript></SCRIPT>
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("各種パラメータ設定")
    	  %>
      </table>
      <center>
		<table border=0><tr><td height=0></td></tr></table>
        <table class="square" border="0" cellspacing="0" cellpadding="0">
          <tr>
           <td>
		  	<table border="0" cellspacing="0" cellpadding="5">
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
								  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1> 												  								  <table width="100%" border="0" cellspacing="2" cellpadding="3">
									<tr> 
									  <td colspan="3" align=left valign=middle nowrap>１．利用者ID情報更新機能</td>				  
									</tr>															
									<tr>
									  <td width="50"></td>
									  <td>強制更新日 &nbsp; 
									  <input type="text" size="15" maxlength="10" name="txtInformDate" id="txtInformDate" class="dateparse">
									  <a href="javascript:void(0);" onClick="g_Calendar.show(event, 'txtInformDate')" title="Show popup calendar">
								      <img src="./Common/calendar.gif" class="cp_img" alt="Open popup calendar"></a>							  
									  </td>
									  <td valign="top">&nbsp;</td>
									</tr>
									<tr>
									  <td colspan="3" align=Center ><font size=2>（注意）１年後にその年の同じ月日の値に自動更新されます。</font></td>
									</tr>
									<tr>
									  <td height="20"></td>
									</tr>	
									<tr> 
									  <td colspan="3" align=left valign=middle nowrap>２．香椎混雑状況情報</td>				  
									</tr>
									
									<tr> 
									  <td width="50"></td>
									  <td colspan="2" nowrap>ターミナルマップ更新</td>						    
									</tr>									
									<tr>
									  <td width="50"></td>
									  <td colspan="2" align="left">
									  	<input type="file" name="txtFileUpload" size="50" onFocus="document.frm.txtFileUpload.select();">
									  </td>									                          
									</tr>
									<tr>
									  <td width="50"></td>
									  <td colspan="2" cowrap>(注意）ファイル名：terminalmap.gif、&nbsp;サイズ：431w*272h</td>
									</tr> 														
									<tr> 									  
									  <td colspan="3" align="center">
									  	<input type="button" value=" アップロード " onClick="fUpload();">										
									  </td>  									 
									</tr>
									<tr>
									  <td height="20"></td>
									</tr>	
									<tr> 
									  <td colspan="3" align=left valign=middle nowrap>３．業務指示お知らせ機能</td>				  
									</tr>																								
									<tr>
									  <td width="50"></td>
									  <td>mail通信間隔 &nbsp; <input type="text" class="num" name="txtMailTime" size="4" maxlength="3">分</td>
									</tr>
									<tr>
									  <td height="20"></td>
									</tr>	
									<tr>									  
									  <td align="center" colspan="3">									  
									  <input type="button" value="  登録  " onClick="fIns();">
									  </td>		  
									</tr>	
								  </table>								  					  
								  <br>
								  <br>
								  <center>
								  <a href="menu.asp">閉じる</a>			
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
