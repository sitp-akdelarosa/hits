<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: upload.asp
'  【プログラム名称】　: 様式アップロード
'
'  （変更履歴）
'
'**********************************************
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()

%>
<!--#include File="./Common/common.inc"-->
<!--#include File="./Common/upload.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<%
	'画面項目変数
	dim obj
	dim i
	dim buf
	dim v_Mode
	dim v_Mode2
	dim v_Msg
	dim v_Filename
	dim v_Filename2
	dim totalByte
	dim v_Data_Cnt
	dim v_Data_Cnt2	
	dim v_ItemName
	dim wkfilename	
	dim v_FocusItem
	dim v_SubmitForm
	dim Arr_Files
	dim Arr_FormFiles
	dim Arr_DelFlag
	dim Arr_DelFlag2
	
	redim Arr_Files(0)
	redim Arr_FormFiles(0)
	redim Arr_DelFlag(0)	
	redim Arr_DelFlag2(0)
	
	Set obj=server.createobject("basp21")
	on error resume next
	totalByte = Request.TotalBytes
	buf	= Request.BinaryRead(totalByte)
	
	v_Mode = ""
	v_Mode2 = ""
	v_Data_Cnt = 0
	v_Data_Cnt2 = 0

	'----------------------------------------
    ' 再描画前の項目取得
    '----------------------------------------	
	call LfRequestItem()  
	
	if v_SubmitForm = "1" then
		if v_Mode = "I" and v_Filename <> "" then		
			if not gfUploadFile(v_Filename,"txtFileUpload",v_Guide) then
				v_Msg = "アップロードは失敗しました。"			
			end if	
		end if
	
		if v_Mode = "D" then
			call LfDeleteFile()		
		end if
	else
		if v_Mode2 = "I" and v_Filename2 <> "" then		
			if not gfUploadFile(v_Filename2,"txtFileUpload2",v_Form) then
				v_Msg = "アップロードは失敗しました。"			
			end if			
		end if	

		if v_Mode2 = "D" then		
			call LfDeleteFormFile()		
		end if
	end if
	
	call LfGetFiles()	
	
'-----------------------------
'   描画前の画面項目を取得
'-----------------------------
function LfRequestItem()
	v_SubmitForm = obj.Form(buf,"frmSubmit")	
	
	if v_SubmitForm = 1 then
		v_Mode = obj.Form(buf,"Gamen_Mode")	
		v_Data_Cnt = obj.Form(buf,"Data_Cnt")
		v_Filename = obj.FormFileName(buf,"txtFileUpload")
		if CInt(v_Data_Cnt) > 0 then
			redim Arr_Files(v_Data_Cnt)		
			redim Arr_DelFlag(v_Data_Cnt)	
	
			for i = 1 to CInt(v_Data_Cnt)					
				Arr_DelFlag(i) = obj.Form(buf,"DelFlag" & i)			
				Arr_Files(i) = obj.Form(buf,"FileNames" & i)
			next			 	
		end if
	else
		v_Mode2 = obj.Form(buf,"Gamen_Mode2")
		v_Data_Cnt2 = obj.Form(buf,"Data_Cnt2")	
		v_Filename2 = obj.FormFileName(buf,"txtFileUpload2")
		
		if CInt(v_Data_Cnt2) > 0 then		
			redim Arr_FormFiles(v_Data_Cnt2)		
			redim Arr_DelFlag2(v_Data_Cnt2)	
	
			for i = 1 to CInt(v_Data_Cnt2)					
				Arr_DelFlag2(i) = obj.Form(buf,"DelFlag2" & i)			
				Arr_FormFiles(i) = obj.Form(buf,"FileNames2" & i)
			next			 	
		end if
	end if
end function

'-----------------------------
'   (ファイル)を取得
'-----------------------------
function LfGetFiles()
	dim ObjFSO,ObjTS,myfile
	dim cnt
	dim param(2)
	
	cnt = 0
	
	call getUploadIni(param,v_Guide)			
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set ObjTS = ObjFSO.GetFolder(param(0))
	
	redim Arr_Files(0)	'2009/07/22 C.Pestano Add
	
	for each myfile in ObjTS.Files		
		cnt = cnt + 1
		redim preserve Arr_Files(cnt)  	
		Arr_Files(cnt) = myfile.Name		
	next
	
	v_Data_Cnt = cnt
	
	cnt = 0
	
	call getUploadIni(param,v_Form)			
	
	Set ObjTS = ObjFSO.GetFolder(param(0))
	
	redim Arr_FormFiles(0)	'2009/07/22 C.Pestano Add
	
	for each myfile in ObjTS.Files				
		cnt = cnt + 1	
		redim preserve Arr_FormFiles(cnt)  	
		Arr_FormFiles(cnt) = myfile.Name				
	next
	
	v_Data_Cnt2 = cnt	

	Set ObjTS = Nothing
	Set ObjFSO = Nothing	
end function

'-----------------------------
'   削除(ファイル)
'-----------------------------
function LfDeleteFile()
	dim ObjFSO,ObjTS
	dim strDir
	dim wkfilename
	dim param(2)		
		
	call getUploadIni(param,v_Guide)
	strDir = param(0)
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")	
	
	for i = 1 to UBOUND(Arr_DelFlag)
		if Arr_DelFlag(i) = "1" and strDir <> "" then		
			wkfilename = strDir & Arr_Files(i)
			gfDeleteFile(wkfilename)
		end if
	next
	
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing	
end function

'-----------------------------
'   削除(ファイル)
'-----------------------------
function LfDeleteFormFile()
	dim ObjFSO,ObjTS
	dim strDir
	dim wkfilename
	dim param(2)		
		
	call getUploadIni(param,v_Form)
	strDir = param(0)
		
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")		
		
	for i = 1 to UBOUND(Arr_DelFlag2)
		if Arr_DelFlag2(i) = "1" and strDir <> "" then					
			wkfilename = strDir & Arr_FormFiles(i)			
			gfDeleteFile(wkfilename)
		end if
	next
			
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing	
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-様式アップロード</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

//画面項目に設定
function finit(){	
	document.frm.Gamen_Mode.value = "<%=v_Mode%>";
	document.frm.Data_Cnt.value = "<%=v_Data_Cnt%>";
	document.frm2.Gamen_Mode2.value = "<%=v_Mode2%>"; 
	document.frm2.Data_Cnt2.value = "<%=v_Data_Cnt2%>"; 		
	
	if (document.frm.Gamen_Mode.value == "I"){
        // エラー等のメッセージとフォーカス
        if ("<%=v_Msg%>" != ""){
            alert("<%=v_Msg%>");

            //フォーカス位置設定
            for( var i=0; i < document.frm.elements.length; i++ ){
                 if ((document.frm.elements[i].type == "file") &&
                     document.frm.elements[i].name == "<%=v_ForcusItem%>"){
                     document.frm.elements[i].focus();  
                     return false;
                 }    
            }
            return false;
		}
	}else{
		document.frm.txtFileUpload.focus();  	
	}
	
	if (document.frm2.Gamen_Mode2.value == "I"){
        // エラー等のメッセージとフォーカス
        if ("<%=v_Msg%>" != ""){
            alert("<%=v_Msg%>");

            //フォーカス位置設定
            for( var i=0; i < document.frm2.elements.length; i++ ){
                 if ((document.frm2.elements[i].type == "file") &&
                     document.frm2.elements[i].name == "<%=v_ForcusItem%>"){
                     document.frm2.elements[i].focus();  
                     return false;
                 }    
            }
            return false;
		}
	}	
}

// アップロードボタンを押下時
function fUpload(){
	if (gfCHKNull(document.frm.txtFileUpload) == false){
		document.frm.txtFileUpload.focus();
        return false;
    } 
	
	document.frm.Gamen_Mode.value = "I"
	document.frm.frmSubmit.value = "1"
	document.frm.submit();
}

// 削除ボタンを押下時
function fDel(){
	var i,cnt;
	var obj;
	
	cnt = 0;
	
	for(i = 1; i <= "<%=v_Data_Cnt%>"; i++){
		obj = eval("document.frm.DelFlag" + i);
        if (obj.value == "1") {  
            cnt++;
        }
    }
	
    if(cnt == 0) {
        window.alert("ファイルを選択してください。");
        return false;
    }

	document.frm.Gamen_Mode.value = "D"
	document.frm.frmSubmit.value = "1"
	document.frm.submit();
}

// アップロードボタンを押下時
function fUpload2(){
	if (gfCHKNull(document.frm2.txtFileUpload2) == false){
		document.frm2.txtFileUpload2.focus();
        return false;
    } 
	
	document.frm2.Gamen_Mode2.value = "I"
	document.frm2.frmSubmit.value = "2"
	document.frm2.submit();
}

// 削除ボタンを押下時
function fDel2(){
	var i,cnt;
	var obj;
	
	cnt = 0;
	
	for(i = 1; i <= "<%=v_Data_Cnt2%>"; i++){
		obj = eval("document.frm2.DelFlag2" + i);
        if (obj.value == "1") {  
            cnt++;
        }
    }
	
    if(cnt == 0) {
        window.alert("ファイルを選択してください。");
        return false;
    }

	document.frm2.Gamen_Mode2.value = "D"
	document.frm2.frmSubmit.value = "2"
	document.frm2.submit();	
}

// ユーザがファイルを選択するとき、<TD>クラスを変えます。
// 削除フラグはつけられています。
function fHighlight(obj,delflag){
	if(obj.className == "chrReadOnly2"){
		obj.className = "highlight"
		delflag.value = "1"
	}else{
		obj.className = "chrReadOnly2"
		delflag.value = "0"
	}	
}
</SCRIPT>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">	
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("様式アップロード")
		%>
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
					  <table width="100%">	    				
						<tr>						  
						  <td align="center">
							<table width="100%">
							<tr>
							  <td align="left">
							  <form name="frm" action="upload.asp" method="post" enctype="multipart/form-data">									  
								  <table width="100%" border="0" cellspacing="2" cellpadding="3">
									<tr> 
									  <td nowrap align="left" colspan="2">●ガイドブック</td>
									  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
									  <INPUT type="hidden" name="Data_Cnt" size="9" readonly tabindex= -1>
									  <INPUT type="hidden" name="frmSubmit" size="9" readonly tabindex= -1>
									</tr>
									<tr>    								  
									  <td width="600">
									    <div class="listbox">																				
										<table width="100%" align="left" border="0" cellspacing="0" cellpadding="0">
										<% for i = 1 to UBOUND(Arr_Files) %>						
											<tr>												
											<td>
											<% v_ItemName = "DelFlag" & cstr(i) %>
											<input type="hidden" name="<%= v_ItemName %>" size="2">																							
											<% v_ItemName = "FileNames" & cstr(i) %>
											<input type="text" name="<%= v_ItemName %>" value="<%=Arr_Files(i)%>" class="chrReadOnly2" size="58" readonly=TRUE tabindex = -1 onClick="fHighlight(this,document.frm.DelFlag<%=cstr(i)%>);">										
											</td>		 
											</tr>	
										<% next %>																			
										</table>										
										</div> 
									  </td>									  
									  <td valign="bottom">										
										<input type="button" name="btnDel" value="   削除  " onClick="fDel();">
									  </td>
									</tr>	
									<tr> 
									  <td height="15" colspan="2"></td>				  
									</tr>
									<tr> 
									  <td nowrap align="left" colspan="2">●ファイルアップロード</td>				  
									</tr>						
									<tr>
									  <td colspan="2"><input type="file" name="txtFileUpload" size="60"></td>
									</tr>
									<tr>
									  <td align="center" colspan="2"><input type="button" name="btnUpload" value=" アップロード " onClick="fUpload();"></td>
									</tr>		
								  </table>
							</form>
							<form name="frm2" action="upload.asp" method="post" enctype="multipart/form-data">
								  <table width="100%" border="0" cellspacing="2" cellpadding="3">
									<tr> 
									  <td nowrap align="left" colspan="2">●各種様式等</td>
									  <INPUT type="hidden" name="Gamen_Mode2" size="9" maxlength="1"  readonly tabindex= -1>
									  <INPUT type="hidden" name="Data_Cnt2" size="9" readonly tabindex= -1>
									  <INPUT type="hidden" name="frmSubmit" size="9" readonly tabindex= -1>
									</tr>
									<tr> 
									  <td width="600">										
										<div class="listbox">
										<table width="100%" align="left" border="0" cellspacing="0" cellpadding="0">
										<% for i = 1 to UBOUND(Arr_FormFiles) %>						
											<tr>												
											<td>
											<% v_ItemName = "DelFlag2" & cstr(i) %>
											<input type="hidden" name="<%= v_ItemName %>" size="2">																							
											<% v_ItemName = "FileNames2" & cstr(i) %>
											<input type="text" name="<%= v_ItemName %>" value="<%=Arr_FormFiles(i)%>" class="chrReadOnly2" size="58" readonly=TRUE tabindex = -1 onClick="fHighlight(this,document.frm2.DelFlag2<%=cstr(i)%>);">										
											</td>		 
											</tr>	
										<% next %>																			
										</table>
										</div> 
									  </td>
									  <td valign="bottom">										
										<input type="button" name="btnDel2" value="   削除  " onClick="fDel2();">
									  </td>
									</tr>
									<tr> 
									  <td height="15" colspan="2"></td>				  
									</tr>	
									<tr> 
									  <td nowrap align=left colspan="2">●ファイルアップロード</td>				  
									</tr>						
									<tr>
									  <td align="center" colspan="2"><input type="file" name="txtFileUpload2" size="60"></td>
									</tr>
									<tr>
									  <td align="center" colspan="2"><input type="button" name="btnUpload2" value=" アップロード " onClick="fUpload2();"></td>
									</tr>		
								  </table>
							</form>						  
								  <center>					  
								  <br>
								  <br>
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
	  <table border=0><tr><td height=20></td></tr></table>	
	</center>
    </td>
 </tr>
 	<%
		DisplayFooter
	%> 
</table>
</body>
</HTML>
