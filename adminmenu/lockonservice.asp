<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: 
'  【プログラム名称】　: 
'
'  （変更履歴）
'
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
%>
<!--#include File="./Common/common.inc"-->

<%
		'ユーザデータ所得
	dim USER, COMPcd  			
	dim v_GamenMode
	dim v_DataCnt2
	dim v_Msg
	
		
	dim Num2	
	dim strOrder2
	dim FieldName2	
	dim ObjRS2,ObjConn2
	dim ObjConnLO, ObjRSLO
	    
	dim wk
	dim i,x,j
	dim v_ItemName
	dim abspage, pagecnt,reccnt	
	
	dim Arr_DriverID()
	dim Arr_Check()
	dim Arr_HiTSUserID()
	dim Arr_NumOfRedCard()
	dim Arr_AcceptStatus()
	
	dim Arr_Users()
	dim v_CompanyName
	
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	'Search Condition Start
	dim SSearchType
	dim SUserID
	dim SDriverName
	dim SDriverCompany
	dim SDriverID
	'Search Condition End
		
	const gcPage = 10
	
	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' 再描画前の項目取得
   	'----------------------------------------			
	call LfGetRequestItem
		
		
	If v_GamenMode = "S" Then
      Call getCompanyName()
	End If
		
	'更新
	If v_GamenMode = "U" Then		
		call LfUpdDriverLimit()
	End If
	
	Call getAllUser()
	Call getDriverInfo()
	
	
Function LfGetRequestItem()
   
	If Request.form("Gamen_Mode") = "" Then
	  v_GamenMode = Request.QueryString("GamenMode")
	Else
	  v_GamenMode = Request.form("Gamen_Mode")
	End If
	
	if Trim(v_GamenMode) = "PS" then
	  SDriverName = Request.QueryString("SDriverName")
	  SDriverCompany = Request.QueryString("SDriverCompany")
	  SDriverID = Request.QueryString("SDriverID")
	  SSearchType = Request.QueryString("searchType")
      v_DataCnt2 = Request.QueryString("Data_Cnt")
      SUserID = Request.QueryString("SUserID")
      v_CompanyName = Request.QueryString("FullName")
	else
	  SDriverName = Request.form("SDriverName")
	  SDriverCompany = Request.form("SDriverCompany")
	  SDriverID = Request.form("SDriverID")
	  SUserID = Request.Form("cmbUser")
	  v_DriverInfo = Request.Form("driverInfo")
	  SSearchType = Request.form("searchType")
      v_DataCnt2 = Request.form("Data_Cnt")
      v_CompanyName = Request.form("FullName")
    end if
    If v_DataCnt2 = "" then
      v_DataCnt2 = 0
    end if
    
    If Trim(SSearchType) = "" Then
      SSearchType = "1"
    End If
    
	ReDimension(v_DataCnt2)

	For i = 1 to (v_DataCnt2) - 1 
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
        Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
        Arr_HiTSUserID(i) = TRIM(Request.form("HiTSUserID" & i))
        Arr_NumOfRedCard(i) = TRIM(Request.form("NumOfRedCard" & i))
        Arr_AcceptStatus(i) = TRIM(Request.Form("AcceptStatusFlag" & i))
	Next
	
End Function

Function ReDimension(index)
   Redim Arr_Check(index)
   Redim Arr_DriverID(index)
   Redim Arr_HitsUserID(index)
   Redim Arr_NumOfRedCard(index)
   Redim Arr_AcceptStatus(index)
End Function

Function getAllUser()   
   dim StrSQL
   dim recCnt
   recCnt = 0
   Redim Arr_Users(recCnt)
   
   StrSQL = "SELECT * FROM mUsers " 
   StrSQL = StrSQL & " WHERE UserCode <> '' "
   StrSQL = StrSQL & " ORDER BY UserCode"
   ConnectSvr ObjConnLO, ObjRSLO
   
   ObjRSLO.Open StrSQL, ObjConnLo, 0, 1, 1
   
   While Not ObjRSLO.EOF
     Redim Preserve Arr_Users(recCnt)
     Arr_Users(recCnt) = ObjRSLO("UserCode")
     recCnt = recCnt + 1
     ObjRSLO.MoveNext
   Wend

   ObjRSLO.Close
   ObjConnLo.Close
End Function

Function getCompanyName
  dim StrSQL
  
  ConnectSvr ObjConnLO, ObjRSLO
  StrSQL = " SELECT * FROM mUsers "
  StrSQL = StrSQL & " WHERE UserCode='" & Trim(SUserID) & "'"
  ObjRSLO.Open StrSQL, ObjConnLo, 0, 1, 1
   
  While Not ObjRSLO.EOF
    v_CompanyName = Trim(ObjRSLO("FullName"))
    ObjRSLO.MoveNext
  Wend
  
  ObjRSLO.Close
  ObjConnLo.Close
  
End Function

Function getDriverInfo()
    dim StrSQL
 
    ConnectSvr ObjConn2, ObjRS2
    
    StrSQL = "SELECT DISTINCT LomDriver.*, mUsers.UserCode, mUsers.FullName, mUsers.TelNo "
    StrSQL = StrSQL & " FROM LomDriver "
    StrSQL = StrSQL & " LEFT JOIN mUsers ON LomDriver.HiTSUserID = mUsers.UserCode AND (LomDriver.AcceptStatus='1' OR LomDriver.AcceptStatus='3') "
    if Trim(SDriverName) <> "" or Trim(SDriverCompany) <> "" or Trim(SDriverID) <> "" or Trim(SUserID) <> "" then
       StrSQL = StrSQL  & " WHERE "
       
       If Trim(SUserID) <> "" Then
         StrSQL = StrSQL  & "LomDriver.HiTSUserID = '" & Trim(SUserID) & "' "
       End If
       
       if Trim(SDriverName) <> "" then
         If Trim(SUserID) <> "" Then
           StrSQL = StrSQL  & " AND "
         End If
         StrSQL = StrSQL  & "LomDriver.LoDriverName LIKE '%" & Trim(SDriverName) & "%' "
       end if
       
       if Trim(SDriverCompany) <> "" then
         if Trim(SDriverName) <> "" Or Trim(SUserID) <> "" then
            StrSQL = StrSQL  & " AND "  
         end if
         StrSQL = StrSQL  & "LomDriver.LoDriverCompany LIKE '%" & Trim(SDriverCompany) & "%' "
       end if
       if Trim(SDriverID) <> "" then
         if Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" Or Trim(SUserID) <> "" then
            StrSQL = StrSQL  & " AND "  
         end if
         StrSQL = StrSQL  & "LomDriver.LoDriverID LIKE '%" & Trim(SDriverID) & "%' "
       end if
    end if
    
    if Trim(SDriverName) = "" and Trim(SDriverCompany) = "" and Trim(SDriverID) = "" And Trim(SUserID) = "" then
      StrSQL = StrSQL  & " WHERE "
    end if
   
    if Trim(SDriverID) <> "" Or Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" Or Trim(SUserID) <> "" then
        StrSQL = StrSQL  & " AND "  
    end if
    
    If Trim(SSearchType) = "1" Then
      StrSQL = StrSQL & " LomDriver.AcceptStatus='3'"
    ElseIf Trim(SSearchType) = "2" Then
      StrSQL = StrSQL & " (LomDriver.AcceptStatus='3' OR LomDriver.AcceptStatus='1') " 
    End If 
	
	'Response.Write StrSQL
	
    ObjRS2.PageSize = 100
	ObjRS2.CacheSize = 100
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2, 0, 1, 1

	Num2 = ObjRS2.recordcount	

	if Num2 > 100 then
		If CInt(Request("pagenum2")) = 0 Then
			ObjRS2.AbsolutePage = 1
		Else
			If CInt(Request("pagenum2")) <= ObjRS2.PageCount Then
				ObjRS2.AbsolutePage = CInt(Request("pagenum2"))
			Else
				ObjRS2.AbsolutePage = 1
			End If
		End If		 
	end if
	
	if err <> 0 then
	    ObjRS2.close
	    ObjConn2.close
		Exit Function
	end if			
	'エラートラップ解除
    on error goto 0	

End Function

Function LfUpdDriverLimit()
    dim StrSQL
    dim ErrFlg
    dim iSeq
	
    ConnectSvr ObjConnLO, ObjRSLO	
	
	For i = 1 to v_DataCnt2-1
      If UCase(Trim(Arr_Check(i))) = "ON" Then
         If Trim(Arr_AcceptStatus(i)) <> "3" Then
           StrSQL = " UPDATE LomDriver SET "
           StrSQL = StrSQL & " AcceptStatus='3', "
           StrSQL = StrSQL & " NumOfRedCard='" & CInt(Trim(Arr_NumOfRedCard(i))) + 1 & "', "
           StrSQL = StrSQL & " UpdtTime='" & Now() & "' "
           StrSQL = StrSQL & " WHERE "
           StrSQL = StrSQL & " LoDriverID='" & Trim(Arr_DriverID(i)) & "'"           
           ObjConnLO.Execute StrSQL
           If err.number <> 0 then
             v_Msg = "変更できません。"
             ObjConnLO.rollbacktrans
             ObjRSLO.Close
             ObjConnLO.Close
             Exit Function
		   End If
		 End If
	  Else
	     StrSQL = " UPDATE LomDriver SET "
         StrSQL = StrSQL & " AcceptStatus='1', "
         StrSQL = StrSQL & " UpdtTime='" & Now() & "' "
         StrSQL = StrSQL & " WHERE "
         StrSQL = StrSQL & " LoDriverID='" & Trim(Arr_DriverID(i)) & "'"           
         ObjConnLO.Execute StrSQL
         if err.number <> 0 then
           v_Msg = "変更できません。"
           ObjConnLO.rollbacktrans
           ObjRSLO.Close
           ObjConnLO.Close
           Exit Function
		 end if
      End If
    Next
    ObjConnLO.Close
    
End Function

function LfPutPage(rec,page,pagecount,link)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo
	dim intNextFlag
	dim strParam
	PageIndex=0
	PageWkNo=0	
	if rec > 0 then	

		if pagecount<page then
			page=pagecount
		end if
		'ページIndexを設定
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		PageWkNo=((gcPage*PageIndex)+1)-gcPage
		
		
		'先頭ページが0より小さい場合は1を設定
		if PageWkNo<=0 Then
			PageWkNo=0
		End If
        

		'パラメータ設定
		
	    'strParam="&InOutF=" & v_InOutFlag
		strParam=""
		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """>最初へ</a>"
			response.write "| &nbsp;"
			if PageWkNo<>0 Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """>前へ</a>"
			Else
				response.write "<font style='color:#FFFFFF;'>前へ</font>"
			End If
		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        		
		'--- インデックス
		'ページが1ページ以上存在する場合
		if pagecount>1 then
			response.write "| &nbsp;"

			'指定ページ数分ループ
			for i=1 to gcPage
				'ページ数算出
				PageWkNo=(gcPage*PageIndex)+i

				'ページが全ページより大きい場合は処理中断
				if pagecount< PageWkNo then
					PageWkNo=PageWkNo-1
					exit for
				end if
				'現在選択されているページの場合
				if PageWkNo=page then
					response.write "&nbsp;" & PageWkNo 
				else
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
					
		if page<pagecount-1 then
			PageWkNo=PageWkNo+1
			If PageWkNo<=LastPage Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """>次へ</a>"'
			Else
				response.write "<font style='color:#FFFFFF;'>次へ</font>"
			End If
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & """>最後へ</a>"'            
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function
'-----------------------------
'   数値変換 (Long型)
'-----------------------------
function gfCLng(str1)
    dim str
    str=gfTrim(str1)
    if isnull(str) then
        gfCLng=0
    elseif trim(str)="" then
        gfCLng=0
    elseif not isNumeric(str) then
        gfCLng=0
    elseif len(str)>9 then
        if instr(str,".")>0 and instr(str,".")<10 then
            gfClng=clng(left(str,instr(str,".")-1))
        else
            gfClng=0
        end if
    else
        gfCLng = CLng(fix(str))
    end if
end function

'-----------------------------
'   Trim　NULLの場合→空値(Space0)
'-----------------------------
function gfTrim(str)
    if isnull(str) then
        gfTrim=""
    else
        gfTrim=trim(str)
    end if
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-ロックオンサービス制限</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<STYLE>
th.hlist {
	position: relative;
}
th {
    border-width: 1px 1px 1px 1px;
    padding: 4px;
    background-color: #ffcc33;
}
SELECT.chr {
    BACKGROUND-COLOR: #ffffff;
    BORDER-BOTTOM: #ffffff 1px solid;
    BORDER-LEFT: #002f7b 0px solid;
    BORDER-RIGHT: #ffffff 0px solid;
    BORDER-TOP: #ffffff 0px solid;
    COLOR: black;
    FONT-FAMILY: 'ＭＳ ゴシック';
    FONT-SIZE: 10px;
    FONT-WEIGHT: normal;
    PADDING-BOTTOM: 2px;
    PADDING-LEFT: 1px;
    PADDING-RIGHT: 2px;
    PADDING-TOP: 3px;
    TEXT-ALIGN: left
}
table {
    border-width: 0px 1px 1px 0px;
}
DIV.center {
	text-align:center;
}
DIV.BDIV1 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
DIV.BDIV2 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
thead tr {
    position: relative;
    top: expression(this.offsetParent.scrollTop);
}
#loading2 {
	font:bold 10px Verdana;
	color:red;
	position:absolute; 
	top:220px; 
	left:390px;
	width:300px;
	height:30px; 
	z-index:69;
	font-size:12pt;
	border:0px;
	vertical-align: middle;
}
#footer {
 position: fixed;
 top: 100%;
 width: 100%;
}

.cmbUser option{
  height:10px;
}

</STYLE>
<SCRIPT Language="JavaScript">
function finit(){
	//データ引継ぎ設定
	  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";  
    document.frm.cmbUser.value="<%=SUserID%>";	
    document.frm.FullName.value="<%=v_CompanyName%>"
    if("<%=SSearchType %>"=="1"){
      document.getElementById("chk1").checked=true;
    }
    else{
      if("<%=SSearchType %>"=="2"){
         document.getElementById("chk2").checked=true;
      }
    }
    if ("<%=v_Msg%>" != ""){
      alert("<%=v_Msg%>");
      return false;
    }
    
}

//データが無い場合の表示制御
function view(){
	var sortedHeight;
	sortedHeight = 0;
	var vHeight;
	var obj2=document.getElementById("BDIV2");
	var rowHeight;
	
	if('<%=Num2%>'!='0'){
	  var rowHeightThead = getRowHeightThead();
	  var rowHeightTbody = getRowHeightTbody();
	  if(rowHeightThead > 0){
	    rowHeightThead=rowHeightThead
	  }
	  if(rowHeightTbody > 0){
	    rowHeight=rowHeightTbody*15
	  }
	  rowHeight=rowHeight+rowHeightThead
    }
    else{
      rowHeight = 0;
      rowHeight=23*15;
    }
    
	if((document.body.offsetWidth) < 240){
		obj2.style.width=62;
		obj2.style.overflowX="auto";	 
	}else if((document.body.offsetWidth)<1037){
		obj2.style.width=document.body.offsetWidth-140;
		obj2.style.overflowX="auto";
	}//else{
	//	obj2.style.width=document.body.offsetWidth-380;
	//	obj2.style.overflowX="auto";
	//}	

	if((document.body.offsetHeight-rowHeight) < 130){ 
	    if(obj2.clientWidth<obj2.scrollWidth)
	    {
	      obj2.style.height = 62;
		  obj2.style.overflowY = "auto";
	    }
	    else{
	      obj2.style.height = 47;
		  obj2.style.overflowY = "auto";
		}
	}else if((document.body.offsetHeight-rowHeight) < 395){
	    vHeight = rowHeight + 68;
		obj2.style.height = document.body.offsetHeight-vHeight;
		obj2.style.overflowY = "auto";
	}else{
		if(obj2.clientWidth < obj2.scrollWidth)
        {
           obj2.style.height = rowHeight+17;
           obj2.style.overflowY="auto";
        }
        else{
           obj2.style.height = rowHeight;
           obj2.style.overflowY="auto";
        } 
	}
}

function getRowHeightThead()
{
  var oRows = document.getElementById('TBInOut').getElementsByTagName('thead');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}

function getRowHeightTbody()
{
  var oRows = document.getElementById('TBInOut').getElementsByTagName('td');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}
function ClearSItem()
{
  document.frm.SDriverName.value="";
  document.frm.SDriverCompany.value="";
  document.frm.SDriverID.value="";
}

function ClearSItem2()
{
  document.frm.cmbUser.value="";
  document.frm.SDriverName.value="";
  document.frm.SDriverCompany.value="";
  document.frm.SDriverID.value="";
  document.frm.Gamen_Mode.value = "S";
  document.frm.submit();
}

function fSearch(){
	document.frm.Gamen_Mode.value = "S";
	ClearSItem();
    document.frm.submit();
}

function fSearchDriver(){
   	document.frm.Gamen_Mode.value = "SD";
    document.frm.submit();

}
</script>

<script type="text/vbscript">

Public Sub Update_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 1

  if chkFlag=1 then
    x=MsgBox("        上の内容で登録します。" & vbCrLf & "             よろしいですか？" ,4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "U"
      document.frm.submit()
    end if
  end if

End Sub
</script>


</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();view();" onResize="view();">
<form name="frm" action="lockonservice.asp" method="post">
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0" style="width:1020px;">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("ロックオンサービス制限")
		%>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
		  <INPUT type=hidden name="FullName" readonly tabindex=-1>        
      </table>
 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
      <td style="width:90px">&nbsp;</td>
      <td>
         <div id="BDIV3">
            <!--Detail Start-->
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td>
                  <DIV style="width:230px; padding:10px;background-color:#FFCCFF; text-align:center;">ドライバ一覧からの制限 on/off</DIV>
                </td>
              </tr>
              <tr><td>&nbsp;</td></tr>
              <tr nowrap>
                <td nowrap>
                  <div style="margin-left:13px">
                  <table>
                    <tr>
                      <td nowrap><input type="radio" name="searchType" id="chk1" value="1" checked=true onclick="ClearSItem2();">制限中ドライバのみ表示</td>
                      <td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
                      <td nowrap><input type="radio" name="searchType" id="chk2" value="2" onclick="ClearSItem2();">全ドライバを表示</td>
                    </tr>
                  </table>
                  </div>
                </td>
              </tr>
              <tr><td>&nbsp;</td></tr>
              <tr>
                <td>
                  <div style="margin-left:20px">
                  <table>
                    <td nowrap style="width:150px">承認ユーザ指定選択</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td> 
	                  <td nowrap>
	                    <select name="cmbUser" style="width:116px;" onchange="fSearch();" class="cmbUser">
	                      <OPTION VALUE = '' SELECTED></OPTION>
	                      <% If UBound(Arr_Users) > 0 Then%>
	                        <% For j=0 to UBound(Arr_Users) %>
	                          <OPTION VALUE = '<%=Arr_Users(j)%>'><%=Arr_Users(j) %></OPTION>
	                        <% Next %>
	                      <% End If %>
			            </select>
	                  </td>
	                  <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	                  <td nowrap>
	                    <%=v_CompanyName%>
	                  </td>
                    </td>
                  </table>
                  </div>
                </td>
              </tr>
              
              <tr nowrap>
                <td nowrap>
                  <div style="margin-left:20px">
                  <!--Search Conditions Start-->
                    <table>
                      <tr>
                        <td nowrap style="width:150px">ドライバ名指定</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap><input type="text" name="SDriverName" value="<%=SDriverName%>" onfocus="this.select();"></td>
                      </tr>
                      <tr>
                        <td nowrap style="width:150px">ドライバＩＤ指定</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap><input type="text" name="SDriverID" value="<%=SDriverID%>" onfocus="this.select();"></td>
                      </tr>
                      <tr>
                        <td nowrap style="width:150px">会社名指定</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap><input type="text" name="SDriverCompany" value="<%=SDriverCompany%>" onfocus="this.select();"></td>
                        <td width=100 align=right nowrap><input type="button" name="Button" value="表示更新" onClick="fSearchDriver();"></td>
                        <td width=150 align=right nowrap>※部分一致検索可</td>
                      </tr>
                    </table>
                  <!--Search Conditions End-->
                  </div>    
                </td>
              </tr>
              <tr nowrap>
                <td nowrap><BR/></td>
              </tr>
              <tr align=right nowrap>
                <td width="100%" height="30" align=right nowrap>
                  <div style="margin-left:20px">
                  <table border="0" cellpadding="0" cellspacing="0">
                    <tr>
		               <td width="100%" align="center" nowrap>
		                 <!--Page Pagination Start-->
		                 <%					
				           If Num2 > 0 Then						
					         abspage = ObjRS2.AbsolutePage
					         pagecnt = ObjRS2.PageCount
					         call LfPutPage(Num2,abspage,pagecnt,"pagenum2")
				           End If									
			             %>
		                 <!--Page Pagination End-->
		               </td>
		            </tr>
		          </table> 
		          </div>
                </td>
              </tr>
	          <tr>		
		        <td nowrap>
		          <div style="margin-left:20px">
			      <div id="BDIV2">
			   	    <% If Num2>0 Then%>
			   		<!--Driver List Start-->	
					<table border="1" cellpadding="0" cellspacing="0" width=100% id="TBInOut">				
					  <thead>
					    <!--HEADER INFORMATION START-->
					    <tr>
					      <th id="H2Col01" class="hlist" align="center" nowrap>制限中</th>
					      <th id="H2Col02" class="hlist" nowrap>ドライバ名</th>
					      <th id="H2Col03" class="hlist" nowrap>ドライバID</th>								
					      <th id="H2Col04" class="hlist" nowrap>会社名</th>
					      <th id="H2Col05" class="hlist" nowrap>承認ユーザ</th>
						  <th id="H2Col06" class="hlist" nowrap>承認ユーザ会社名</th>
						  <th id="H2Col07" class="hlist" nowrap>承認ユーザ<BR/>電話番号</th>
						  <th id="H2Col08" class="hlist" nowrap>制限<BR />実施日時</th>		
						  <th id="H2Col09" class="hlist" nowrap>制限<BR />回数</th>																									
				        </tr>
					    <!--HEADER INFORMATION END-->
					  </thead>																
					  <tbody>
					    <!--DETAIL INFORMATION START-->
                        <% 
						  x = 1 							
						  For i=1 To ObjRS2.PageSize
						    If Not ObjRS2.EOF Then
							  x = x + 1
						%>
						<tr bgcolor="#CCFFFF">	
						  <% v_ItemName = "chkInOut" + cstr(i) %>
						  <td id="D2Col01" align="center" nowrap style="width:60px;">
						    <% If Trim(ObjRS2("AcceptStatus")) = "3" Then %>
						      <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR />
						    <% Else %>
						      <input type="checkbox" name="<%= v_ItemName %>" ><BR />
						    <% End If %>
						  </td>
						  <td id="D2Col02" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("LoDriverName"))%><BR />
                          </td>
						  <td id="D2Col03" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("LoDriverID"))%><BR />
                          </td>
						  <td id="D2Col04" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("LoDriverCompany"))%><BR />
                          </td>

                          <td id="D2Col05" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("UserCode"))%><BR />
                          </td>
                          <td id="D2Col06" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("FullName"))%><BR />
                          </td>
                          <td id="D2Col07" align="center" valign="middle" nowrap style="width:100px;">
                            <%=Trim(ObjRS2("TelNo"))%><BR />
                          </td>
                          <td id="D2Col08" align="center" valign="middle" nowrap style="width:130px;">
                            <% If Trim(ObjRS2("AcceptStatus")) = "3" Then %>
                               <%=Trim(Year(ObjRS2("UpdtTime"))) & "-" & Right("0" & Trim(Month(ObjRS2("UpdtTime"))),2) & "-" & Right("0" & Trim(Day(ObjRS2("UpdtTime"))),2) & " " & Trim(FormatDateTime(ObjRS2("UpdtTime"),4))%>                            
                            <% End If%><BR />
                          </td>
                          <td id="D2Col09" align="center" valign="middle" nowrap style="width:40px;">
                            <%=Trim(ObjRS2("NumOfRedCard"))%>
                            <BR />
                          </td>

                          <% v_ItemName = "LODriverID" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("LODriverID"))%>">	
						  <% v_ItemName = "HiTSUserID" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("HiTSUserID"))%>">
						  <% v_ItemName = "NumOfRedCard" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("NumOfRedCard"))%>">
						  <% v_ItemName = "AcceptStatusFlag" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("AcceptStatus"))%>">
						  
						</tr>
						<% 
						      ObjRS2.MoveNext 		
						    End If
						  Next	
						  ObjRS2.close    
						  ObjConn2.close
						%>  
						<!--DETAIL INFORMATION END-->	
						    									
					  </tbody>								
					</table>
					<!--Driver List End-->
					<INPUT type="hidden" name="Data_Cnt" value="<%=x%>">
					
				    <% Else %>
					  <table border="0" cellPadding="2" cellSpacing="0" id="NODATA">						
					    <TR class=bgw><TD nowrap style="color:Red;">ドライバーの登録がありません</TD></TR>
					  </table>
				    <% End If %>		
			      </div>
			      </div>
		        </td>
	          </tr>
	          <tr><td>&nbsp;</td></tr>  
	          <tr>
	            <td>
	              <div style="margin-left:20px">
	              <p style="color:#993300;">※「制限中」のチェックを入れることで、ドライバはロックオンができなくなります。<BR />
                        &nbsp;&nbsp;&nbsp;&nbsp;また、そのドライバへの指示も出せなくなります。<BR />
                        &nbsp;&nbsp;&nbsp;&nbsp;解除する場合はチェックを外します。
                  </p>
                  </div>
	            </td>
	          </tr>
	          <tr><td>&nbsp;</td></tr>
	          <tr>		
		        <td>
		          <div>
		            <div style="margin-left:20px">
			        <table border="0" cellpadding="2" cellspacing="0">
			          <tr>
			            <td><input type="button" name="Update"  value="上の内容で登録"></td>
			          </tr>
			        </table>
			        </div>	
			      </div>		
		        </td>
	          </tr> 
	          <tr><td>&nbsp;</td></tr>
	          <tr><td><center><a href="menu.asp">閉じる</a></center></td></tr>   
	          <tr><td>&nbsp;</td></tr>
            </table>
          </div>
      </td>
      <td style="width:30px">&nbsp;</td>   
      </tr></table>  
    </td>
 </tr>
	<%
		DisplayFooter
	%>
</table>
 
</form>
</body>
</HTML>
