<%
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
  
  
'-----------------------------
function gfHTMLEncode(str)
	if gftrim(str)<>"" then
		gfHTMLEncode=server.htmlencode(gftrim(str))
	else
		gfHTMLEncode=""
	end if
end function
'-----------------------------
function gfSQLEncode(str)
	if gftrim(str)<>"" then
		gfSQLEncode=replace(gftrim(str),"'","''")
	else
		gfSQLEncode=""
	end if
end function

'-----------------------------
function gfXlsPDFChg(xlspath, pdfpath)
Dim objExcelApp, objWbk1
	On Error Resume Next
	
	gfXlsPDFChg = ""
	
	' Excelのオブジェクトの参照を取得
	Set objExcelApp = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
		gfXlsPDFChg = "PDF作成エラー0：" & Err.Description
		exit function
	end if

	' Excelウィンドウを非表示
	objExcelApp.Visible = false
	'Excelオープン
	Set objWbk1 = objExcelApp.Workbooks.Open(xlspath, False, True)

	If Err.Number <> 0 Then
		gfXlsPDFChg = "PDF作成エラー1：" & Err.Description
	else
		'PDF保存
		Call objWbk1.ExportAsFixedFormat(0, pdfpath)
		If Err.Number <> 0 Then
			gfXlsPDFChg = "PDF作成エラー2：" & Err.Description
		End If
	end if
	
	' 指定ブックを閉じる
	objWbk1.Saved = True
	objWbk1.Close False
	Set objWbk1 = Nothing

	' Excel終了
	'objExcelApp.Quit
	Set objExcelApp = Nothing
end function

'----------------------------------
'2017/03/02 T.Okui Add-S
'オペ社コードを会社名に変換
function gfConvertOperator(Operator)
	dim OpeName
	OpeName = ""
	Select Case Operator
	  Case "HKK"
	  	OpeName = "博多港運（株）"
	  Case "KAM"
	  	OpeName = "（株）上組"
	  Case "KTC"
	  	OpeName = "（株）ジェネック"
	  Case "MLC"
	  	OpeName = "三菱倉庫（株）"
	  Case "NEC"
	  	OpeName = "日本通運（株）"
	  Case "SOG"
	  	OpeName = "相互運輸（株）"
	  Case Else
	  	OpeName = Operator
	End Select
	gfConvertOperator = OpeName
end function
'2017/03/02 T.Okui Add-E
%>