<%
'-----------------------------
'   ���l�ϊ� (Long�^)
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
'   Trim�@NULL�̏ꍇ����l(Space0)
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
	
	' Excel�̃I�u�W�F�N�g�̎Q�Ƃ��擾
	Set objExcelApp = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
		gfXlsPDFChg = "PDF�쐬�G���[0�F" & Err.Description
		exit function
	end if

	' Excel�E�B���h�E���\��
	objExcelApp.Visible = false
	'Excel�I�[�v��
	Set objWbk1 = objExcelApp.Workbooks.Open(xlspath, False, True)

	If Err.Number <> 0 Then
		gfXlsPDFChg = "PDF�쐬�G���[1�F" & Err.Description
	else
		'PDF�ۑ�
		Call objWbk1.ExportAsFixedFormat(0, pdfpath)
		If Err.Number <> 0 Then
			gfXlsPDFChg = "PDF�쐬�G���[2�F" & Err.Description
		End If
	end if
	
	' �w��u�b�N�����
	objWbk1.Saved = True
	objWbk1.Close False
	Set objWbk1 = Nothing

	' Excel�I��
	'objExcelApp.Quit
	Set objExcelApp = Nothing
end function

'----------------------------------
'2017/03/02 T.Okui Add-S
'�I�y�ЃR�[�h����Ж��ɕϊ�
function gfConvertOperator(Operator)
	dim OpeName
	OpeName = ""
	Select Case Operator
	  Case "HKK"
	  	OpeName = "�����`�^�i���j"
	  Case "KAM"
	  	OpeName = "�i���j��g"
	  Case "KTC"
	  	OpeName = "�i���j�W�F�l�b�N"
	  Case "MLC"
	  	OpeName = "�O�H�q�Ɂi���j"
	  Case "NEC"
	  	OpeName = "���{�ʉ^�i���j"
	  Case "SOG"
	  	OpeName = "���݉^�A�i���j"
	  Case Else
	  	OpeName = Operator
	End Select
	gfConvertOperator = OpeName
end function
'2017/03/02 T.Okui Add-E
%>