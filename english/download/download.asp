<!--#include File="../../adminmenu/Common/common.inc"-->
<!--#include File="download.inc"-->

<%	dim file1,gerrmsg
	
	Dim param(2)
	dim v_Filename
	
	v_Filename = request.querystring("guide")
	
	if v_Filename <> "" then	
		call getDownloadIni(param,v_Guide)
	else
		v_Filename = request.querystring("form")	
		call getDownloadIni(param,v_Form)
	end if	
	
	response.buffer=true
		
	if param(0) <> "" then
		if not gfdownloadFile(param(0) & v_Filename, v_Filename) then
			response.clear
			response.write gerrmsg
		end if
	else
		session("ErrMsg")="パラメータエラー!!"
		response.write session("ErrMsg")
	end if	
%>