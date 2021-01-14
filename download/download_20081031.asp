<!--#include File="../download/download.inc"-->
<%	dim file1,gerrmsg
	
	Dim param(2)
	getDownloadIni param
	
	response.buffer=true
	
	'response.write param(2)
	'response.end 
	
	if param(0) <> "" then
		if not gfdownloadFile(param(0), Mid(param(0), InstrRev(param(0), "\")+1)) then
			response.clear
			response.write gerrmsg
		end if
	else
		session("ErrMsg")="パラメータエラー!!"
		response.write session("ErrMsg")
	end if	
%>