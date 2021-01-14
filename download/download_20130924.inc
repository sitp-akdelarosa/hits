<%
	function gfdownloadFile(fname1, fname2)
		on error resume next
		err.clear
		dim bo, barry, file1		
		'file1	= server.mappath(fname1)
		file1	= fname1		
		set bo = Server.CreateObject("basp21")
		barry = bo.BinaryRead(file1)
		if not isArray(barry) then
			response.write file1
			gErrMsg="ファイルの読み込みに失敗しました。<br>" & fname2
			gfdownloadFile=false
			set bo = nothing
			exit function
		end if
		Response.Addheader "Content-Disposition", "attachment ; filename=" & fname2
		Response.Contenttype = "application/x-binary"
		Response.Binarywrite barry
		set bo = nothing
		if err.number=0 then
			gfdownloadFile=true
		else
			gErrMsg=gErrMsg & vbcrlf & err.description
			gfdownloadFile=false
		end if
	end function
	
	'''INIファイルからパラメータを読込む
	'''Input :Array(1)
	'''OutPut:Array(1)=USER_GUIDE_FILE
	function getDownloadIni(param,strDir)
	dim ObjFSO,ObjTS,tmpStr

	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- ファイルを開く（読み取り専用） ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("../adminmenu/ini/admin.ini"),1,false)

	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 11, 1)
		Select Case tmpStr(0)			
			Case strDir
				param(0) = tmpStr(1)			
		End Select
	Loop
	
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
	end function
%>
