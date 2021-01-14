<%		
	function gfUploadFile(fname1,upload,strDir)		
		dim ret
		dim wkfilename 		
		dim param(2)
		
		gfUploadFile = True
		
		call getUploadIni(param,strDir)
		
		wkfilename	= param(0) & Mid(fname1, InstrRev(fname1, "\")+1)		
		
		ret	= obj.FormSaveAs(buf,upload,wkfilename)
		
		if ret > 0 then
			gfUploadFile = True
		else		
			gfUploadFile = False
		end if
	end function
	
	function gfUploadFile2(fname1, upload)
		dim ret
		dim wkfilename 		
		dim serverpath
		
		gfUploadFile2 = True
		
		serverpath = Server.MapPath("../")
		
		wkfilename	= serverpath & "\" & Mid(fname1, InstrRev(fname1, "\")+1)
		
		ret	= obj.FormSaveAs(buf,upload,wkfilename)
		
		if ret > 0 then
			gfUploadFile2 = true
		else		
			gfUploadFile2 = false
		end if
	end function

	function gfUploadFile3(fname1, upload,strDir)
		dim ret
		dim wkfilename 		
		
		gfUploadFile3 = True
		
		wkfilename	= strDir & Mid(fname1, InstrRev(fname1, "\")+1)
		
		ret	= obj.FormSaveAs(buf,upload,wkfilename)
		
		if ret > 0 then
			gfUploadFile3 = true
		else		
			gfUploadFile3 = false
		end if
	end function
	
	function gfDeleteFile(fname1)		
		dim ObjFSO,ObjTS,tmpStr
		dim cnt
		
		Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		objFSO.DeleteFile(fname1)		
		Set ObjFSO = Nothing	
	end function
		
	'''INIファイルからパラメータを読込む
	'''Input :Array(1), Variable Name
	function getUploadIni(param,strVariable)
		dim ObjFSO,ObjTS,tmpStr
		Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		'--- ファイルを開く（読み取り専用） ---
		Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/admin.ini"),1,false)
		'--- ファイルデータの読込み ---
		Do Until ObjTS.AtEndofStream
			tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)			
			Select Case tmpStr(0)							
				Case strVariable							
					param(0) = tmpStr(1)
			End Select
		Loop
		ObjTS.Close
		Set ObjTS = Nothing
		Set ObjFSO = Nothing
	end function	
	
		'''INIファイルからパラメータを読込む
	'''Input :Array(1)
	'''OutPut:Array(1)=PDF_FOLDER_PATH
	function getInformIni(param,strDir)
	dim ObjFSO,ObjTS,tmpStr

	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- ファイルを開く（読み取り専用） ---
	Set ObjTS = ObjFSO.OpenTextFile(strDir,1,false)

	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)
		Select Case tmpStr(0)			
			Case "PDF_FOLDER_PATH"
				param(0) = tmpStr(1)			
		End Select
	Loop
	
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
	end function

%>
