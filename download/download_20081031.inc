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
			gErrMsg="�t�@�C���̓ǂݍ��݂Ɏ��s���܂����B<br>" & fname2
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
	
	'''INI�t�@�C������p�����[�^��Ǎ���
	'''Input :Array(1)
	'''OutPut:Array(1)=USER_GUIDE_FILE
	function getDownloadIni(param)
	dim ObjFSO,ObjTS,tmpStr
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("/download/ini/download.ini"),1,false)
	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)
		Select Case tmpStr(0)			
			Case "USER_GUIDE_FILE"
				param(0) = tmpStr(1)			
		End Select
	Loop
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing

	end function
%>