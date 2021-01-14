<%
'''SQLサーバの接続
'''Input :DB接続Object、レコードセットObject
Function ConnDBH(ObjConn,ObjRS)
  Set ObjConn = Server.CreateObject("ADODB.Connection")
  ObjConn.Open "Provider=SQLOLEDB; Data Source = SI-KURT-LAPTOP\SQLEXPRESS; Initial Catalog = OPE3_KAM; User Id = sa; Password=P@55w0rd"
  ' ObjConn.Open "wavedb","wave","wave"

  ObjConn.CursorLocation = 3	'''3:adUseClient

  Set ObjRS = Server.CreateObject("ADODB.Recordset")
  ObjConn.BeginTrans
End Function


'''SQLサーバの切断
'''Input :DB接続Object、レコードセットObject
Function DisConnDBH(ObjConn,ObjRS)
  Set ObjRS = Nothing
  ObjConn.CommitTrans
  ObjConn.Close
  Set ObjConn = Nothing
End Function


'''サーバ日時の取得
'''Input :Array(5)=(YYYY,MM,DD,hh,mm)
Function getDayTime(DayTime)
	dim MyDate
	MyDate = Now
	DayTime = Array(Year(MyDate),Right("0"&Month(MyDate),2),Right(Date(),2),Right("0"&Hour(MyDate),2),Right("0"&Minute(MyDate),2))
End Function


'''INIファイルからパラメータを読込む
'''Input :Array(1)
'''OutPut:Array(0)=PDF_FOLDER_PATH、Array(1)=PERIOD
Function getIni(param)
	dim ObjFSO,ObjTS,tmpStr
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- ファイルを開く（読み取り専用） ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("/inform/ini/inform.ini"),1,false)
	'--- ファイルデータの読込み ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)
		Select Case tmpStr(0)
			Case "PDF_FOLDER_PATH"
				param(0) = tmpStr(1)
			Case "PERIOD"
				param(1) = tmpStr(1)
		End Select
	Loop
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
End Function
%>
