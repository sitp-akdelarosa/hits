<%@Language="VBScript"%>

<!--#include file="./Common/Common.inc"-->
<%
'  （変更履歴）
'   2013-09-26   Y.TAKAKUWA   スマートフォンのカウントを追加。
%>
<%
	'変数宣言
	Dim sDateF,sDateT,sMode,tmpDate
	Dim iCount,iLoop,iFileCnt,iDateCnt,i,j,k,iHdRow,iGSum(),iTSum(),iMTSum(),iRSum
	Dim HDate()
	Dim iTKind,iMTKind
	Dim PageNum(),WkNum(),PageTitle(),SubTitle(),Count()
	Dim MPageNum(),MWkNum(),MPageTitle(),MSubTitle(),MCount()
	'2013-09-30 Y.TAKAKUWA Add-S
	Dim SPageNum(),SWkNum(),SPageTitle(),SSubTitle(),SCount(),iSTKind
	'2013-09-30 Y.TAKAKUWA Add-E
	Dim strTitleFileName,sHdValue,strFileName
	Dim FPageNum(),FWkNum(),FDate(),FCount()

	' Tempファイル属性のチェック

	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' ダウンロードファイルの取得

	strFileName = Session.Contents("tempfile")
	If strFileName="" Then
		' セッションが切れているとき
		Response.Redirect("accesstotal.asp")	 '利用件数Topへ
		Response.End
	End If
	strFileName="../temp/" & strFileName
	' ダウンロードファイルのOpen
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' ファイルのダウンロード
	Response.ContentType="application/octet-stream"
	Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

	'ヘッダの取得

	anyTmp=Split(ti.ReadLine,",")
	sDateF=anyTmp(0)
	sDateT=anyTmp(1)
	sMode=anyTmp(2)
	
	iLoop=0
	iFileCnt=0
	'CSVファイルデータを変数に格納
	Do While Not ti.AtEndOfStream
		if iLoop<>0 then
			anyTmp=Split(ti.ReadLine,",")
			ReDim Preserve FPageNum(iFileCnt)
			ReDim Preserve FWkNum(iFileCnt)
			ReDim Preserve FDate(iFileCnt)
			ReDim Preserve FCount(iFileCnt)
			FPageNum(iFileCnt)=anyTmp(0)
			FWkNum(iFileCnt)=anyTmp(1)
			if sMode="D" then
				FDate(iFileCnt)=left(anyTmp(2),4) & "/" & mid(anyTmp(2),5,2) & "/" & right(anyTmp(2),2)
			else
				FDate(iFileCnt)=left(anyTmp(2),4) & "/" & mid(anyTmp(2),5,2) 
			end if
			FCount(iFileCnt)=anyTmp(3)
			iFileCnt=iFileCnt+1
		end if
		iLoop=iLoop+1
	Loop

	'ヘッダ書き込み
	if sMode="D" then
			Response.Write left(sDateF,4) & "年" & mid(sDateF,5,2) & "月" & right(sDateF,2) & "日から" & left(sDateT,4) & "年" & mid(sDateT,5,2) & "月" & right(sDateT,2) & "日まで,"
	Else
			Response.Write left(sDateF,4) & "年" & mid(sDateF,5,2) & "月から" & left(sDateT,4) & "年" & mid(sDateT,5,2) & "月まで,"
	End If
	Response.Write Chr(13) & Chr(10)
	Response.Write "＜パソコン＞,"
	Response.Write Chr(13) & Chr(10)

	'パソコンタイトル行出力
	Response.Write "メニュー項目,"
	Response.Write "画面,"
	Response.Write "画面No," 
	
	iDateCnt=0
	iLoop=0
	'ヘッダ日付設定
	'日別の場合
	if sMode="D" then
		tmpDate=left(sDateF,4) & "/" & mid(sDateF,5,2) & "/" & right(sDateF,2)
		iCount=DateDiff("d", left(sDateF,4) & "/" & mid(sDateF,5,2) & "/" & right(sDateF,2), left(sDateT,4) & "/" & mid(sDateT,5,2) & "/" & right(sDateT,2))
		
		do 
			if iCount<iLoop then
				exit do
			end if
			ReDim Preserve HDate(iDateCnt)
			HDate(iDateCnt)=tmpDate
			Response.Write tmpDate & ","
			tmpDate=DateAdd("d", 1, tmpDate)
			iDateCnt=iDateCnt+1
			iLoop=iLoop+1
		Loop  
	else
		iCount=DateDiff("m", left(sDateF,4) & "/" & mid(sDateF,5,2) & "/01", left(sDateT,4) & "/" & mid(sDateT,5,2) & "/01")

		tmpDate=left(sDateF,4) & "/" & mid(sDateF,5,2) & "/" & right(sDateF,2)
		do 
			if iCount<iLoop then
				exit do
			end if
			ReDim Preserve HDate(iDateCnt)
			HDate(iDateCnt)=left(tmpDate,7)
			Response.Write left(tmpDate,7) & ","
			tmpDate=DateAdd("m", 1, tmpDate)
			iDateCnt=iDateCnt+1
			iLoop=iLoop+1
		Loop

	end if
	Response.Write "合計"

	Response.Write Chr(13) & Chr(10)

	'パソコン用 ログタイトル取得
	strTitleFileName="../logweb.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iTKind=0
	
	
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		If anyTmpTitle(2) <> "" Then 
			ReDim Preserve PageNum(iTKind)
			ReDim Preserve WkNum(iTKind)
			ReDim Preserve PageTitle(iTKind)
			ReDim Preserve SubTitle(iTKind)
			PageTitle(iTKind) = anyTmpTitle(2)
			PageNum(iTKind) = ""
			WkNum(iTKind) = ""
			SubTitle(iTKind) = ""
			iTKind=iTKind+1
		end if
		ReDim Preserve PageNum(iTKind)
		ReDim Preserve WkNum(iTKind)
		ReDim Preserve PageTitle(iTKind)
		ReDim Preserve SubTitle(iTKind)
		PageNum(iTKind) = anyTmpTitle(0)
		WkNum(iTKind) = anyTmpTitle(1)
		PageTitle(iTKind) = ""
		SubTitle(iTKind) = anyTmpTitle(3)
		iTKind=iTKind+1
	Loop
	ti.Close

	'携帯用 ログタイトル取得
	strTitleFileName="../logija.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iMTKind=0
	
	
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		If anyTmpTitle(2)<>"" Then 
			ReDim Preserve MPageNum(iMTKind)
			ReDim Preserve MWkNum(iMTKind)
			ReDim Preserve MPageTitle(iMTKind)
			ReDim Preserve MSubTitle(iMTKind)
			MPageTitle(iMTKind) = anyTmpTitle(2)
			MPageNum(iMTKind) = ""
			MWkNum(iMTKind) = ""
			'2017/10/13 Add-S CIS
			'SubTitle(iMTKind) = ""
			MSubTitle(iMTKind) = ""
			'2017/10/13 Add-E CIS
			iMTKind=iMTKind+1
		end if
		ReDim Preserve MPageNum(iMTKind)
		ReDim Preserve MWkNum(iMTKind)
		ReDim Preserve MPageTitle(iMTKind)
		ReDim Preserve MSubTitle(iMTKind)
		MPageNum(iMTKind) = anyTmpTitle(0)
		MWkNum(iMTKind) = anyTmpTitle(1)
		MPageTitle(iMTKind) = ""
		MSubTitle(iMTKind) = anyTmpTitle(3)
		iMTKind=iMTKind+1
	Loop
	ti.Close
	
	'2013-09-30 Y.TAKAKUWA Add-S
	'携帯用 ログタイトル取得
	strTitleFileName="../logsumafo.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iSTKind=0
	
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		If anyTmpTitle(2)<>"" Then 
			ReDim Preserve SPageNum(iSTKind)
			ReDim Preserve SWkNum(iSTKind)
			ReDim Preserve SPageTitle(iSTKind)
			ReDim Preserve SSubTitle(iSTKind)
			SPageTitle(iSTKind) = anyTmpTitle(2)
			SPageNum(iSTKind) = ""
			SWkNum(iSTKind) = ""
			'2017/10/13 Add-S CIS
			'SubTitle(iSTKind) = ""
			SSubTitle(iSTKind) = ""
			'2017/10/13 Add-E CIS
			iSTKind=iSTKind+1
		end if
		ReDim Preserve SPageNum(iSTKind)
		ReDim Preserve SWkNum(iSTKind)
		ReDim Preserve SPageTitle(iSTKind)
		ReDim Preserve SSubTitle(iSTKind)
		SPageNum(iSTKind) = anyTmpTitle(0)
		SWkNum(iSTKind) = anyTmpTitle(1)
		SPageTitle(iSTKind) = ""
		SSubTitle(iSTKind) = anyTmpTitle(3)
		iSTKind=iSTKind+1
	Loop
	ti.Close
	'2013-09-30 Y.TAKAKUWA Add-E

	'パソコン用データ整備
	ReDim Count(iTKind-1,iDateCnt-1)
	ReDim iGSum(iDateCnt-1)
	ReDim iTSum(iDateCnt-1)
	for iLoop=0 to iDateCnt-1
		iTSum(iLoop)=0
	Next
	'パソコン表示項目分ループ
	For i=0 to iTKind-1
		'メニュー項目が変わった場合
		if sHdValue<>PageTitle(i) and trim(PageTitle(i))<>"" then
			'先頭行以外
			if i<>0 then
				for iLoop=0 to iDateCnt-1
					Count(iHdRow,iLoop)=iGSum(iLoop)
				Next
			end if

			iHdRow=i
			for iLoop=0 to iDateCnt-1
				iGSum(iLoop)=0
			Next
			sHdValue=PageTitle(i)
		end if
		'カウントクリア
		For iLoop=0 to iDateCnt-1
			Count(i,iLoop)=0
		Next
		
		'ファイル行数分ループ
		For j=0 to iFileCnt-1
			'画面番号、作業番号が同じ場合
			If PageNum(i)=FPageNum(j) and WkNum(i)=FWkNum(j) then
				'日付分ループ
				For k=0 to iDateCnt-1
					'日付が同じデータの場合
					if cstr(HDate(k))=cstr(FDate(j)) then
						Count(i,k)=Count(i,k)+FCount(j)
						iGSum(k)=iGSum(k)+FCount(j)
						iTSum(k)=iTSum(k)+FCount(j)
						Exit for
					end if
				Next
			end if
		Next
	Next
	'最終行のデータを足しこむ
	For iLoop=0 to iDateCnt-1
		Count(iHdRow,iLoop)=iGSum(iLoop)
	Next	
    '-------------------------------------------------------------------------------------------------------
	'携帯用データ整備
	sHdValue=""
	ReDim MCount(iMTKind-1,iDateCnt-1)
	ReDim iGSum(iDateCnt-1)
	ReDim iMTSum(iDateCnt-1)
	for iLoop=0 to iDateCnt-1
		iMTSum(iLoop)=0
	Next
	'携帯表示項目分ループ
	For i=0 to iMTKind-1
		'メニュー項目が変わった場合
		if sHdValue<>MPageTitle(i) and trim(MPageTitle(i))<>"" then
			'先頭行以外
			if i<>0 then
				for iLoop=0 to iDateCnt-1
					MCount(iHdRow,iLoop)=iGSum(iLoop)
				Next
			end if

			iHdRow=i
			for iLoop=0 to iDateCnt-1
				iGSum(iLoop)=0
			Next
			sHdValue=MPageTitle(i)
		end if
		'カウントクリア
		For iLoop=0 to iDateCnt-1
			MCount(i,iLoop)=0
		Next
		
		'ファイル行数分ループ
		For j=0 to iFileCnt-1
			'画面番号、作業番号が同じ場合
			If MPageNum(i)=FPageNum(j) and MWkNum(i)=FWkNum(j) then
				'日付分ループ
				For k=0 to iDateCnt-1
					'日付が同じデータの場合
					if cstr(HDate(k))=cstr(FDate(j)) then
						MCount(i,k)=MCount(i,k)+FCount(j)
						iGSum(k)=iGSum(k)+FCount(j)
						iMTSum(k)=iMTSum(k)+FCount(j)
						Exit for
					end if
				Next
			end if
		Next
	Next
	'最終行のデータを足しこむ
	For iLoop=0 to iDateCnt-1
		MCount(iHdRow,iLoop)=iGSum(iLoop)
	Next	
    '------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-S
    '------------------------------------------------------------------------------------------------------------
    'スマトフォンデータ整備
	sHdValue=""
	ReDim SCount(iSTKind-1,iDateCnt-1)
	ReDim iGSum(iDateCnt-1)
	ReDim iSTSum(iDateCnt-1)
	for iLoop=0 to iDateCnt-1
		iSTSum(iLoop)=0
	Next
	'携帯表示項目分ループ
	For i=0 to iSTKind-1
		'メニュー項目が変わった場合
		if sHdValue<>SPageTitle(i) and trim(SPageTitle(i))<>"" then
			'先頭行以外
			if i<>0 then
				for iLoop=0 to iDateCnt-1
					SCount(iHdRow,iLoop)=iGSum(iLoop)
				Next
			end if

			iHdRow=i
			for iLoop=0 to iDateCnt-1
				iGSum(iLoop)=0
			Next
			sHdValue=SPageTitle(i)
		end if
		'カウントクリア
		For iLoop=0 to iDateCnt-1
			SCount(i,iLoop)=0
		Next
		
		'ファイル行数分ループ
		For j=0 to iFileCnt-1
			'画面番号、作業番号が同じ場合
			If SPageNum(i)=FPageNum(j) and SWkNum(i)=FWkNum(j) then
				'日付分ループ
				For k=0 to iDateCnt-1
					'日付が同じデータの場合
					if cstr(HDate(k))=cstr(FDate(j)) then
						SCount(i,k)=SCount(i,k)+FCount(j)
						iGSum(k)=iGSum(k)+FCount(j)
						iSTSum(k)=iSTSum(k)+FCount(j)
						Exit for
					end if
				Next
			end if
		Next
	Next
	'最終行のデータを足しこむ
	For iLoop=0 to iDateCnt-1
		SCount(iHdRow,iLoop)=iGSum(iLoop)
	Next	
    '------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-E
    '------------------------------------------------------------------------------------------------------------
    
	'パソコンファイルへ出力
	For iLoop=0 to iTKind-1
		Response.Write  PageTitle(iLoop) &","
		Response.Write  SubTitle(iLoop) &","
		if trim(PageNum(iLoop))<>"" then
			Response.Write  PageNum(iLoop) & "-" & WkNum(iLoop) &","		
		else
			Response.Write ","
		end if
		iRSum=0
		'日付分ループ
		for j=0 to iDateCnt-1
			Response.Write  Count(iLoop,j) &","
			iRSum=iRSum+Count(iLoop,j)
		next
		Response.Write iRSum & ","
		Response.Write Chr(13) & Chr(10)
	Next

	'合計書き込み
	Response.Write  "合計,,,"
	iRSum=0
	for j=0 to iDateCnt-1
		Response.Write  iTSum(j) &","
		iRSum=iRSum+iTSum(j)
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)

	Response.Write Chr(13) & Chr(10)
    '-------------------------------------------------------------------------------------------------------------
	'携帯ファイル出力
	Response.Write "＜携帯電話＞,"
	Response.Write Chr(13) & Chr(10)

	'携帯タイトル行出力
	Response.Write "メニュー項目,"
	Response.Write "画面,"
	Response.Write "画面No," 
	
	For iLoop=0 to iDateCnt-1
		Response.Write HDate(iLoop) & ","
	next
	Response.Write "合計"

	Response.Write Chr(13) & Chr(10)

	'パソコンファイルへ出力
	For iLoop=0 to iMTKind-1
		Response.Write  MPageTitle(iLoop) &","
		Response.Write  MSubTitle(iLoop) &","
		if trim(MPageNum(iLoop))<>"" then
			Response.Write  MPageNum(iLoop) & "-" & MWkNum(iLoop) &","		
		else
			Response.Write ","
		end if
		iRSum=0
		'日付分ループ
		for j=0 to iDateCnt-1
			Response.Write  MCount(iLoop,j) &","
			iRSum=iRSum+MCount(iLoop,j)
		next
		Response.Write iRSum & ","
		Response.Write Chr(13) & Chr(10)
	Next

	'合計書き込み
	Response.Write  "合計,,,"
	iRSum=0
	for j=0 to iDateCnt-1
		Response.Write  iMTSum(j) &","
		iRSum=iRSum+iMTSum(j)
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)
	Response.Write Chr(13) & Chr(10)
    '-------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-S
    '-------------------------------------------------------------------------------------------------------------
    'スマートフォンファイル出力
	Response.Write "＜スマートフォン＞,"
	Response.Write Chr(13) & Chr(10)

	'携帯タイトル行出力
	Response.Write "メニュー項目,"
	Response.Write "画面,"
	Response.Write "画面No," 
	
	For iLoop=0 to iDateCnt-1
		Response.Write HDate(iLoop) & ","
	next
	Response.Write "合計"

	Response.Write Chr(13) & Chr(10)

	'パソコンファイルへ出力
	For iLoop=0 to iSTKind-1
		Response.Write  SPageTitle(iLoop) &","
		Response.Write  SSubTitle(iLoop) &","
		if trim(SPageNum(iLoop))<>"" then
			Response.Write  SPageNum(iLoop) & "-" & SWkNum(iLoop) &","		
		else
			Response.Write ","
		end if
		iRSum=0
		'日付分ループ
		for j=0 to iDateCnt-1
			Response.Write  SCount(iLoop,j) &","
			iRSum=iRSum+SCount(iLoop,j)
		next
		Response.Write iRSum & ","
		Response.Write Chr(13) & Chr(10)
	Next

	'合計書き込み
	Response.Write  "合計,,,"
	iRSum=0
	for j=0 to iDateCnt-1
		Response.Write  iSTSum(j) &","
		iRSum=iRSum+iSTSum(j)
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)
	Response.Write Chr(13) & Chr(10)
    '-------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-E
    '-------------------------------------------------------------------------------------------------------------
    
	'総合計書き込み
	Response.Write  "総合計,,,"
	iRSum=0
	for j=0 to iDateCnt-1
	    '2013-09-30 Y.TAKAKUWA Upd-S
		'Response.Write  iTSum(j)+iMTSum(j) &","
		'iRSum=iRSum+iTSum(j)+iMTSum(j)
		Response.Write  iTSum(j)+iMTSum(j)+iSTSum(j) &","
		iRSum=iRSum+iTSum(j)+iMTSum(j)+iSTSum(j)		
		'2013-09-30 Y.TAKAKUWA Upd-E
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)


	' ダウンロード終了
	Response.End

%>
