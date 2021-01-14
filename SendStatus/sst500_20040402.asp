<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst500.asp				_/
'_/	Function	:ステータス配信mail即時送信			_/
'_/	Date			:2004/01/07				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''Microsoft ADO用のadovbs.incにて提供されている
	Const adBoolean = 11
	Const adDBTimeStamp = 135
	Const adInteger = 3
	Const adChar = 129
	Const adParamInput = &H0001
	Const adParamReturnValue = &H0004

	'''セッションの有効性をチェック
	CheckLoginH

	'''送信しました画面にて「最新の情報に更新」でSubmitされた場合の対策
	if Session.Contents("SendMailSubmitted") = "False" then

		'''データ取得
		Dim USER, KIND, NUMBER, ErrCode, NewDelMode
		Dim Email1, Email2, Email3, Email4, Email5
		Dim UserName

		USER = Session.Contents("userid")
		KIND = Request.Form("ContORBL")
		NUMBER = Request.Form("ContBLNo")
		NewDelMode = Request.Form("Mode")
		ErrCode = 0

		'''サーバ日付の取得
		Dim DayTime
		getDayTime DayTime

		'''DB接続
		Dim ObjConn, ObjRS, StrSQL
		ConnDBH ObjConn, ObjRS

		'''指定コンテナ番号,ＢＬ番号の存在チェック
		Dim Num
		if KIND = 1 then '''コンテナ番号指定
			StrSQL = "SELECT count(ContNo) AS CNUM FROM ImportCont WHERE ContNo='" & NUMBER & "'"
		elseif KIND = 2 then	'''ＢＬ番号指定の場合。
			StrSQL = "SELECT count(BLNo) AS CNUM FROM BL WHERE BLNo='"& NUMBER & "'"
		else
			response.write("KIND error!")
			response.end
		end if
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			'''DB切断
			DisConnDBH ObjConn, ObjRS
			jumpErrorP "2","c104","01","ステータス配信mail即時送信","101","SQL:<BR>"&strSQL
		end if
		Num = ObjRS("CNUM")
		ObjRS.close

		'''指定されたコンテナ番号またはＢＬ番号が存在する場合
		if Num > 0 then
			'''ステータス配信先メールアドレスとログインユーザ名の抽出
			StrSQL = "SELECT TI.Email1, TI.Email2, TI.Email3, TI.Email4, TI.Email5, MU.FullName "
			StrSQL = StrSQL & " FROM TargetItems TI, mUsers MU "
			StrSQL = StrSQL & " WHERE TI.UserCode='" & USER & "' AND TI.UserCode=MU.UserCode "
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				'''DB切断
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "2","c104","01","ステータス配信mail即時送信","101","SQL:<BR>"&strSQL
			end if
			if ObjRS.EOF then		'''ログインユーザ用のステータス配信項目定義レコードが存在しない場合
				ObjRS.close
				ErrCode = 1
			else	'''ログインユーザ用のステータス配信項目定義レコードが存在する場合
				Email1=Trim(ObjRS("Email1"))
				Email2=Trim(ObjRS("Email2"))
				Email3=Trim(ObjRS("Email3"))
				Email4=Trim(ObjRS("Email4"))
				Email5=Trim(ObjRS("Email5"))
				UserName=Trim(ObjRS("FullName"))
				ObjRS.close

				if IsNull(Email1) and IsNull(Email2) and IsNull(Email3) and IsNull(Email4) and IsNull(Email5) then
				'''ログインユーザ用のステータス配信項目定義レコードが存在するが、メールアドレスが１つも登録されていない場合
					ErrCode = 2

				'''１つでもメールアドレスの定義が存在する場合、メール送信対象となるコンテナをすべて抽出する。
				else
					Dim ETA, TA, InTime, ListDate, DOStatus, DelPermitDate, FreeTime, FreeTimeExt
					Dim CYDelTime, DetentionFreeTime, ReturnTime
					Dim OLTICFlag, OLTICNo, OLTDateFrom, OLTDateTo
					Dim ContainerNumber, RcdNum, i
					Dim VslCode, VoyCtrl
					Dim sp, p0, p1, p2, p3, p4

					if KIND = 1 then		'''コンテナ番号指定の場合
						StrSQL = "SELECT VslCode, VoyCtrl FROM ImportCont "
						StrSQL = StrSQL & " WHERE ContNo='"& NUMBER &"'"
						StrSQL = StrSQL & " AND UpdtTime = (SELECT max(UpdtTime) FROM ImportCont WHERE ContNo='"& NUMBER &"') "

						ObjRS.Open StrSQL, ObjConn, 3, 1
						if err <> 0 then
							'''DB切断
							DisConnDBH ObjConn, ObjRS
							jumpErrorP "2","c104","01","ステータス配信mail即時送信","101","SQL:<BR>"&strSQL
						end if
						ReDim ContainerNumber(1), VslCode(1), VoyCtrl(1)
						ContainerNumber(0) = NUMBER
						VslCode(0) = ObjRS("VslCode")
						VoyCtrl(0) = ObjRS("VoyCtrl")
						RcdNum = 1
						ObjRS.close

					elseif KIND = 2 then		'''ＢＬ番号指定の場合、対象コンテナ番号をすべて取り出す
						StrSQL = "SELECT IC.VslCode, IC.VoyCtrl, IC.ContNo FROM BL, ImportCont IC "
						StrSQL = StrSQL & " WHERE BL.BLNo='"& NUMBER &"'"
						StrSQL = StrSQL & " AND BL.VslCode = IC.VslCode "
						StrSQL = StrSQL & " AND BL.VoyCtrl = IC.VoyCtrl "
						StrSQL = StrSQL & " AND BL.BLNo = IC.BLNo "
						StrSQL = StrSQL & " AND BL.UpdtTime = (SELECT max(BL.UpdtTime) FROM BL WHERE BL.BLNo='"& NUMBER &"') "

						ObjRS.Open StrSQL, ObjConn, 3, 1
						if err <> 0 then
							'''DB切断
							DisConnDBH ObjConn, ObjRS
							jumpErrorP "2","c104","01","ステータス配信mail即時送信","101","SQL:<BR>"&strSQL
						end if
						RcdNum = ObjRS.RecordCount
						ReDim ContainerNumber(RcdNum), VslCode(RcdNum), VoyCtrl(RcdNum)
						for i=0 to RcdNum-1
							ContainerNumber(i) = ObjRS("ContNo")
							VslCode(i) = ObjRS("VslCode")
							VoyCtrl(i) = ObjRS("VoyCtrl")
							ObjRS.MoveNext
						next
						ObjRS.close
					end if

					Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
					Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4

					'''SMTPサーバ名の設定
					svName   = "slitdns2.hits-h.com"
					attachedFiles = ""
					mailFlag1 = 0
					mailFlag2 = 0
					mailFlag3 = 0
					mailFlag4 = 0
					'''メール送信元アドレスの設定
					mailFrom = "mrhits@hits-h.com"
					mailTo = ""

					if IsNull(Email1) = false then
						mailTo = mailTo & Email1
						mailFlag1 = 1
					else
						mailFlag1 = 0
					end if

					if IsNull(Email2) = false then
						if mailFlag1 = 1 then
							mailTo = mailTo & vbtab & Email2
						else
							mailTo = mailTo & Email2
						end if
						mailFlag2 = 1
					else
						mailFlag2 = 0
					end if

					if IsNull(Email3) = false then
						if mailFlag1 = 1 or mailFlag2 = 1 then
							mailTo = mailTo & vbtab & Email3
						else
							mailTo = mailTo & Email3
						end if
						mailFlag3 = 1
					else
						mailFlag3 = 0
					end if

					if IsNull(Email4) = false then
						if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
							mailTo = mailTo & vbtab & Email4
						else
							mailTo = mailTo & Email4
						end if
						mailFlag4 = 1
					else
						mailFlag4 = 0
					end if

					if IsNull(Email5) = false then
						if mailFlag1 = 1  or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
							mailTo = mailTo & vbtab & Email5
						else
							mailTo = mailTo & Email5
						end if
					end if

					Dim rc, fp, fobj, tfile, sendTime
					Set ObjMail = Server.CreateObject("BASP21")

					Dim S_Flag

					'''各パラメータの格納用配列の宣言
					ReDim ETA(RcdNum), TA(RcdNum), InTime(RcdNum), ListDate(RcdNum), DOStatus(RcdNum)
					ReDim DelPermitDate(RcdNum), FreeTime(RcdNum), FreeTimeExt(RcdNum)
					ReDim CYDelTime(RcdNum), DetentionFreeTime(RcdNum), ReturnTime(RcdNum)
					ReDim OLTICFlag(RcdNum), OLTICNo(RcdNum), OLTDateFrom(RcdNum), OLTDateTo(RcdNum)
					ReDim rc(RcdNum), sendTime(RcdNum)

					'''搬出可否判定用ストアードプロシジャの呼び出しのための設定
					set sp = Server.CreateObject("ADODB.Command")
					set sp.ActiveConnection = ObjConn
					sp.CommandText = "{?=call DelPermitCheck(?,?,?)}"
					Set p0 = sp.CreateParameter("ret", adBoolean, adParamReturnValue)
					sp.Parameters.Append p0
					Set p1 = sp.CreateParameter("VslCode", adChar, adParamInput, 7)
					sp.Parameters.Append p1
					Set p2 = sp.CreateParameter("VoyCtrl", adInteger, adParamInput)
					sp.Parameters.Append p2
					Set p3 = sp.CreateParameter("ContNo", adChar, adParamInput, 12)
					sp.Parameters.Append p3

					'''抽出したコンテナの数だけループさせて、コンテナ毎に状態をメール送信する。
					for i=0 to RcdNum-1
						StrSQL = "SELECT VP.ETA, VP.TA, IC.InTime, CT.ListDate, IC.DOStatus, IC.DelPermitDate, IC.FreeTime, "
						StrSQL = StrSQL & " IC.FreeTimeExt, IC.CYDelTime, IC.DetentionFreeTime, IC.ReturnTime, "
						StrSQL = StrSQL & " IC.OLTICFlag, IC.OLTICNo, IC.OLTDateFrom, IC.OLTDateTo "
						StrSQL = StrSQL & " FROM VslPort VP, ImportCont IC, Container CT "
						StrSQL = StrSQL & " WHERE IC.ContNo='" & ContainerNumber(i) & "'"
						StrSQL = StrSQL & " AND VP.PortCode='JPHKT' "
						StrSQL = StrSQL & " AND IC.VslCode=VP.VslCode "
						StrSQL = StrSQL & " AND IC.VoyCtrl=VP.VoyCtrl "
						StrSQL = StrSQL & " AND CT.ContNo=IC.ContNo "
						StrSQL = StrSQL & " AND IC.VslCode=CT.VslCode "
						StrSQL = StrSQL & " AND IC.VoyCtrl=CT.VoyCtrl "

						ObjRS.Open StrSQL, ObjConn
						if err <> 0 then
							'''DB切断
							DisConnDBH ObjConn, ObjRS
							jumpErrorP "2","c104","01","ステータス配信mail即時送信","101","SQL:<BR>"&strSQL  & i
						end if

						ETA(i)=ObjRS("ETA")
						TA(i)=ObjRS("TA")
						InTime(i)=ObjRS("InTime")
						ListDate(i)=ObjRS("ListDate")
						DOStatus(i)=ObjRS("DOStatus")
						DelPermitDate(i)=ObjRS("DelPermitDate")
						FreeTime(i)=ObjRS("FreeTime")
						FreeTimeExt(i)=ObjRS("FreeTimeExt")
						CYDelTime(i)=ObjRS("CYDelTime")
						DetentionFreeTime(i)=ObjRS("DetentionFreeTime")
						ReturnTime(i)=ObjRS("ReturnTime")
						OLTICFlag(i)=ObjRS("OLTICFlag")
						OLTICNo(i)=ObjRS("OLTICNo")
						OLTDateFrom(i)=ObjRS("OLTDateFrom")
						OLTDateTo(i)=ObjRS("OLTDateTo")
						ObjRS.close

						Dim mailSubject, mailBody
						'''メールタイトルの設定
						if KIND = 1 then
							mailSubject = "輸入ステータスのお知らせ(コンテナ番号：" & ContainerNumber(i) & ")"
						elseif KIND = 2 then
							mailSubject = "輸入ステータスのお知らせ(ＢＬ番号：" & NUMBER & ")"
						end if

						'''メール本文の作成
						mailBody = ""
						mailBody = UserName & " 殿" & vbCrLf & vbCrLf
						mailBody = mailBody & "輸入ステータスのお知らせ　　　" & DayTime(0) & "年" & DayTime(1) & "月" & DayTime(2) & "日" & DayTime(3) & "時現在"  & vbCrLf & vbCrLf
						mailBody = mailBody & "●対象コンテナ" & vbCrLf
						mailBody = mailBody & "　" & ContainerNumber(i) & vbCrLf & vbCrLf
						mailBody = mailBody & "●ステータス" & vbCrLf

						mailBody = mailBody & "　(1)入港時間" & vbCrLf
						if IsNull(ETA(i)) = false then
							if Hour(ETA(i)) = 0 and Minute(ETA(i)) = 0 and Second(ETA(i)) = 0 then
								mailBody = mailBody & "　　予定・・・" & Year(ETA(i)) & "年" & Right("0"&Month(ETA(i)),2) & "月" & Right("0"&Day(ETA(i)),2) & "日" & vbCrLf
							else
								mailBody = mailBody & "　　予定・・・" & Year(ETA(i)) & "年" & Right("0"&Month(ETA(i)),2) & "月" & Right("0"&Day(ETA(i)),2) & "日 " & Right("0"&Hour(ETA(i)),2) & ":" & Right("0"&Minute(ETA(i)),2) & vbCrLf
							end if
						elseif IsNull(TA(i)) = false then
							if Hour(TA(i)) = 0 and Minute(TA(i)) = 0 and Second(TA(i)) = 0 then
								mailBody = mailBody & "　　完了・・・" & Year(TA(i)) & "年" & Right("0"&Month(TA(i)),2) & "月" & Right("0"&Day(TA(i)),2) & "日" & vbCrLf
							else
								mailBody = mailBody & "　　完了・・・" & Year(TA(i)) & "年" & Right("0"&Month(TA(i)),2) & "月" & Right("0"&Day(TA(i)),2) & "日 " & Right("0"&Hour(TA(i)),2) & ":" & Right("0"&Minute(TA(i)),2) & vbCrLf
							end if
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "　(2)ＣＹ搬入時間" & vbCrLf
						if IsNull(InTime(i)) = false then
							mailBody = mailBody & "　　" & Year(InTime(i)) & "年" & Right("0"&Month(InTime(i)),2) & "月" & Right("0"&Day(InTime(i)),2) & "日 " & Right("0"&Hour(InTime(i)),2) & ":" & Right("0"&Minute(InTime(i)),2) & vbCrLf
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "　(3)通関許可状況" & vbCrLf
						if IsNull(ListDate(i)) = false then
							mailBody = mailBody & "　　○　通関許可日=" & Year(ListDate(i)) & "年" & Right("0"&Month(ListDate(i)),2) & "月" & Right("0"&Day(ListDate(i)),2) & "日" & vbCrLf
						else
							mailBody = mailBody & "　　×" & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "　(4)ＤＯクリア状況" & vbCrLf
						if DOStatus(i) = "Y" then
							mailBody = mailBody & "　　○" & vbCrLf
						else
							mailBody = mailBody & "　　×" & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						'''搬出可否判定
						mailBody = mailBody & "　(5)搬出可否" & vbCrLf
						'''ＣＹ搬出されている場合は「済」を送信する  Modified 20040312
						if IsNull(CYDelTime(i)) = false then
							mailBody = mailBody & "　　済" & vbCrLf
						else
						'''ImportContテーブルのVslCode, VoyCtrl, ContNoが同じでBLNoだけが異なるレコードが存在する場合、
						'''当該レコードについても条件をクリアできているかチェックする。
							sp("VslCode") = VslCode(i)
							sp("VoyCtrl") = VoyCtrl(i)
							sp("ContNo") = ContainerNumber(i)
							'''ストアードプロシジャの呼び出し
							sp.Execute
							'''ストアードプロシジャの呼び出し結果の判定
							if sp("ret") = True then 
								mailBody = mailBody & "　　○　搬出可能日=" & Year(DelPermitDate(i)) & "年" & Right("0"&Month(DelPermitDate(i)),2) & "月" & Right("0"&Day(DelPermitDate(i)),2) & "日" & vbCrLf
							else
								mailBody = mailBody & "　　×" & vbCrLf
							end if
						end if
						mailBody = mailBody & vbCrLf

						''''''あと何日の表示をするのはFreeTimeExtまたはFreeTimeがmail即時送信実行日より将来の場合としている
						mailBody = mailBody & "　(6)デマレージフリータイム" & vbCrLf
						if IsNull(FreeTimeExt(i)) = false then
							if FreeTimeExt(i) > Date then
								mailBody = mailBody & "　　" & Year(FreeTimeExt(i)) & "年" & Right("0"&Month(FreeTimeExt(i)),2) & "月" & Right("0"&Day(FreeTimeExt(i)),2) & "日　あと" & DateDiff("d",Date,FreeTimeExt(i)) & "日" & vbCrLf
							else
								mailBody = mailBody & "　　" & Year(FreeTimeExt(i)) & "年" & Right("0"&Month(FreeTimeExt(i)),2) & "月" & Right("0"&Day(FreeTimeExt(i)),2) & "日" & vbCrLf
							end if
						elseif IsNull(FreeTime(i)) = false then
							if FreeTime(i) > Date then
								mailBody = mailBody & "　　" & Year(FreeTime(i)) & "年" & Right("0"&Month(FreeTime(i)),2) & "月" & Right("0"&Day(FreeTime(i)),2) & "日　あと" & DateDiff("d",Date,FreeTime(i)) & "日" & vbCrLf
							else
								mailBody = mailBody & "　　" & Year(FreeTime(i)) & "年" & Right("0"&Month(FreeTime(i)),2) & "月" & Right("0"&Day(FreeTime(i)),2) & "日" & vbCrLf
							end if
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "　(7)ＣＹ搬出時間" & vbCrLf
						if IsNull(CYDelTime(i)) = false then
							mailBody = mailBody & "　　" & Year(CYDelTime(i)) & "年" & Right("0"&Month(CYDelTime(i)),2) & "月" & Right("0"&Day(CYDelTime(i)),2) & "日 " & Right("0"&Hour(CYDelTime(i)),2) & ":" & Right("0"&Minute(CYDelTime(i)),2) & vbCrLf
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						'''あと何日の表示をするのはディテンションフリータイムが将来となる場合としている。
						'''また、DetentionFreeTimeに「0」が設定されている場合、すなわち返却予定日数として
						'''「未入力」「５日以上」または「リストオフ」が指定されている場合、あと何日の表示はしない。
						mailBody = mailBody & "　(8)ディテンションフリータイム" & vbCrLf
						if not IsNull(DetentionFreeTime(i)) and not IsNull(CYDelTime(i)) then
							if DateAdd("d",DetentionFreeTime(i),DateValue(CYDelTime(i)))>Date then
								mailBody = mailBody & "　　搬出日から" & Trim(DetentionFreeTime(i)) & "日以内　あと" & DateDiff("d",Date,DateAdd("d",DetentionFreeTime(i),DateValue(CYDelTime(i)))) & "日" & vbCrLf
							else
								mailBody = mailBody & "　　搬出日から" & Trim(DetentionFreeTime(i)) & "日以内" & vbCrLf
							end if
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "　(9)空コン返却状況" & vbCrLf
						if IsNull(ReturnTime(i)) = false then
							mailBody = mailBody & "　　○　空コン返却日時=" & Year(ReturnTime(i)) & "年" & Right("0"&Month(ReturnTime(i)),2) & "月" & Right("0"&Day(ReturnTime(i)),2) & "日 " & Right("0"&Hour(ReturnTime(i)),2) & ":" & Right("0"&Minute(ReturnTime(i)),2) & vbCrLf
						else
							mailBody = mailBody & "　　×" & vbCrLf
						end if

						'''メール送信処理

						rc(i)=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
						sendTime(i)=Now
					Next

					for i=0 to RcdNum-1
						if rc(i)="" then
							S_Flag = 0
						else
							S_Flag = 1
							exit for
						end if
					next

					if S_Flag = 0 then		'''メール送信成功
						'''削除画面からmail即時送信をやった場合TargetContainersテーブルの最終送信日時を更新する。
						'''新規登録画面からmail即時送信をやった場合は対象レコードがまだinsertされていないので最終送信日時の更新は不要。
						if NewDelMode = 2 then
							StrSQL = "UPDATE TargetContainers SET UpdtTime='" & Now() & "', UpdtPgCd='STATUS01',"
							StrSQL =  StrSQL & " UpdtTmnl='" & USER & "', LatestSentTime='" & Now() & "'"
							if KIND = 1 then		'''対象がコンテナ番号
								StrSQL =  StrSQL & " WHERE ContNo='" & NUMBER & "' AND UserCode='" & USER & "'"
							elseif KIND = 2 then		'''対象がＢＬ番号
								StrSQL =  StrSQL & " WHERE BLNo='" & NUMBER & "' AND UserCode='" & USER & "'"
							end if
							StrSQL =  StrSQL & " AND Process='R' OR Process='N'"
						end if
						ObjConn.Execute(StrSQL)
						if err <> 0 then
							Set ObjRS = Nothing
							jumpErrorPDB ObjConn,"1","c104","14","ステータス配信mail即時送信","104","SQL:<BR>"&StrSQL
						end if

						'''ログ出力
						WriteLogH "c104", "ステータス配信mail即時送信","01",""
						ErrCode = 0

					else		'''メール送信失敗
						fp = Server.MapPath("./mailerror") & "\error.txt"
						set fobj = Server.CreateObject("Scripting.FileSystemObject")

						for i=0 to RcdNum-1
							if rc(i)<>"" then
								if fobj.FileExists(fp) = True then
									set tfile = fobj.OpenTextFile(fp,8)
								else
									set tfile = fobj.CreateTextFile(fp,True,False)
								end if
								tfile.WriteLine sendTime(i) & " " & rc(i)
								tfile.Close
								ErrCode = 8
							end if
						next

					end if		'''メール送信成功、失敗処理の終わり
				end if		'''メールアドレスが１つでも定義されている場合の処理の終わり
			end if		'''ステータス配信項目が定義されている場合の処理の終わり
		else		'''指定されたコンテナ番号、ＢＬ番号が存在しない
			ErrCode = 9
		end if

		'''DB接続解除
		DisConnDBH ObjConn, ObjRS
		'''エラートラップ解除
		on error goto 0

		Session.Contents("SendMailSubmitted") = "True"


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信mail即時送信</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信mail即時送信結果画面--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst500" method="POST">
	<TR><TD>　</TD></TR>
<% if ErrCode=0 then %>
	<TR>
		<TD align="center">
			メール送信しました。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
<% elseif ErrCode=1 or ErrCode=2 then %>
	<TR>
		<TD align="center">
			メール送信先が設定されていません。<BR>「設定」メニューにてメールアドレスを登録してください。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
<% elseif ErrCode=8 then %>
	<TR>
		<TD align="center">
			メール送信に失敗しました。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
<% elseif ErrCode=9 then %>
	<TR>
		<TD align="center">
			指定されたコンテナ番号またはＢＬ番号は存在しません。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
<% end if %>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>

<%''' if Session.Contents("SendMailSubmitted") = "False"のelse処理 %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信mail即時送信</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信mail即時送信結果画面--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst500" method="POST">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			処理は既に完了しています。<BR><BR><BR>
		</TD>
	</TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("SendMailSubmitted") = "False"のendif処理 %>
<% end if %>
