<%
'**********************************************
'  【プログラムＩＤ】　: dml000D
'  【プログラム名称】　: ドライバ承認
'
'  （変更履歴）
'   2013-05-10   Y.TAKAKUWA   作成
'   2013-06-27   Y.TAKAKUWA   ロギングの機能を追加
'**********************************************
Option Explicit
Response.Expires = 0
'HTTPコンテンツタイプ設定
Response.ContentType = "text/html; charset=Shift_JIS"
Response.AddHeader "Pragma", "no-cache" 
%>
<%'**********************************************
  '共通の前提処理
  '共通関数  (Commonfunc.inc)
%>
<!--#include File="Common.inc"-->


<%
	'**********************************************

	'セッションの有効性をチェック
	CheckLoginH
	
	'ユーザデータ所得
	dim USER, COMPcd
	dim v_GamenMode
	dim v_SearchType      '2013-06-28  Y.TAKAKUWA  Add
	dim v_DataCnt2
		
	dim Num2	
	dim strOrder2
	dim FieldName2	
	dim ObjRS2,ObjConn2
	
	dim wk
	dim i,x
	dim v_ItemName
	dim v_ItemValue
	dim abspage, pagecnt,reccnt	
	
	dim Arr_DriverID()
	dim Arr_Check()
	
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	dim Arr_SendStat()
	'2013-06-27 Y.TAKAKUWA Add-S
	dim v_AdminMailAddress
	dim v_SendDetail
	'2013-06-27 Y.TAKAKUWA Add-E
	const gcPage = 10
	
	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' 再描画前の項目取得
   	'----------------------------------------			
	call LfGetRequestItem
	Call getAdminMailAddress()
			
	If v_GamenMode = "AP" then
	  call LfSendEmail()
	end if
	'2013-06-27 Y.TAKAKUWA Add-S
	If v_GamenMode = "D" Then
	  call LfDeleteDriverInfo()
	End If
	'2013-06-27 Y.TAKAKUWA Add-E
	Call getDriverInfo()
	
Function LfGetRequestItem()
   
	If Request.form("Gamen_Mode") = "" then
	  v_GamenMode = Request.QueryString("GamenMode")
	else
	  v_GamenMode = Request.form("Gamen_Mode")
	end if
	
	If Request.Form("SearchApprovalType") = "" Then
	  v_SearchType = Request.QueryString("SearchType")
	Else
	  v_SearchType = Request.Form("SearchApprovalType")
	End If
	if Trim(v_GamenMode) = "PS" then
	  'v_DriverInfo = Request.QueryString("driverInfo")
      v_DataCnt2 = Request.QueryString("DataCnt")
	else
	  v_DriverInfo = Request.Form("driverInfo")
      v_DataCnt2 = Request.form("DataCnt2")
    end if
    If v_DataCnt2 = "" then
      v_DataCnt2 = 0
    end if
	ReDimension(v_DataCnt2)
	
    v_SendDetail = Request.form("Send_Detail") 
    
	For i = 1 to (v_DataCnt2) - 1 
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
        Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
	Next
End Function

Function ReDimension(index)
   Redim Arr_Check(index)
   Redim Arr_DriverID(index)
End Function

Function getDriverInfo()
    dim StrSQL
 
    ConnDBH ObjConn2, ObjRS2
    
    StrSQL = "SELECT * FROM LomDriver "
    StrSQL = StrSQL & " WHERE "
    StrSQL = StrSQL & " HiTSUserID = '" & USER & "'"
    '2013-06-28 Y.TAKAKUWA Add-S
    If Trim(v_SearchType) = "S2" Then
      StrSQL = StrSQL & " AND (AcceptStatus = '1' OR AcceptStatus = '2')"
    Else
      'StrSQL = StrSQL & " AND AcceptStatus <> '1' AND AcceptStatus <> '3' " '2013-06-27 Y.TAKAKUWA Del
      StrSQL = StrSQL & " AND (AcceptStatus = '' OR AcceptStatus = NULL OR AcceptStatus='0')"
    End If
    StrSQL = StrSQL & " ORDER BY LomDriver.LoDriverID "
    '2013-06-28 Y.TAKAKUWA Add-E
    ObjRS2.PageSize = 50
	ObjRS2.CacheSize = 50
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2

	Num2 = ObjRS2.recordcount	
	
	if Num2 > 50 then 
		If CInt(Request("pagenum2")) = 0 Then
			ObjRS2.AbsolutePage = 1
		Else
			If CInt(Request("pagenum2")) <= ObjRS2.PageCount Then
				ObjRS2.AbsolutePage = CInt(Request("pagenum2"))
			Else
				ObjRS2.AbsolutePage = 1
			End If
		End If		 
	end if
	
	if err <> 0 then
		DisConnDBH ObjConn2, ObjRS2	'DB切断
		jampErrerP "2","b301","01","ロックオン事前情報","102","SQL:<BR>" & StrSQL & err.description & Err.number
		Exit Function
	end if			
	'エラートラップ解除
    on error goto 0	

End Function

'2013-06-27 Y.TAKAKUWA Add-S
Function getAdminMailAddress()   
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
    
   ConnDBH ObjConnLO, ObjRSLO	
   StrSQL = "SELECT * FROM mUsers " 
   StrSQL = StrSQL & " WHERE UserCode = '" & USER & "' "
   StrSQL = StrSQL & " ORDER BY UserCode"
   ObjRSLO.Open StrSQL, ObjConnLO
   While Not ObjRSLO.EOF
     v_AdminMailAddress = ObjRSLO("MailAddress")
     ObjRSLO.MoveNext
   Wend

   ObjRSLO.Close
   DisConnDBH ObjConnLO, ObjRSLO	'DB切断
   
End Function
'2013-06-27 Y.TAKAKUWA Add-E

'2013-06-27 Y.TAKAKUWA Add-S
Function LfDeleteDriverInfo()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	'2013-09-26 Y.TAKAKUWA Add-S
    WriteLogH "b502", "ドライバ承認（非承認）実行", "01", ""
    '2013-09-26 Y.TAKAKUWA Add-E
    
	For i = 1 to v_DataCnt2-1
      If UCase(Trim(Arr_Check(i))) = "ON" Then
        'QUERY VALUES FOR Delete
        StrSQL = "SELECT * FROM LomDriver WHERE LoDriverID ='" & Arr_DriverID(i)  & "'"
        'response.Write StrSQL                                                
        ObjRSLO.Open StrSQL, ObjConnLO
        If ObjRSLO.recordcount > 0 then
            StrSQL = " DELETE FROM LomDriver WHERE "
            StrSQL = StrSQL & "LoDriverID='" & Trim(Arr_DriverID(i)) & "'"        
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		    end if
	    end if
	    ObjRSLO.Close
	    
	    StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID='" & USER & "' AND LoDriverID ='" & Arr_DriverID(i)  & "'"
	    ObjRSLO.Open StrSQL, ObjConnLO
	    If ObjRSLO.recordcount > 0 then
		'DELETE ALSO IN GROUP
		  StrSQL = " DELETE FROM LoGroupeDriver WHERE "
          StrSQL = StrSQL & " HiTSUserID='" & USER & "'" 
          StrSQL = StrSQL & " AND LoDriverID='" & Trim(Arr_DriverID(i)) & "'"  
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
		    Set ObjRSLO = Nothing				
		    jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		  end if
		end if
		ObjRSLO.Close
		
      end if
    Next

    DisConnDBH ObjConnLO, ObjRSLO	'DB切断
    
End function
'2013-06-27 Y.TAKAKUWA Add-E

Function LfSendEmail()

  Dim ObjConnLO, ObjRSLO
  Dim ErrFlg
  Dim StrSQL

  Dim v_LoDriverName
  Dim v_LoDriverID
  Dim v_LoDriverPassword
  Dim v_LoEmailAddress
  Dim v_ErrMsg	
  ConnDBH ObjConnLO, ObjRSLO	
  WriteLogH "", "", "", ""
  '2013-09-26 Y.TAKAKUWA Add-S
  WriteLogH "b502", "ドライバ承認（非承認）実行", "01",""
  '2013-09-26 Y.TAKAKUWA Add-E
  Dim objMail
  Dim rc	
  
  Dim svName 
  Dim mailTo 
  Dim mailFrom 
  Dim mailSubject
  Dim strBody
  Dim attachedFiles
  Dim sendDetailArr

  If Trim(v_SendDetail) <> "" Then
    sendDetailArr = Split(v_SendDetail,"/")
  End If
  If Ubound(sendDetailArr) > 0 Then
  If Trim(sendDetailArr(1)) <> "" Then
  'For i = 1 to v_DataCnt2-1                         '2013-06-27 Y.TAKAKUWA Del
  'For i = 1 to UBound(sendDetailArr)                 '2013-06-27 Y.TAKAKUWA Add
    'If UCase(Trim(Arr_Check(i))) = "ON" Then        '2013-06-27 Y.TAKAKUWA Del
      
      'ドライバ内容取込-S
      StrSQL = "SELECT * FROM LomDriver WHERE LoDriverID ='" & sendDetailArr(0)  & "'"
      ObjRSLO.Open StrSQL, ObjConnLO
      If ObjRSLO.recordcount > 0 Then
        v_LoDriverName = Trim(ObjRSLO("LoDriverName")) 
        v_LoDriverID  = Trim(ObjRSLO("LoDriverID"))
        v_LoDriverPassword  = MID(Trim(ObjRSLO("LoDriverPW")),1,1) & String(Len(Trim(ObjRSLO("LoDriverPW"))) -1,"*")
        v_LoEmailAddress  = Trim(ObjRSLO("MailAddress"))
      End If 
      'ドライバ内容取込-E
      
      
      svName = "153.150.17.106"
      'svName = "221.186.126.66"
      'svName = "192.168.17.243"
      mailTo = Trim(sendDetailArr(1))
      If Trim(sendDetailArr(2)) <> "" Then
        mailTo = mailTo & vbtab & sendDetailArr(2) 'Trim(v_LoEmailAddress)  '2013-06-27 Y.TAKAKUWA Upd
      End If
      mailFrom = "mrhits@hits-h.com" 
      mailSubject = "HiTSドライバ承認"
      attachedFiles = ""
  
      'メール内容-S
      strBody = v_LoDriverName & " 様" & vbCrLf & vbCrLf
      strBody = strBody & "HiTS利用ドライバとして承認されました。" & vbCrLf  
      strBody = strBody & "　　ドライバID＝" & v_LoDriverID  & vbCrLf 
      strBody = strBody & "　　パスワード＝" & v_LoDriverPassword & vbCrLf & vbCrLf
      strBody = strBody & "次の手順で専用アプリをインストールしてください。" & vbCrLf & vbCrLf
      strBody = strBody & "○Android携帯の場合" & vbCrLf
      strBody = strBody & "１．「提供元不明のアプリ」のインストールを許可してください。" & vbCrLf
      strBody = strBody & "　(例）設定→アプリケーション設定　内に項目があります" & vbCrLf & vbCrLf
      strBody = strBody & "２．下記のURLをクリックして専用アプリをダウンロードしてください。" & vbCrLf & vbCrLf
      strBody = strBody & "３．ダウンロードが終わりましたら、インストールを行ってください。" & vbCrLf
      strBody = strBody & "　(例）通知パネルにある「HiTS.apk」をタップして、「インストール」をタップしてください。" & vbCrLf & vbCrLf
      strBody = strBody & "４．ダウンロードアプリの一覧に「HiTS」アイコンが追加されます。" & vbCrLf
      strBody = strBody & "　アプリを起動し、ドライバ登録で設定されたドライバIDとパスワードを入力すればログインできます。" & vbCrLf & vbCrLf
      strBody = strBody & "５．インストールが終了後、必要に応じて「提供元不明のアプリ」のインストールを不許可に戻してください。" & vbCrLf & vbCrLf
      strBody = strBody & "　https://www.hits-h.com/sp/android/download.html" & vbCrLf & vbCrLf & vbCrLf
      strBody = strBody & "○iPhone携帯の場合" & vbCrLf
      strBody = strBody & "１．下記のURLをタップしてください。" & vbCrLf & vbCrLf
      strBody = strBody & "２．アプリダウンロード用の認証画面が表示されますので、下記のIDとパスワードを入力してください。" & vbCrLf
      strBody = strBody & "　　ユーザ名　: hits （全て小文字）" & vbCrLf
      strBody = strBody & "　　パスワード: Logi-app （最初のみ大文字）" & vbCrLf 
      strBody = strBody & "　　※ドライバ登録時のものではありません。"& vbCrLf & vbCrLf
      strBody = strBody & "３．入力後、「インストール」をタップしてください。" & vbCrLf & vbCrLf
      strBody = strBody & "４．インストール完了後、ホーム画面に「HiTS」アイコンが追加されます。" & vbCrLf
      strBody = strBody & "　アプリを起動し、ドライバ登録で設定されたドライバIDとパスワードを入力すればログインできます。" & vbCrLf & vbCrLf
      strBody = strBody & "　https://www.hits-h.com/sp/iOS/download.html" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
      strBody = strBody & "※このメールに関してのお問い合わせは、下記URLのページにあります連絡先まで、ご連絡をお願いいたします。" & vbCrLf
      strBody = strBody & "　http://www.hits-h.com/request.asp" & vbCrLf & vbCrLf
      strBody = strBody & "※このメールにお心当たりのない方は、他の方が間違って本サービスにメールアドレスを登録された可能性があります。" &vbCrLf
      strBody = strBody & "　お手数をおかけしますが、このメールを破棄していただけますよう、お願いいたします。"
      strBody = Server.HTMLEncode(strBody)
      'メール内容-E
      
      If svName <> "" And mailTo <> "" Then
        Set ObjMail = Server.CreateObject("BASP21")
        rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, strBody, attachedFiles)

        if rc <> "" then
            StrSQL = " UPDATE LomDriver SET "
            '2013/07/30 Upd-S Fujiyama メール送信エラーは正常扱いにする
            'StrSQL = StrSQL & "AcceptStatus='2', "                          'AcceptStatus
            StrSQL = StrSQL & "AcceptStatus='1', "                          'AcceptStatus
            '2013/07/30 Upd-E Fujiyama メール送信エラーは正常扱いにする
            StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
            StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
            StrSQL = StrSQL & "UpdtTmnl='" & USER & "' "                   'UpdtTmnl
            'StrSQL = StrSQL & "MailAddress='" & sendDetailArr(1) & "' "     'MailAddress
            StrSQL = StrSQL & "WHERE LoDriverID='" & Trim(sendDetailArr(0)) & "'"      '2013-06-27 Y.TAKAKUWA Upd
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
	          Set ObjRSLO = Nothing				
	          jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	        end if
	        v_ErrMsg = "失敗"
        else
          StrSQL = " UPDATE LomDriver SET "
          StrSQL = StrSQL & "AcceptStatus='1', "                          'AcceptStatus
          StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
          StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
          StrSQL = StrSQL & "UpdtTmnl='" & USER & "', "                    'UpdtTmnl
          StrSQL = StrSQL & "MailAddress='" & sendDetailArr(1) & "' "     'MailAddress
          StrSQL = StrSQL & "WHERE LoDriverID='" & Trim(sendDetailArr(0)) & "'"     '2013-06-27 Y.TAKAKUWA Upd
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
      	    Set ObjRSLO = Nothing				
	        jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	      end if
	      v_ErrMsg = "成功"
        end if
        ObjRSLO.Close                '2013-06-27 Y.TAKAKUWA  Del
      Else
          StrSQL = " UPDATE LomDriver SET "
          '2013/07/30 Upd-S Fujiyama メール送信エラーは正常扱いにする
          'StrSQL = StrSQL & "AcceptStatus='2', "                          'AcceptStatus
          StrSQL = StrSQL & "AcceptStatus='1', "                          'AcceptStatus
          '2013/07/30 Upd-E Fujiyama
          StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
          StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
          StrSQL = StrSQL & "UpdtTmnl='" & USER & "' "                    'UpdtTmnl
          'StrSQL = StrSQL & "MailAddress='" & sendDetailArr(1) & "' "     'MailAddress
          StrSQL = StrSQL & "WHERE LoDriverID='" & Trim(sendDetailArr(0)) & "'"      '2013-06-27 Y.TAKAKUWA Upd
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
	        Set ObjRSLO = Nothing				
	        jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	      end if
	      v_ErrMsg = "失敗"
      End If
      '2013-06-28 Y.TAKAKUWA Add-S
      If Trim(sendDetailArr(2)) <> "" Then
        If Trim(v_AdminMailAddress) <> Trim(sendDetailArr(2)) Then
          'Response.Write "Admin updated:" & v_AdminMailAddress
          StrSQL = "SELECT * FROM mUsers WHERE UserCode ='" & USER  & "'"
          ObjRSLO.Open StrSQL, ObjConnLO
          If ObjRSLO.recordcount > 0 Then
            StrSQL = " UPDATE mUsers SET "
            StrSQL = StrSQL & "MailAddress='" & sendDetailArr(2) & "' "     'MailAddress
            StrSQL = StrSQL & "WHERE UserCode='" & USER & "'"
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
	          Set ObjRSLO = Nothing				
	          jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	        Else
	          v_AdminMailAddress = sendDetailArr(2)
	        end if
          End If
        End If
      End If
      '2013-06-28 Y.TAKAKUWA Add-E
    'End If
    
  'Next
  End If
  End If
  'Set objMsg = Nothing
  DisConnDBH ObjConnLO, ObjRSLO	'DB切断

End Function

function LfPutPage(rec,page,pagecount,link)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo
	dim intNextFlag
	dim strParam
	PageIndex=0
	PageWkNo=0	
	if rec > 0 then	

		if pagecount<page then
			page=pagecount
		end if
		'ページIndexを設定
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		PageWkNo=((gcPage*PageIndex)+1)-gcPage
		
		
		'先頭ページが0より小さい場合は1を設定
		if PageWkNo<=0 Then
			PageWkNo=0
		End If
        

		'パラメータ設定
		
	    'strParam="&InOutF=" & v_InOutFlag
		strParam=""
		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & "&SearchType=" & v_SearchType & """>最初へ</a>"
			response.write "| &nbsp;"
			if PageWkNo<>0 Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&SearchType=" & v_SearchType & """>前へ</a>"
			Else
				response.write "<font style='color:#FFFFFF;'>前へ</font>"
			End If
		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        		
		'--- インデックス
		'ページが1ページ以上存在する場合
		if pagecount>1 then
			response.write "| &nbsp;"

			'指定ページ数分ループ
			for i=1 to gcPage
				'ページ数算出
				PageWkNo=(gcPage*PageIndex)+i

				'ページが全ページより大きい場合は処理中断
				if pagecount< PageWkNo then
					PageWkNo=PageWkNo-1
					exit for
				end if
				'現在選択されているページの場合
				if PageWkNo=page then
					response.write "&nbsp;" & PageWkNo 
				else
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam  & "&SearchType=" & v_SearchType & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
					
		if page<pagecount-1 then
			PageWkNo=PageWkNo+1
			If PageWkNo<=LastPage Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>次へ</a>"'
			Else
				response.write "<font style='color:#FFFFFF;'>次へ</font>"
			End If
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & """>最後へ</a>"'            
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function
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
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR>



<STYLE>
body {
    background-image:url('../gif/back.gif');
    margin:0;
    padding:0;
}

th.hlist {
	position: relative;
}
th {
    border-width: 1px 1px 1px 1px;
    padding: 4px;
    background-color: #ffcc33;
}
SELECT.chr {
    BACKGROUND-COLOR: #ffffff;
    BORDER-BOTTOM: #ffffff 1px solid;
    BORDER-LEFT: #002f7b 0px solid;
    BORDER-RIGHT: #ffffff 0px solid;
    BORDER-TOP: #ffffff 0px solid;
    COLOR: black;
    FONT-FAMILY: 'ＭＳ ゴシック';
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
    PADDING-BOTTOM: 2px;
    PADDING-LEFT: 1px;
    PADDING-RIGHT: 2px;
    PADDING-TOP: 3px;
    TEXT-ALIGN: left
}
table {
    border-width: 0px 1px 1px 0px;
}
DIV.center {
	text-align:center;
}
DIV.BDIV1 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
DIV.BDIV2 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
thead tr {
    position: relative;
    top: expression(this.offsetParent.scrollTop);
}
#loading2 {
	font:bold 12px Verdana;
	color:red;
	position:absolute; 
	top:220px; 
	left:390px;
	width:300px;
	height:30px; 
	z-index:69;
	font-size:12pt;
	border:0px;
	vertical-align: middle;
}
#footer {
 position: fixed;
 top: 100%;
 width: 100%;
}

</style>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT Language="JavaScript">

function finit(){
	//データ引継ぎ設定  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";
    if("<%=v_Searchtype %>"=="S1"){
      document.getElementById("chk1").checked=true;
    }
    else{
      if("<%=v_Searchtype %>"=="S2"){
         document.getElementById("chk2").checked=true;
      }
    }
}

//データが無い場合の表示制御
function view(){
	var sortedHeight;
	sortedHeight = 0;
	var vHeight;
	var obj2=document.getElementById("BDIV2");
	var rowHeight;
	
	if('<%=Num2%>'!='0'){
	  var rowHeightThead = getRowHeightThead();
	  var rowHeightTbody = getRowHeightTbody();
	  
	
	  if(rowHeightThead > 0){
	    rowHeightThead=rowHeightThead
	  }
	  if(rowHeightTbody > 0){
	    rowHeight=rowHeightTbody*15
	  }
	  rowHeight=rowHeight+rowHeightThead
    }
    else{
      rowHeight = 0;
      rowHeight=23*15;
    }
    
	if((document.body.offsetWidth-50) < 50){
		obj2.style.width=50;
		obj2.style.overflowX="auto";	 
	}else if((document.body.offsetWidth-50)  < 813){
		//obj2.style.width=document.body.offsetWidth-200;
		obj2.style.width=document.body.offsetWidth-220;
		obj2.style.overflowX="auto";
	}else{
		obj2.style.width=document.body.offsetWidth-220;
		obj2.style.overflowX="auto";
	}	
	
	if((document.body.offsetHeight-rowHeight) < 100){ 
	    if(obj2.clientWidth<obj2.scrollWidth)
	    {
	      obj2.style.height = 40;
		  obj2.style.overflowY = "auto";
	    }
	    else{
	      obj2.style.height = 25;
		  obj2.style.overflowY = "auto";
		}
	}else if((document.body.offsetHeight-rowHeight) < 395){
	    vHeight = rowHeight + 68;
		obj2.style.height = document.body.offsetHeight-vHeight;
		obj2.style.overflowY = "auto";
	}else{
		if(obj2.clientWidth < obj2.scrollWidth)
        {
           obj2.style.height = rowHeight+17;
           obj2.style.overflowY="auto";
        }
        else{
           obj2.style.height = rowHeight;
           obj2.style.overflowY="auto";
        } 
	}
    
	var obj3=document.getElementById("BDIV3");

}

function getRowHeightThead()
{
  var oRows = document.getElementById('TBInOut').getElementsByTagName('thead');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}
function getRowHeightTbody()
{
  var oRows = document.getElementById('TBInOut').getElementsByTagName('td');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}

function LockOnReg(){
	document.frm.Gamen_Mode.value = "I";
    document.frm.submit();
}


function fRSearch(){
	document.frm.Gamen_Mode.value = "S";
    document.frm.submit();
}

function fDelete(){
    var i;
    var chkFlag;
    chkFlag = 0;

    for(i=1;i<=parseInt(document.frm.DataCnt2.value)-1; i++){
      if(document.getElementById("checkbox" + i.toString()).checked == true){
         chkFlag = 1;
      }
    }
    
    if(chkFlag == 1){
      
      var show_modal = showModalDialog("dmlModal.asp?ActionType=D", window, "dialogWidth:300px; dialogHeight:175px; center:1; scroll: no;");
      if(show_modal){
        document.frm.Gamen_Mode.value = "D"
        document.frm.submit();
      }
    }
}

function fApproval(){
    var i;
    var chkFlag;
    var selectedCnt;
    var gTitle;
    var rowCnt;
    var show_modal;
    //show_modal = null;
    rowCnt = 0;
    chkFlag = 0;
    
    for(i=1;i<=parseInt(document.frm.DataCnt2.value)-1; i++){
      if(document.getElementById("checkbox" + i.toString()).checked == true){
         chkFlag = 1;
         selectedCnt = i;
         rowCnt = rowCnt + 1;
      }
    }

    if(rowCnt > 1){
       alert("複数選択は許可されていません");
       return;
    }
    
    if(chkFlag == 1){
      if(document.getElementById("chk1").checked==true){
        gTitle = "S1";
      }
      else{
        gTitle = "S2";
      }
      show_modal = showModalDialog("dmlModal.asp?ActionType=S&SendTo=" + document.getElementById("InputMailAddress" + selectedCnt.toString()).value.toString() + "&DriverID=" + document.getElementById("InputDriverID" + selectedCnt.toString()).value.toString() + "&AdminMailAddress=" + '<%=v_AdminMailAddress%>' + "&GamenTitle=" + gTitle.toString(), window, "dialogWidth:350px; dialogHeight:210px; center:1; scroll: no;");
      if(show_modal != false && (typeof(show_modal) != 'undefined' && show_modal != null)){
        document.frm.Gamen_Mode.value = "AP"
        document.frm.Send_Detail.value = show_modal;
        document.frm.submit();
      }
      
      
      
    }
}

</SCRIPT>

<script type="text/vbscript">
//Public Sub Approval_onclick()
//  Dim chkFlag
//  Dim x
//  Dim i
//  
//  chkFlag = 0
//  
//  for i = 1 to CInt(document.frm.DataCnt2.value-1)
//     If document.frm.elements("chkInOut" + CStr(i)).checked then
//       chkFlag = 1
//     end if
//  Next
//  
//  if chkFlag=1 then
//    x=MsgBox("選択したドライバを承認します。よろしいですか？",4,"Confirm")
//    if x = vbYes then
//      document.frm.Gamen_Mode.value = "AP"
//      document.frm.submit()
//    end if
//  end if
//End Sub
</script>

</HEAD>
<BODY onLoad="finit();view();" onResize="view();">
<form name="frm" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
   <td rowspan="2">
	 <IMG height=73 src="Image/DriverApproval.png" width=507>
   </td>
   <td height="19" bgcolor="#000099" align="right">
	 <IMG height=19 src="../gif/logo_hits_ver2_1.gif">
   </td>
</tr>
<tr>
   <td align="right" width="100%" height="45">
		
   </td>
</tr>
</table>

<table width="100%" height="82%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="120" nowrap>&nbsp;</td><td>
  <div id="BDIV3" style="width: 300px; height: 100%; padding-top:32px;">
  <!--Hidden Values Start-->
  <INPUT type=hidden name="Gamen_Mode" size="9" readonly tabindex= -1>
  <INPUT type=hidden name="Send_Detail" size="9" readonly tabindex= -1>  <!--2013-06-27 Y.TAKAKUWA Add-->
  <!--Hidden Values End-->
  <!--Added Start-->
  <!--<div style="width:150px; padding:10px;background-color:#FFCCFF; text-align:center;">承認待ちドライバ一覧</div>-->
  <div style="margin-left:30px;">
  <!--2013-06-28 Y.TAKAKUWA Add-S-->
  <table>
  <tr>
     <td>
        <input type=radio name="SearchApprovalType" id="chk1" value="S1" checked=true onclick="fRSearch();"/>承認待ちドライバ一覧
     </td>
  </tr>
  <tr>
     <td>
        <input type=radio name="SearchApprovalType" id="chk2" value="S2" onclick="fRSearch();"/>承認ドライバ一覧
     </td>
  </tr>
  </table>
  <!--2013-06-28 Y.TAKAKUWA Add-E-->
  <table border="0" cellpadding="0" cellspacing="0">
    <tr align=right nowrap>
      <td width="100%" height="30" align=right nowrap>
          <table border="0" cellpadding="0" cellspacing="0">
            <tr>
		      <td width="100%" align="center" nowrap>
		      <!--Page Pagination Start-->
		        <%
				  If Num2 > 0 Then
					abspage = ObjRS2.AbsolutePage
					pagecnt = ObjRS2.PageCount
					call LfPutPage(Num2,abspage,pagecnt,"pagenum2")
				  End If
			     %>
		      <!--Page Pagination End-->
		      </td>
		    </tr>
		  </table> 
      </td>
    </tr>
	<tr>
		<!--Place Here Start-->
		<td nowrap>
			<div id="BDIV2">
			   	<% If Num2>0 Then%>
			   		<!--Work List Start-->	
					<table border="1" cellpadding="0" cellspacing="0" width=100% id="TBInOut">
						<thead>
						   <!--HEADER INFORMATION START-->
							<tr>
								<th id="H2Col01" class="hlist" align="center" nowrap>選択</th>
								<th id="H2Col02" class="hlist" nowrap>氏名</th>
								<th id="H2Col03" class="hlist" nowrap>ドライバID</th>
								<th id="H2Col04" class="hlist" nowrap>パスワード</th>
								<th id="H2Col05" class="hlist" nowrap>会社名</th>
								<th id="H2Col06" class="hlist" nowrap>携帯番号</th>
								<th id="H2Col07" class="hlist" nowrap>メールアドレス</th>
								<!-- 2013-06-27 Y.TAKAKUWA Del-S -->
								<!--<th id="H2Col08" class="hlist" nowrap>通知結果</th>-->
								<!-- 2013-06-27 Y.TAKAKUWA Del-E -->
							</tr>
						    <!--HEADER INFORMATION END-->
						</thead>
						<tbody>
						    <!--DETAIL INFORMATION START-->
                            <% 
								x = 1
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
									x = x + 1
							%>
							<tr bgcolor="#CCFFFF">	
							  <% v_ItemName = "chkInOut" + cstr(i) %>
							<td id="D2Col01" align="center" width="30" align="center" nowrap>
							  <input type="checkbox" name="<%= v_ItemName %>" id="checkbox<%=cstr(i)%>"><BR>
							</td>
							<td id="D2Col02" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("LoDriverName"))%><BR />
                            </td>
							<td id="D2Col03" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("LoDriverID"))%><BR />
                            </td>
							<td id="D2Col04" align="center" valign="middle" nowrap>
                              <%=MID(Trim(ObjRS2("LoDriverPW")),1,1) & String(Len(Trim(ObjRS2("LoDriverPW")))-1,"*")%>
                              <BR />
                            </td>
                            
							<td id="D2Col05" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("LoDriverCompany"))%><BR />
                            </td>
                            
                            <td id="D2Col06" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("PhoneNum"))%><BR />
                            </td>
                            
							<td id="D2Col07" align="center" valign="middle" nowrap>
                              <a href="mailto:<%=Trim(ObjRS2("MailAddress"))%>"><%=Trim(ObjRS2("MailAddress"))%></a>
                              <BR />
                            </td>
                            <!--2013-06-27 Y.TAKAKUWA Del-S-->
                            <!--
                            <%If Trim(ObjRS2("AcceptStatus")) = "2" Then%>
                              <td id="D2Col08"  align="center" valign="middle" nowrap style="color:Red">失敗<BR /></td>
                            <%Else %>
                              <td id="D2Col08" align="center" valign="middle" nowrap><BR /></td>
                            <%End If %>
                            -->
                            <!--2013-06-27 Y.TAKAKUWA Del-E-->
                            <% v_ItemName = "LODriverID" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" id="InputDriverID<%=CStr(i)%>" value="<%=Trim(ObjRS2("LODriverID"))%>">
							
							<% v_ItemName = "MailAddress" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" id="InputMailAddress<%=CStr(i)%>" value="<%=Trim(ObjRS2("MailAddress"))%>">
							
							</tr>
						    <% 
									ObjRS2.MoveNext 		
									End If
								Next	
							  ObjRS2.close    
						      DisConnDBH ObjConn2, ObjRS2
						    %>  
						    <!--DETAIL INFORMATION END-->	    									
						</tbody>								
					</table>
					<!--Work List End-->
					<INPUT type=hidden name="DataCnt2" value="<%=x%>">
				<% Else %>
				    
					<table border="1" cellPadding="2" cellSpacing="0" id="NODATA">						
					  <TR class=bgw><TD nowrap>ドライバーの登録がありません</TD></TR>
					</table>
					
				<% End If %>		
			</div>
		</td>
		<!--Place Here End-->
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr>
	  <td>
	    <div>※承認メールを送信しても、携帯端末側の迷惑メール設定等で受信できない場合があります。</div>
        <div style="margin-left:12px; ">承認メールが届かない場合は、「承認ドライバ一覧」表示から再送してください。</div>
        <div style="margin-left:12px; margin-top:20px;">承認メールを運行管理者より指定ドライバへ転送することで、指定URLよりアプリをダウンロードすることが可能となります。</div>
	  </td>
	</tr>
    <tr><td>&nbsp;</td></tr>
	<tr>		
		<td>
		    <div>
			  <table border="0" cellpadding="2" cellspacing="0">
			  <tr>
			    <%
			      If v_SearchType="S2" Then
			         v_ItemValue = "承認メールを再送"
			      Else
			         v_ItemValue = "選択したドライバを承認"
			      End If
			    %>
			    <%If Num2>0 then%>
			    <td><input type="button" name="Approval" onclick="fApproval();" value="<%=v_ItemValue%>"></td>
			    <%else%>
			    <td><input type="button" name="Approval" onclick="fApproval();" value="<%=v_ItemValue%>" disabled></td>
			    <%end if%>
			  </tr>
			  <%If v_SearchType <> "S2" then%>
			  <tr><td><br /></td></tr>
			  <tr>
			    <%If Num2>0 then%>
			    <td><input type="button" name="Delete" onclick="fDelete();" value="選択したドライバを承認せず削除"></td>
			    <%else%>
			    <td><input type="button" name="Delete" onclick="fDelete();" value="選択したドライバを承認せず削除" disabled></td>
			    <%end if%>
			  </tr>
			  <%End If%>
			  </table>
			</div>		
		</td>
	</tr>    
  </table>
  </div>
  <!--Added End-->
  </div>
</td></tr>  
</table>
</form>
  
<div id="footer">
<MAP name=map>
<AREA coords=22,0,0,22,105,22,105,0 href="http://www.hits-h.com/index.asp" target="_parent" shape=POLY>
</MAP>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TR>
    <TD align=right vAlign=bottom>
      <IMG border=0 height=23 src="Image/b-home.gif" useMap="#map" width=105></TD></TR>
  <TR><TD colspan=2 bgColor=#000099 height=10></TD></TR>
</TABLE>
</div>  
</BODY>
</HTML>
