<%
'**********************************************
'  【プログラムＩＤ】　: dml000D
'  【プログラム名称】　: グループ登録
'
'  （変更履歴）
'   2013-03-18   Y.TAKAKUWA   作成
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
	dim v_DataCnt2
	dim v_DataCnt
	
	dim ErrMsg

    dim Num	
	dim strOrder
	dim FieldName	
	dim ObjRS,ObjConn
			
	dim Num2	
	dim strOrder2
	dim FieldName2	
	dim ObjRS2,ObjConn2
	
	dim wk
	dim i,x,j
	dim v_ItemName
	dim abspage, pagecnt,reccnt	
	
	dim Arr_DriverID()
	dim Arr_Check1_0()
	dim Arr_Check1_1()
	dim Arr_Check1_2()
	dim Arr_Check1_3()
	dim Arr_Check1_4()
	dim Arr_Check1_5()
	dim Arr_Check1_6()
	dim Arr_Check1_7()
	dim Arr_Check1_8()
	dim Arr_Check1_9()
	dim Arr_HiTSUserID()
	
	dim Arr_DriverID2()
	dim Arr_Check2_0()
    dim Arr_Check2_1()
	dim Arr_Check2_2()
	dim Arr_Check2_3()
	dim Arr_Check2_4()
	dim Arr_Check2_5()
	dim Arr_Check2_6()
	dim Arr_Check2_7()
	dim Arr_Check2_8()
	dim Arr_Check2_9()
	dim Arr_HiTSUserID2()
	
	dim Group0_cnt
	dim Group1_cnt
	dim Group2_cnt
	dim Group3_cnt
	dim Group4_cnt
	dim Group5_cnt
	dim Group6_cnt
	dim Group7_cnt
	dim Group8_cnt
	dim Group9_cnt
	
	Group0_cnt = 0
	Group1_cnt = 0
	Group2_cnt = 0
	Group3_cnt = 0
	Group4_cnt = 0
	Group5_cnt = 0
	Group6_cnt = 0
	Group7_cnt = 0
	Group8_cnt = 0
	Group9_cnt = 0
	
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	'Search Condition Start
	dim SDriverName
	dim SDriverCompany
	dim SDriverID
	
	dim SDriverID2
	'Search Condition End
	
	'Option Condition Start
	dim v_LogOnUser
	'Option Condition End
	
	dim v_GroupID         'Group ID
	dim v_GroupID2        'Group ID
	dim v_GroupName       'Input Group Name
	dim v_GroupName2      'Display Group Name    
	dim v_GroupNameChgFlag
	dim v_OwnGroup 
	dim Arr_OwnGroup()
	
	v_GroupNameChgFlag = ""
		
	const gcPage = 10

	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' 再描画前の項目取得
   	'----------------------------------------			
	call LfGetRequestItem
		
    If Trim(v_GroupName) <> Trim(v_GroupName2) then
      v_GroupNameChgFlag = "1"
    end if
		
	'登録
	if v_GamenMode = "I" then		
		'call LfUpdLOInfo()
	end if

    if v_GamenMode = "GI" then
      call LfSetGroupName()
    end if

	'Delete Driver
	If v_GamenMode = "D" then
	  call LfDeleteLoDriverInfo()
	end if
	
	'Delete Group Driver
	If v_GamenMode = "DG" then
	  call LfDeleteLoGroupDriverInfo()
	end if
	
	'Register Group Driver Start
	If v_GamenMode = "R" Or v_GamenMode = "RO" then
	  call LfRegisterLoGroupeDriver()
	end if
	'Register Group Driver End
	
	'Get Driver Information Start
	Call getDriverInfo()
	Call getDriverInfo2()
	'If v_GamenMode = "SD" then   
	'End If
	'Get Driver Information End
	
	'Get Group Data Start
	Call LfGetGroupData()
	Call CheckOwnGroup()
	'Get Group Data End
	
	If v_GamenMode = "SGN" then
	  Call LfGetGroupName()
	  Response.Redirect "./dml000C.asp?GamenMode=SGN2" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&LogOnUser2=" & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#groupName"
	end if
	
	If v_GamenMode = "SD" then
	  Response.Redirect "./dml000C.asp?GamenMode=SD2" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&LogOnUser2=" & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#table2"
	end if
	
Function LfGetRequestItem()

	If Request.form("Gamen_Mode") = "" then
	  v_GamenMode = Request.QueryString("GamenMode")
	else
	  v_GamenMode = Request.form("Gamen_Mode")
	end if
	if Trim(v_GamenMode) = "PS" OR Trim(v_GamenMode) = "SGN2" OR Trim(v_GamenMode) = "SD2" then
	  SDriverName = Request.QueryString("SDriverName")
	  SDriverCompany = Request.QueryString("SDriverCompany")
	  SDriverID = Request.QueryString("SDriverID")
	  SDriverID2 = Request.QueryString("SDriverID2")            'Search Input Driver ID
	  v_LogOnUser = Request.QueryString("LogOnUser")
	  v_GroupName = Request.QueryString("GroupName")            'Input Group Name
	  v_GroupName2 = Request.QueryString("GroupName2")          'Display Group Name
	  v_GroupID = Request.QueryString("GroupID")                'GroupID
	  v_GroupID2 = Request.QueryString("GroupID2")              'GroupID
	  v_DataCnt2 = Request.QueryString("DataCnt2")
 	  v_DataCnt = Request.QueryString("DataCnt")
	else	
	  v_GroupName = Request.form("SGroupName")          
	  If v_GamenMode = "SGN" Or v_GamenMode = "GI" Then
	    v_GroupID = Request.Form("cmbGroup")             
	  Else
	    v_GroupID = Request.Form("cmbGroup2")            
	  End If
	  v_GroupID2 = Request.Form("cmbGroup2")            
	  if v_GroupID = "" then
	    v_GroupID = "G0"                                 
	  end if
	  if v_GroupID2 = "" then
	    v_GroupID2 = "G0"                                
	  end if
	  v_DataCnt2 = Request.form("DataCnt2")
 	  v_DataCnt = Request.form("DataCnt")
	  if v_GamenMode = "S" Or v_GamenMode = "" then
	    SDriverName = ""
	    SDriverCompany = ""
	    SDriverID = ""
	    SDriverID2 = ""                                    
	    v_DriverInfo = Request.form("driverInfo")
	    v_LogOnUser = "3"
	  else
	    SDriverName = Request.form("SDriverName")
	    SDriverCompany = Request.form("SDriverCompany")
	    SDriverID = Request.form("SDriverID")
	    SDriverID2 = Request.form("SDriverID2")            
	    v_DriverInfo = Request.form("driverInfo")
	    v_LogOnUser = Request.form("selectCompany")
	    if v_LogOnUser <> "3" then
	      v_GroupID2 = "G0"
	    end if
	  end if
	end if

	if Trim(v_DataCnt) = "" then
	  v_DataCnt = 0
	end if
	if Trim(v_DataCnt2) = "" then
	  v_DataCnt2 = 0
	end if
	
	ReDimension(v_DataCnt)
    ReDimension2(v_DataCnt2)
    
	For i = 1 to (v_DataCnt) - 1 
	  Arr_Check1_0(i) = Trim(Request.form("chkInOutG0_" & i))
	  Arr_Check1_1(i) = Trim(Request.form("chkInOutG1_" & i))
	  Arr_Check1_2(i) = Trim(Request.form("chkInOutG2_" & i))
	  Arr_Check1_3(i) = Trim(Request.form("chkInOutG3_" & i))
	  Arr_Check1_4(i) = Trim(Request.form("chkInOutG4_" & i))
	  Arr_Check1_5(i) = Trim(Request.form("chkInOutG5_" & i))
	  Arr_Check1_6(i) = Trim(Request.form("chkInOutG6_" & i))
	  Arr_Check1_7(i) = Trim(Request.form("chkInOutG7_" & i))
	  Arr_Check1_8(i) = Trim(Request.form("chkInOutG8_" & i))
	  Arr_Check1_9(i) = Trim(Request.form("chkInOutG9_" & i))
      Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
      Arr_HiTSUserID(i) = TRIM(Request.form("HiTSUserID" & i))
	Next
	
	For i = 1 to (v_DataCnt2) - 1 
      Arr_Check2_0(i) = Trim(Request.form("chkInOut2G0_" & i))
      Arr_Check2_1(i) = Trim(Request.form("chkInOut2G1_" & i))
      Arr_Check2_2(i) = Trim(Request.form("chkInOut2G2_" & i))
      Arr_Check2_3(i) = Trim(Request.form("chkInOut2G3_" & i))
      Arr_Check2_4(i) = Trim(Request.form("chkInOut2G4_" & i))
      Arr_Check2_5(i) = Trim(Request.form("chkInOut2G5_" & i))
      Arr_Check2_6(i) = Trim(Request.form("chkInOut2G6_" & i))
      Arr_Check2_7(i) = Trim(Request.form("chkInOut2G7_" & i))
      Arr_Check2_8(i) = Trim(Request.form("chkInOut2G8_" & i))
      Arr_Check2_9(i) = Trim(Request.form("chkInOut2G9_" & i))
      Arr_DriverID2(i) = TRIM(Request.form("LODriverID2_" & i))
      Arr_HiTSUserID2(i) = TRIM(Request.form("HiTSUserID2_" & i))
	Next
	
End Function

Function ReDimension(index)
   Redim Arr_Check1_0(index)
   Redim Arr_Check1_1(index)
   Redim Arr_Check1_2(index)
   Redim Arr_Check1_3(index)
   Redim Arr_Check1_4(index)
   Redim Arr_Check1_5(index)
   Redim Arr_Check1_6(index)
   Redim Arr_Check1_7(index)
   Redim Arr_Check1_8(index)
   Redim Arr_Check1_9(index)
   Redim Arr_DriverID(index)
   Redim Arr_HiTSUserID(index)
End Function

Function ReDimension2(index)
   Redim Arr_Check2_0(index)
   Redim Arr_Check2_1(index)
   Redim Arr_Check2_2(index)
   Redim Arr_Check2_3(index)
   Redim Arr_Check2_4(index)
   Redim Arr_Check2_5(index)
   Redim Arr_Check2_6(index)
   Redim Arr_Check2_7(index)
   Redim Arr_Check2_8(index)
   Redim Arr_Check2_9(index)
   Redim Arr_DriverID2(index)
   Redim Arr_HiTSUserID2(index)
End Function

Function getDriverInfo()
    dim StrSQL
    dim i
    ConnDBH ObjConn, ObjRS
    
    'StrSQL = "SELECT "
    'StrSQL = StrSQL & " LomDriver.*, ISNULL(LoGroupeDriver.LoDriverID,'') As GroupDriverID, ISNULL(LoGroupeDriver.LoGroupID,'') As GroupID "
    'StrSQL = StrSQL & " FROM LomDriver "
    'StrSQL = StrSQL & " INNER JOIN LoGroupeDriver ON LoGroupeDriver.LoDriverID = LomDriver.LoDriverID AND LomDriver.HiTSUserID = LoGroupeDriver.HiTSUserID "
    'If Trim(v_LogOnUser) = "3" Then
    '  StrSQL = StrSQL & " AND LoGroupeDriver.LoGroupID='" & Trim(v_GroupID) & "'"
    'End If
    
    StrSQL = " SELECT DISTINCT "
    StrSQL = StrSQL & " LomDriver.LoDriverName, LomDriver.LoDriverID, LomDriver.LoDriverCompany, LomDriver.HiTSUserID  "
    StrSQL = StrSQL & " , ISNULL(G0Group.LoGroupID,'') AS Group0 "
    StrSQL = StrSQL & " , ISNULL(G1Group.LoGroupID,'') AS Group1 "
    StrSQL = StrSQL & " , ISNULL(G2Group.LoGroupID,'') AS Group2 "
    StrSQL = StrSQL & " , ISNULL(G3Group.LoGroupID,'') AS Group3 "
    StrSQL = StrSQL & " , ISNULL(G4Group.LoGroupID,'') AS Group4 "
    StrSQL = StrSQL & " , ISNULL(G5Group.LoGroupID,'') AS Group5 "
    StrSQL = StrSQL & " , ISNULL(G6Group.LoGroupID,'') AS Group6 "
    StrSQL = StrSQL & " , ISNULL(G7Group.LoGroupID,'') AS Group7 "
    StrSQL = StrSQL & " , ISNULL(G8Group.LoGroupID,'') AS Group8 "
    StrSQL = StrSQL & " , ISNULL(G9Group.LoGroupID,'') AS Group9 " 
    StrSQL = StrSQL & " FROM LomDriver " 
    
    For i = 0 to 9
      If CInt(MID(Trim(v_GroupID2),2,1)) = i Then
         If Trim(v_LogOnUser) = "3" Then
           StrSQL = StrSQL & " INNER JOIN LoGroupeDriver AS G" & CStr(i) & "Group ON " 
         Else
           StrSQL = StrSQL & " LEFT JOIN LoGroupeDriver AS G" & CStr(i) & "Group ON " 
         End If
           StrSQL = StrSQL & " LomDriver.LoDriverID = G" & CStr(i) & "Group.LoDriverID AND "
           'StrSQL = StrSQL & " LomDriver.HiTSUserID = G" & CStr(i) & "Group.HiTSUserID AND " 
           StrSQL = StrSQL & " G" & i & "Group.LoGroupID='" & Trim(v_GroupID2) & "'"
           StrSQL = StrSQL & " AND G" & i & "Group.HiTSUserID='" & USER & "'"
      Else
          StrSQL = StrSQL & " LEFT JOIN LoGroupeDriver AS G" & CStr(i) & "Group ON " 
          StrSQL = StrSQL & " LomDriver.LoDriverID = G" & CStr(i) & "Group.LoDriverID AND "
          'StrSQL = StrSQL & " LomDriver.HiTSUserID = G" & CStr(i) & "Group.HiTSUserID AND " 
          StrSQL = StrSQL & " G" & i & "Group.LoGroupID='G" & Trim(CStr(i)) & "'" 
          StrSQL = StrSQL & " AND G" & i & "Group.HiTSUserID='" & USER & "'"
      End if
    Next
    If Trim(v_LogOnUser) = "2" Or Trim(v_LogOnUser) = "3" Then
       If Trim(v_LogOnUser) = "2" Then
         StrSQL = StrSQL & " RIGHT JOIN "
       ElseIf Trim(v_LogOnUser) = "3" Then
         StrSQL = StrSQL & " INNER JOIN "
       End If
       StrSQL = StrSQL & " ( "
       StrSQL = StrSQL & " SELECT DISTINCT LoGroupeDriver.LoGroupID FROM LomGroup "
       StrSQL = StrSQL & " INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "
       'StrSQL = StrSQL & " INNER JOIN LoGroupeDriver ON LomDriver.LoDriverID = LoGroupeDriver.LoDriverID "
       'StrSQL = StrSQL & " INNER JOIN LomGroup ON LomDriver.HiTSUserID = LomGroup.HiTSUserID "
       StrSQL = StrSQL & " WHERE LomGroup.HiTSUserID = '" & USER & "' "
       If Trim(v_LogOnUser) = "3" Then
         StrSQL = StrSQL + " AND LomGroup.LoGroupID ='" & Trim(v_GroupID2)  & "'"
       End If
       StrSQL = StrSQL & " ) OWNGROUP ON "
       StrSQL = StrSQL & " G0Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G1Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G2Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G3Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G4Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G5Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G6Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G7Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G8Group.LoGroupID = OWNGROUP.LoGroupID OR "
       StrSQL = StrSQL & " G9Group.LoGroupID = OWNGROUP.LoGroupID "
    End If
    
    
    
    if Trim(SDriverName) <> "" or Trim(SDriverCompany) <> "" or Trim(SDriverID) <> "" then
       StrSQL = StrSQL  & " WHERE "
       if Trim(SDriverName) <> "" then
         StrSQL = StrSQL  & "LomDriver.LoDriverName LIKE '%" & Trim(SDriverName) & "%'"
       end if
       
       if Trim(SDriverCompany) <> "" then
         if Trim(SDriverName) <> "" then
            StrSQL = StrSQL  & " AND "  
         end if
         StrSQL = StrSQL  & "LomDriver.LoDriverCompany LIKE '%" & Trim(SDriverCompany) & "%'"
       end if
       if Trim(SDriverID) <> "" then
         if Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" then
            StrSQL = StrSQL  & " AND "  
         end if
         StrSQL = StrSQL  & "LomDriver.LoDriverID LIKE '%" & Trim(SDriverID) & "%'"
       end if
    end if
    
    if Trim(SDriverName) = "" and Trim(SDriverCompany) = "" and Trim(SDriverID) = "" and Trim(v_LogOnUser) <> "3" then
      StrSQL = StrSQL  & " WHERE "
    end if
    
    if Trim(v_LogOnUser) = "1" Or Trim(v_LogOnUser) = "" then
      if Trim(SDriverID) <> "" Or Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" then
        StrSQL = StrSQL  & " AND "  
      end if
      StrSQL = StrSQL  & " LomDriver.HiTSUserID='" & USER & "' "
      StrSQL = StrSQL  & " AND LomDriver.AcceptStatus='1' "
    elseIf Trim(v_LogOnUser) = "2" then
      if Trim(SDriverID) <> "" Or Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" then
        StrSQL = StrSQL  & " AND "  
      end if
      StrSQL = StrSQL  & " LomDriver.HiTSUserID<>'" & USER & "' "
      StrSQL = StrSQL  & " AND LomDriver.AcceptStatus='1' "
    end if
    StrSQL = StrSQL & " ORDER BY LomDriver.LoDriverID "
    'response.Write StrSQL
    ObjRS.PageSize = 100
	ObjRS.CacheSize = 100
	ObjRS.CursorLocation = 3
	ObjRS.Open StrSQL, ObjConn
	
	Num = ObjRS.recordcount	
	
	if Num > 100 then
		If CInt(Request("pagenum")) = 0 Then
			ObjRS.AbsolutePage = 1
		Else
			If CInt(Request("pagenum")) <= ObjRS.PageCount Then
				ObjRS.AbsolutePage = CInt(Request("pagenum"))
			Else
				ObjRS.AbsolutePage = 1
			End If
		End If		 
	end if
	
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "2","b301","01","ロックオン事前情報","102","SQL:<BR>" & StrSQL & err.description & Err.number
		Exit Function
	end if			
	'エラートラップ解除
    on error goto 0	

End Function

Function LfRegisterLoGroupeDriver()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	dim strDriverID
	dim strDriverGroupID
	dim strGroupID
	dim strHiTSUserID
	dim dataCnt

    ConnDBH ObjConnLO, ObjRSLO
    '2013-09-26 Y.TAKAKUWA Add-S	
	WriteLogH "b504", "グループ登録（削除）実行", "01", ""
	'2013-09-26 Y.TAKAKUWA Add-E
	ErrFlg = false
	
	If Trim(v_GamenMode) = "R" Then
	  dataCnt = v_DataCnt2
	Else
	  dataCnt = v_DataCnt
	End If
	
    For i = 1 to dataCnt-1
      For j = 0 to 9  
        If Trim(v_GamenMode) = "R" Then
          strDriverID = Trim(Arr_DriverID2(i))
          strHiTSUserID = Trim(Arr_HiTSUserID2(i))
        Else
          strDriverID = Trim(Arr_DriverID(i))
          strHiTSUserID = Trim(Arr_HiTSUserID(i))
        End If
        
        Select Case j
          Case 0
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_0(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_0(i))
            End If
            strGroupID = "G0"
          Case 1
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_1(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_1(i))
            End If
            strGroupID = "G1"
          Case 2
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_2(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_2(i))
            End If
            strGroupID = "G2"
          Case 3
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_3(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_3(i))
            End If
            strGroupID = "G3"
          Case 4
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_4(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_4(i))
            End If
            strGroupID = "G4"
          Case 5
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_5(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_5(i))
            End If
            strGroupID = "G5"
          Case 6
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_6(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_6(i))
            End If
            strGroupID = "G6"
          Case 7
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_7(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_7(i))
            End If
            strGroupID = "G7"
          Case 8
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_8(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_8(i))
            End If
            strGroupID = "G8"
          Case 9
            If Trim(v_GamenMode) = "R" Then
              strDriverGroupID = Trim(Arr_Check2_9(i))
            Else
              strDriverGroupID = Trim(Arr_Check1_9(i))
            End If
            strGroupID = "G9"
        End Select
        
        If strDriverID <> "" and strDriverGroupID = "on" then
          'QUERY IF GROUP EXIST
          StrSQL = "SELECT * FROM LomGroup WHERE HiTSUserID ='" & USER  & "'"&_
                                       " AND LoGroupID='" & Trim(strGroupID) & "'"  
          ObjRSLO.Open StrSQL, ObjConnLO
          'response.Write StrSQL
          If ObjRSLO.recordcount = 0 then
             StrSQL = " INSERT INTO LomGroup (HiTSUserID, LoGroupID, UpdtTime, UpdtPgCd, UpdtTmnl)"
             StrSQL = StrSQL & " VALUES ( "
             StrSQL = StrSQL & "'" & USER & "',"                            'HiTSUserID
             StrSQL = StrSQL & "'" & Trim(strGroupID) & "',"                 'LoGroupID
             StrSQL = StrSQL & "'" & Now() & "',"                           'UpdtTime
             StrSQL = StrSQL & "'" & "PREDEF01" & "',"                      'UpdtPgCd
             StrSQL = StrSQL & "'" & USER & "' "                            'UpdtTmnl
             StrSQL = StrSQL & ")"
             ObjConnLO.Execute(StrSQL)  
             if err <> 0 then
                Set ObjRSLO = Nothing				
                jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
             end if
          End If
          ObjRSLO.Close
          
          'QUERY VALUES BEFORE INSERTION
          StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID ='" & USER & "'"&_
                                                 " AND LoGroupID='" & Trim(strGroupID) & "'"&_
                                                 " AND LoDriverID='" & Trim(strDriverID) & "'"
                                                
          ObjRSLO.Open StrSQL, ObjConnLO
          If ObjRSLO.recordcount > 0 then
            'UPDATE
            StrSQL = " UPDATE LoGroupeDriver SET "
            StrSQL = StrSQL & "HiTSUserID='" & USER & "'," 'Trim(strHiTSUserID) & "',"                'HiTSUserID
            StrSQL = StrSQL & "LoGroupID='" & Trim(strGroupID) & "', "                         'LoGroupID
            StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                                     'UpdtTime
            StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"                                'UpdtPgCd
            StrSQL = StrSQL & "UpdtTmnl='" & USER & "' "                  'UpdtTmnl
            StrSQL = StrSQL & "WHERE HiTSUserID ='" & USER & "'"&_
                               " AND LoGroupID='" & Trim(strGroupID) & "'"&_
                               " AND LoDriverID='" & Trim(strDriverID) & "'"
            ObjConnLO.Execute(StrSQL)

            if err <> 0 then
		      Set ObjRSLO = Nothing				
		      jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	        end if
	        ObjRSLO.Close
          else
            'INSERT
            StrSQL = " INSERT INTO LoGroupeDriver (HiTSUserID, LoGroupID, UpdtTime, UpdtPgCd, UpdtTmnl, LoDriverID )"
            StrSQL = StrSQL & " VALUES ( "
            StrSQL = StrSQL & "'" & USER & "'," 'Trim(strHiTSUserID) & "',"        'HiTSUserID
            StrSQL = StrSQL & "'" & Trim(strGroupID) & "',"                 'LoGroupID
            StrSQL = StrSQL & "'" & Now() & "',"                           'UpdtTime
            StrSQL = StrSQL & "'" & "PREDEF01" & "',"                      'UpdtPgCd
            StrSQL = StrSQL & "'" & USER & "',"        'UpdtTmnl
            StrSQL = StrSQL & "'" & Trim(strDriverID) & "' "          'LoDriverID
            StrSQL = StrSQL & ")"
            ObjConnLO.Execute(StrSQL)
          
            if err <> 0 then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		    end if
		    ObjRSLO.Close          
          end if
        Else
          'QUERY VALUES FOR Delete
          StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID ='" & USER & "'"&_
                                                 " AND LoGroupID ='" & Trim(strGroupID) & "'"&_
                                                 " AND LoDriverID ='" & Trim(strDriverID) & "'"                                               
          ObjRSLO.Open StrSQL, ObjConnLO
          If ObjRSLO.recordcount > 0 then
             StrSQL = " DELETE FROM LoGroupeDriver WHERE "
             StrSQL = StrSQL & "      HiTSUserID ='" & USER  & "'"&_
                                " AND LoGroupID ='" & Trim(strGroupID) & "'"&_
                                " AND LoDriverID ='" & Trim(strDriverID) & "'" 
             ObjConnLO.Execute(StrSQL)
             if err <> 0 then
		       Set ObjRSLO = Nothing				
		     jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		     end if
	       end if
	       ObjRSLO.Close
	    end if
	  Next
    Next
    
    DisConnDBH ObjConnLO, ObjRSLO	'DB切断

End Function

Function getDriverInfo2()
    dim StrSQL
 
    ConnDBH ObjConn2, ObjRS2
    
    'StrSQL = " SELECT "
    'StrSQL = StrSQL & " LomDriver.*, ISNULL(LoGroupeDriver.LoDriverID,'') As GroupDriverID "
    'StrSQL = StrSQL & " FROM LomDriver "
    'StrSQL = StrSQL & " LEFT JOIN LoGroupeDriver ON LoGroupeDriver.LoDriverID = LomDriver.LoDriverID AND LomDriver.HiTSUserID = LoGroupeDriver.HiTSUserID AND LoGroupeDriver.LoGroupID='" & Trim(v_GroupID) & "'"
    
    StrSQL = " SELECT "
    StrSQL = StrSQL & " LomDriver.LoDriverName, LomDriver.LoDriverID, LomDriver.LoDriverCompany, LomDriver.HiTSUserID, LomDriver.PhoneNum, LomDriver.MailAddress "
    StrSQL = StrSQL & " , ISNULL(G0Group.LoGroupID,'') AS Group0 "
    StrSQL = StrSQL & " , ISNULL(G1Group.LoGroupID,'') AS Group1 "
    StrSQL = StrSQL & " , ISNULL(G2Group.LoGroupID,'') AS Group2 "
    StrSQL = StrSQL & " , ISNULL(G3Group.LoGroupID,'') AS Group3 "
    StrSQL = StrSQL & " , ISNULL(G4Group.LoGroupID,'') AS Group4 "
    StrSQL = StrSQL & " , ISNULL(G5Group.LoGroupID,'') AS Group5 "
    StrSQL = StrSQL & " , ISNULL(G6Group.LoGroupID,'') AS Group6 "
    StrSQL = StrSQL & " , ISNULL(G7Group.LoGroupID,'') AS Group7 "
    StrSQL = StrSQL & " , ISNULL(G8Group.LoGroupID,'') AS Group8 "
    StrSQL = StrSQL & " , ISNULL(G9Group.LoGroupID,'') AS Group9 " 
    StrSQL = StrSQL & " FROM LomDriver " 
    
    For i = 0 to 9
          StrSQL = StrSQL & " LEFT JOIN LoGroupeDriver AS G" & i & "Group ON " 
          StrSQL = StrSQL & " LomDriver.LoDriverID = G" & i & "Group.LoDriverID AND "
          'StrSQL = StrSQL & " LomDriver.HiTSUserID = G" & i & "Group.HiTSUserID AND " 
          StrSQL = StrSQL & " G" & i & "Group.LoGroupID='G" & Trim(i) & "'"
          StrSQL = StrSQL & " AND G" & i & "Group.HiTSUserID='" & USER & "'"
    Next
    StrSQL = StrSQL  & " WHERE LomDriver.HiTSUserID<>'" & USER & "'"
    
    'if Trim(SDriverID2) <> "" then
       StrSQL = StrSQL  & " AND LomDriver.LoDriverID = '" & Trim(SDriverID2) & "'"
       StrSQL = StrSQL  & " AND LomDriver.AcceptStatus='1' "
    'end if
    StrSQL = StrSQL & " ORDER BY LomDriver.LoDriverID "
    'Response.Write StrSQL
    ObjRS2.PageSize = 100
	ObjRS2.CacheSize = 100
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2
	
	Num2 = ObjRS2.recordcount	
	
	if Num2 > 100 then
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

Function LfDeleteLoDriverInfo()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
	
	For i = 1 to v_DataCnt-1
      If UCase(Trim(Arr_Check1_0(i))) = "ON" Then
        'QUERY VALUES FOR Delete
        StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID ='" & USER  & "'"&_
                                               " AND LoGroupID ='" & Trim(v_GroupID2) & "'"&_
                                               " AND LoDriverID ='" & Trim(Arr_DriverID(i)) & "'" 
                                                        
        ObjRSLO.Open StrSQL, ObjConnLO
        If ObjRSLO.recordcount > 0 then
            StrSQL = " DELETE FROM LoGroupeDriver WHERE "
            StrSQL = StrSQL & "      HiTSUserID ='" & USER  & "'"&_
                               " AND LoGroupID ='" & Trim(v_GroupID2) & "'"&_
                               " AND LoDriverID ='" & Trim(Arr_DriverID(i)) & "'" 
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


Function LfDeleteLoGroupDriverInfo
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
	
	
	For i = 1 to v_DataCnt-1
        'QUERY VALUES FOR Delete
        StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID ='" & USER  & "'"&_
                                               " AND LoGroupID ='" & Trim(v_GroupID2) & "'"
                                                        
        ObjRSLO.Open StrSQL, ObjConnLO
        If ObjRSLO.recordcount > 0 then
            StrSQL = " DELETE FROM LoGroupeDriver WHERE "
            StrSQL = StrSQL & "      HiTSUserID ='" & USER  & "'"&_
                               " AND LoGroupID ='" & Trim(v_GroupID2) & "'"
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		    end if
	    end if
	    ObjRSLO.Close
    Next
    
    'Delete in Master Group
    StrSQL = "SELECT * FROM LomGroup WHERE HiTSUserID ='" & USER  & "'"&_
                                               " AND LoGroupID ='" & Trim(v_GroupID2) & "'"                                          
    ObjRSLO.Open StrSQL, ObjConnLO
    If ObjRSLO.recordcount > 0 then
      StrSQL = " DELETE FROM LomGroup WHERE "
      StrSQL = StrSQL & "      HiTSUserID ='" & USER & "'"&_
                         " AND LoGroupID ='" & Trim(v_GroupID2) & "'"
      ObjConnLO.Execute(StrSQL)
      if err <> 0 then
		 Set ObjRSLO = Nothing				
		 jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	  end if
	end if
	ObjRSLO.Close
	    
    DisConnDBH ObjConnLO, ObjRSLO	'DB切断

End Function

Function LfGetGroupData
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""

      'QUERY VALUES FOR INSERTION
      StrSQL = "SELECT * FROM LomGroup WHERE HiTSUserID ='" & USER  & "'"&_
                                       " AND LoGroupID='" & Trim(v_GroupID2) & "'"
      ObjRSLO.Open StrSQL, ObjConnLO

      If ObjRSLO.recordcount > 0 then
         v_GroupName = Trim(ObjRSLO("LoDriverCompany"))
         v_GroupName2 = Trim(ObjRSLO("LoDriverCompany")) 
      Else
         v_GroupName = ""
         v_GroupName2 = ""
      End if
      ObjRSLO.Close
      
      If (v_GroupID2 <> v_GroupID) Then
        StrSQL = "SELECT * FROM LomGroup WHERE HiTSUserID ='" & USER  & "'"&_
                                         " AND LoGroupID='" & Trim(v_GroupID) & "'"
        ObjRSLO.Open StrSQL, ObjConnLO

        If ObjRSLO.recordcount > 0 then
          v_GroupName = Trim(ObjRSLO("LoDriverCompany"))
        Else
          v_GroupName = ""
        End if
        ObjRSLO.Close
      End If

    DisConnDBH ObjConnLO, ObjRSLO	'DB切断
    
End Function

Function LfGetGroupName
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""

      'QUERY VALUES FOR INSERTION
      StrSQL = "SELECT * FROM LomGroup WHERE HiTSUserID ='" & USER  & "'"&_
                                       " AND LoGroupID='" & Trim(v_GroupID) & "'"
      ObjRSLO.Open StrSQL, ObjConnLO

      If ObjRSLO.recordcount > 0 then
         v_GroupName = Trim(ObjRSLO("LoDriverCompany")) 
      Else
         v_GroupName = ""
      End if
      ObjRSLO.Close
    DisConnDBH ObjConnLO, ObjRSLO	'DB切断
    
End Function

Function LfSetGroupName
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim iSeq
	dim ByteLength
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""

    'Calculate first the Byte Width of the String
    StrSQL = "SELECT TOP 1 DATALENGTH('" & Trim(v_GroupName) & "') AS ByteLength "
    ObjRSLO.Open StrSQL, ObjConnLO
    If ObjRSLO.recordcount > 0 then
       ByteLength = Trim(ObjRSLO("ByteLength"))
    end if
    ObjRSLO.Close
    If Trim(ByteLength) <> "" then

    If Cint(Trim(ByteLength)) <= 20 then
      StrSQL = "SELECT * FROM LomGroup WHERE HiTSUserID ='" & USER  & "'"&_
                                       " AND LoGroupID='" & Trim(v_GroupID) & "'"                   
      ObjRSLO.Open StrSQL, ObjConnLO
   
     
      If ObjRSLO.recordcount > 0 then
        if Trim(v_GroupName) <> "" Then
          StrSQL = " UPDATE LomGroup SET "
          StrSQL = StrSQL & "HiTSUserID='" & USER & "', "                 'HiTSUserID
          StrSQL = StrSQL & "LoGroupID='" & Trim(v_GroupID) & "', "       'LoGroupID
          StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
          StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
          StrSQL = StrSQL & "UpdtTmnl='" & USER & "', "                   'UpdtTmnl
          StrSQL = StrSQL & "LoDriverCompany='" & v_GroupName & "' "      'LoDriverCompany
          StrSQL = StrSQL & "WHERE HiTSUserID ='" & USER  & "'"&_         
                           " AND LoGroupID='" & Trim(v_GroupID) & "'"
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
		    Set ObjRSLO = Nothing				
		    jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	      end if
	      ObjRSLO.Close
	    Else
	      StrSQL = " DELETE FROM LomGroup "
          StrSQL = StrSQL & "WHERE HiTSUserID ='" & USER  & "'"&_         
                             " AND LoGroupID='" & Trim(v_GroupID) & "'"
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
		    Set ObjRSLO = Nothing				
		    jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	      end if
	      ObjRSLO.Close
	    End If
      else
        StrSQL = " INSERT INTO LomGroup (HiTSUserID, LoGroupID, UpdtTime, UpdtPgCd, UpdtTmnl, LoDriverCompany)"
        StrSQL = StrSQL & " VALUES ( "
        StrSQL = StrSQL & "'" & USER & "',"                            'HiTSUserID
        StrSQL = StrSQL & "'" & Trim(v_GroupID) & "',"                 'LoGroupID
        StrSQL = StrSQL & "'" & Now() & "',"                           'UpdtTime
        StrSQL = StrSQL & "'" & "PREDEF01" & "',"                      'UpdtPgCd
        StrSQL = StrSQL & "'" & USER & "',"                            'UpdtTmnl
        StrSQL = StrSQL & "'" & Trim(v_GroupName) & "' "               'LoDriverCompany
        StrSQL = StrSQL & ")"
        'response.Write strSQL
        ObjConnLO.Execute(StrSQL)  
        if err <> 0 then
		  Set ObjRSLO = Nothing				
		  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		end if
		ObjRSLO.Close
      end if
    else
       ErrMsg="登録なし: 全角１０文字、半角２０文字までのみ。" 
    end if
    end if
    DisConnDBH ObjConnLO, ObjRSLO	'DB切断
    
End Function

Function CheckOwnGroup()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim icnt
        
    ConnDBH ObjConnLO, ObjRSLO
    v_OwnGroup = ""
      StrSQL = "SELECT * FROM LomGroup WHERE "
      StrSQL = StrSQL + " HiTSUserID ='" & USER  & "'"
      'response.Write strSQL
      ObjRSLO.Open StrSQL, ObjConnLO
      If ObjRSLO.recordcount > 0 then
        While Not ObjRSLO.EOF 
          v_OwnGroup = v_OwnGroup & Trim(ObjRSLO("LoGroupID")) & ","
          ObjRSLO.MoveNext
        Wend
      else
        v_OwnGroup = ""
      end if
      Redim Arr_OwnGroup(9)
      For icnt = 0 to 9 
        If InStr(v_OwnGroup,"G" & icnt) > 0 Then
          Arr_OwnGroup(icnt) = "G" & icnt  
        Else
          Arr_OwnGroup(icnt) = ""
        End If
      Next
      
      ObjRSLO.Close
    DisConnDBH ObjConnLO, ObjRSLO	'DB切断
End function

function LfPutPage(rec,page,pagecount,link,focus)
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
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """>最初へ</a>"
			response.write "| &nbsp;"
			'Y.TAKAKUWA Upd-S 2015-03-13
			'if PageWkNo<>0 Then
			'	response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & v_DataCnt2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """>前へ</a>"
			'Else
			'	response.write "<font style='color:#FFFFFF;'>前へ</font>"
			'End If
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & page-1 & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & v_DataCnt2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """>前へ</a>"
			'Y.TAKAKUWA Upd-E 2015-03-13
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
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&LogOnUser2=" & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
		
        'Y.TAKAKUWA Upd-S 2015-03-13
        'if page<pagecount-1 then		
		if page<pagecount then
		'Y.TAKAKUWA Upd-E 2015-03-13
			'Y.TAKAKUWA Upd-S 2015-03-13
			'PageWkNo=PageWkNo+1
			'If PageWkNo<=LastPage Then
			'	response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&LogOnUser2=" & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """>次へ</a>"'
			'Else
			'	response.write "<font style='color:#FFFFFF;'>次へ</font>"
			'End If
			PageWkNo=page+1
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&LogOnUser2=" & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """>次へ</a>"'
			'Y.TAKAKUWA Upd-E 2015-03-13
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&SDriverID2=" & SDriverID2 & "&LogOnUser=" & v_LogOnUser & "&LogOnUser2=" & "&DataCnt=" & v_DataCnt & "&DataCnt2=" & v_DataCnt2 & "&GroupID=" & v_GroupID & "&GroupName=" & v_GroupName & "&GroupName2=" & v_GroupName2 & "&GroupID2=" & v_GroupID2 & "#" & focus & """>最後へ</a>"'            
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
<meta http-equiv="X-UA-Compatible" content="IE=11; IE=10; IE=9; IE=8; IE=7; IE=EDGE" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.js"></script>
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


</style>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT Language="JavaScript">
// Y.TAKAKUWA Add-S 2014-07-19
function MakeStaticHeader(gridId, height, width, headerHeight, isFooter) 
{
    width = document.body.offsetWidth-280;
    var rowHeightThead = getRowHeightThead2();
	rowHeightThead = rowHeightThead
    var tbl = document.getElementById(gridId);
	var tbl2 = document.getElementById('TBInOut2');

    if (tbl) {
        var DivHR = document.getElementById('DivHeaderRow');
        var DivMC = document.getElementById('BDIV1');
        
        //*** Set divheaderRow Properties ****
        DivHR.style.height = rowHeightThead + 'px';
        DivHR.style.position = 'relative';
        DivHR.style.top = '0px';
        DivHR.style.zIndex = '10';
        DivHR.style.verticalAlign = 'top';

        //*** Set divMainContent Properties ****
        DivMC.style.width = width + 'px';
        DivMC.style.height = height + 'px';
        DivMC.style.position = 'relative';
        DivMC.style.top = -rowHeightThead + 'px';
        DivMC.style.zIndex = '1';
        //****Copy Header in divHeaderRow****
        DivHR.appendChild(tbl.cloneNode(true));
	
		if(DivMC.offsetWidth<=DivMC.clientWidth){
          DivHR.style.width = (parseInt(width)) + 'px';
		}
	else{
			DivHR.style.width = (parseInt(width)-17) + 'px';
		}
    }
}
function OnScrollDiv(Scrollablediv) {
    document.getElementById('DivHeaderRow').scrollLeft = Scrollablediv.scrollLeft;
}
// Y.TAKAKUWA Add-S 2014-07-19
function finit(){
    var i;
	//データ引継ぎ設定  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";  
    //document.frm.SGroupName2.value="<%=v_GroupName2%>";  
    document.frm.cmbGroup2.value="<%=v_GroupID2%>"; 
    //alert("<%=v_GamenMode%>")
    if(document.frm.Gamen_Mode.value == "SGN2" || document.frm.Gamen_Mode.value == "GI"){
      document.frm.cmbGroup.value="<%=v_GroupID%>";
      document.frm.SGroupName.value="<%=v_GroupName%>"
    }
    else{
      document.frm.cmbGroup.value="<%=v_GroupID2%>";
      document.frm.SGroupName.value="<%=v_GroupName2%>"
    }
    document.frm.SDriverID2.value="<%=SDriverID2%>";
    
    document.frm.SDriverName.value="<%=SDriverName%>";
    document.frm.SDriverCompany.value="<%=SDriverCompany%>";
    document.frm.SDriverID.value="<%=SDriverID%>";
    
      if("<%=v_LogOnUser %>"=="1"){
        document.getElementById("chk1").checked=true;
      }
      else{
        if("<%=v_LogOnUser %>"=="2"){
          document.getElementById("chk2").checked=true;
        }
        else{
          if("<%=v_LogOnUser %>"=="3"){
            document.getElementById("chk3").checked=true;
            //document.getElementById("<%=v_GroupID%>_AllCheck").checked=true;
          }
        }
      }   
}

//データが無い場合の表示制御
function view(){

	var sortedHeight;
	sortedHeight = 0;
	var vHeight;
	var obj1=document.getElementById("BDIV1");
	var obj2=document.getElementById("BDIV2");
        var obj4=document.getElementById("DivHeaderRow"); // Y.TAKAKUWA Add-S 2014-07-19
	var rowHeight;
	if('<%=Num2%>'!='0'){
	  var rowHeightThead = getRowHeightThead();
	  var rowHeightTbody = getRowHeightTbody3();
	  
	  if(rowHeightThead > 0){
	    rowHeightThead=rowHeightThead;
	  }
	  if(rowHeightTbody > 0){
	    if(parseInt('<%=Num2%>')>9){
	      rowHeight=rowHeightTbody*10;
	    }
	    else
	    {
	      rowHeight=rowHeightTbody*parseInt('<%=Num2%>');
	    }
	    
	  }
	  rowHeight=rowHeight+rowHeightThead;
    }
    else{
      rowHeight = 0;
      //rowHeight=23*10;
    }
    
	if((document.body.offsetWidth-50) < 230){
		obj2.style.width=30;
		obj2.style.overflowX="auto";
	}else if((document.body.offsetWidth-50)  < 813){
		obj2.style.width=document.body.offsetWidth-280;
		obj2.style.overflowX="auto";
	}else{
		obj2.style.width=document.body.offsetWidth-280;
		obj2.style.overflowX="auto";
	}	

    if(obj2.clientWidth < obj2.scrollWidth)
    {
       obj2.style.height = rowHeight+17;
       obj2.style.overflowY="auto";
    }
    else{   
       obj2.style.height = rowHeight;
       obj2.style.overflowY="auto";
    } 

    if('<%=Num%>'!='0'){
	  var rowHeightThead = getRowHeightThead2(); // Y.TAKAKUWA Upd-S 2014-07-19
	  var rowHeightTbody = getRowHeightTbody2();
	  
	  if(rowHeightThead > 0){
	    rowHeightThead=rowHeightThead;
	  }
	  if(rowHeightTbody > 0){
	    rowHeight=rowHeightTbody*10;
	  }
	  rowHeight=rowHeight+rowHeightThead;
    }
    else{
      rowHeight = 0;
      rowHeight=23*10;
    }
	if((document.body.offsetWidth-50) < 230){
		obj1.style.width=30;
		obj4.style.width=document.body.offsetWidth-30;     // Y.TAKAKUWA Add-S 2014-07-19
		obj1.style.overflowX="auto";	
                obj4.style.overflowX="hidden";		 	   // Y.TAKAKUWA Add-S 2014-07-19
	}else if((document.body.offsetWidth-50)  < 813){
		obj1.style.width=document.body.offsetWidth-280;
		obj4.style.width=document.body.offsetWidth-280;   // Y.TAKAKUWA Add-S 2014-07-19
		obj1.style.overflowX="auto";
		obj4.style.overflowX="hidden";                    // Y.TAKAKUWA Add-S 2014-07-19
	}else{
		obj1.style.width=document.body.offsetWidth-280;
		obj4.style.width=document.body.offsetWidth-280;  // Y.TAKAKUWA Add-S 2014-07-19
		obj1.style.overflowX="auto"; 
		obj4.style.overflowX="hidden";                   // Y.TAKAKUWA Add-S 2014-07-19
	}	
	//document.frm.ScreenWidth.value = document.body.offsetWidth
   if(obj1.clientWidth < obj2.scrollWidth)
   {
      obj1.style.height = rowHeight+17;
      obj1.style.overflowY="auto";                     
   }
   else{
      obj1.style.height = rowHeight;
      obj4.style.width=document.body.offsetWidth-280;  // Y.TAKAKUWA Add-S 2014-07-19
      obj1.style.overflowY="auto";                    
   } 
    var obj3=document.getElementById("BDIV3");
}

function getRowHeightThead()
{
  var oRows = document.getElementById('TBInOut3').getElementsByTagName('thead');  // Y.TAKAKUWA Upd-S 2014-07-19
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}
// Y.TAKAKUWA Add-S 2014-07-19
function getRowHeightThead2()
{
  var oRows = document.getElementById('TBInOut2').getElementsByTagName('thead');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}
// Y.TAKAKUWA Add-E 2014-07-19
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
function getRowHeightTbody2()
{
  var oRows = document.getElementById('TBInOut2').getElementsByTagName('td');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}
function getRowHeightTbody3()
{
  var oRows = document.getElementById('TBInOut3').getElementsByTagName('td');
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


function fSearch(){
	document.frm.Gamen_Mode.value = "S";
	document.getElementById("chk3").checked = true;
    document.frm.submit();
}

function fSearchGrpName(){
	document.frm.Gamen_Mode.value = "SGN";
    document.frm.submit();
}

function fSDriver(){
   document.frm.Gamen_Mode.value = "SD";
   document.frm.submit();
}

function fGDriver(){
   document.frm.Gamen_Mode.value = "SG";
   document.frm.submit();
}

function fDel()
{
  var chkFlag;
  chkFlag = 0;
  for(i=1; i <= (parseInt(document.frm.DataCnt2.value)-1); i++){
    obj = eval("document.frm.chkInOut2G0_" + i);
    if (obj.checked==true) {
       chkFlag = 1;
	}
  }
  
  if(chkFlag==1){
  var msg = confirm("選択したドライバを削除します。よろしいですか？",1,4,0);
    if(msg == true){
      document.frm.Gamen_Mode.value = "D";
      document.frm.submit();
    }
  }

}
function ClearSItem(stat)
{
  if(stat==1){ 
    document.frm.SDriverName.value=""
    document.frm.SDriverCompany.value=""
    document.frm.SDriverID.value=""
  }
  else{
    document.frm.SDriverID2.value=""
  }
  document.frm.Gamen_Mode.value = "SG";
  document.frm.submit();
}
function refreshParent() 
{
    if('<%=ErrMsg%>' ==""){
      var vmsg = "元の画面に反映するには、元の画面左の「コンテナロック」メニューをクリックして再描画してください。";
      if('<%=v_GamenMode%>'=='D'){
        if('<%=Num%>'=="0"){
          alert(vmsg);
        }
      }
      if('<%=v_GamenMode%>'=='GI'){
        if('<%=v_GroupNameChgFlag%>' != ''){
          alert(vmsg);
        }
      }
      if('<%=v_GamenMode%>'=='DG'){
        alert(vmsg);
      }
      if('<%=v_GamenMode%>'=='R'){
        if('<%=Num%>'!="0" && '<%=v_DataCnt%>'=="0"){
            alert(vmsg);
        }
      }
    }
    
}

function fCheckAllGroup(GNum,obj){
  var elementName;
  for(var i=1;i<=parseInt('<%=Num%>');i++){ 
    elementName = "chkInOutG" + GNum + "_" + i.toString(); //+ 1.toString() + "_" + 1.toString();
    if(obj.checked==true){
      document.getElementById(elementName).checked=true;
    }
    else{
      document.getElementById(elementName).checked=false;    
    }
    
  }
}

// Y.TAKAKUWA Add-S 2014-07-08
function RegisterOwnG()
{
   var chkFlag;
   var x;
   var i;
 
   chkFlag = 1
   if(chkFlag == 1){
      var x = confirm("        上の内容で更新します" + "\r\n             よろしいですか？");
      if (x == true) {
		 document.frm.Gamen_Mode.value = "RO"
         document.frm.submit()
     }    
   }
}

function RegisterDriverToGroup()
{
   var chkFlag;
   var x;
   var i;
  
  chkFlag = 1
  
  if(chkFlag == 1){
     var x = confirm("         この内容で更新します" +  "\r\n             よろしいですか？");
      if (x == true) {
		 document.frm.Gamen_Mode.value = "R"
         document.frm.submit()
     }    
  }
}

function UpdateGroup()
{
   var chkFlag;
   var x;
   var i;
   
   chkFlag = 0

  if(document.getElementById("SGroupName").value != "<%=v_GroupName%>") 
  {
     chkFlag = 1
  }

  if(chkFlag == 1){
     var x = confirm("         グループ名を更新します。" + "\r\n             よろしいですか？");
      if (x == true) {
        document.frm.Gamen_Mode.value = "GI"
        document.frm.submit()
     }    
  }
}
</SCRIPT>
<script type="text/vbscript">
Public Sub Delete_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 0
  
  for i = 1 to CInt(document.frm.DataCnt.value-1)
     If document.frm.elements("chkInOut" + CStr(i)).checked then
       chkFlag = 1
     end if
  Next
  
  if chkFlag=1 then
    x=MsgBox("選択したドライバをグループから除外します。" & vbCrLf & "       （ドライバ情報自体は残ります）" & vbCrLf & "               よろしいですか？",4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "D"
      document.frm.submit()
    end if
  end if

End Sub


Public Sub DeleteGroup_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 1
  
  if chkFlag=1 then
    x=MsgBox("        このグループを削除します。" & vbCrLf & "    （ドライバ情報自体は残ります）" & vbCrLf & "            よろしいですか？",4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "DG"
      document.frm.submit()
    end if
  end if
End Sub


Public Sub RegisterOwn_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 1

  if chkFlag=1 then
    x=MsgBox("        上の内容で更新します" & vbCrLf & "             よろしいですか？" ,4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "RO"
      document.frm.submit()
    end if
  end if
End Sub

Public Sub Update_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 0
  
  If document.frm.elements("SGroupName").value <> "<%=v_GroupName%>" Then
     chkFlag = 1
  End If
  
  if chkFlag=1 then
    x=MsgBox("         グループ名を更新します。" & vbCrLf & "             よろしいですか？" ,4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "GI"
      document.frm.submit()
    end if
  end if
End Sub

Public Sub Register_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 1
  
  if chkFlag=1 then
    x=MsgBox("         この内容で更新します" & vbCrLf & "             よろしいですか？",4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "R"
      document.frm.submit()
    end if
  end if
End Sub
</script>

</HEAD>
<BODY onLoad="finit();view();refreshParent();MakeStaticHeader('TBInOut2', 300, 400, 52, true);" onResize="view();MakeStaticHeader('TBInOut2', 300, 400, 52, true);">
<form name="frm" method="post">
<!--Hidden Values Start-->
<INPUT type=hidden name="Gamen_Mode" size="9" readonly tabindex= -1>
<!--<INPUT name="ScreenWidth" size="9" readonly tabindex= -1>-->
<!--Hidden Values End-->
<!--HEADER SECTION START-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
   <td rowspan="2">
	 <IMG height=73 src="Image/Group.png" width=507>
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
<!--HEADER SECTION END-->

<!--DETAIL SECTION START-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="110" nowrap>&nbsp;</td>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <!--Search Condition Start-->
      <tr nowrap>
        <td nowrap>
          <a name="table1"></a>
          <BR />
	      <DIV style="width:230px; padding:10px;background-color:#FFCCFF; text-align:center;">現在の登録状況の表示と更新</DIV>
	      <BR />
	      <DIV style="margin-left:30px">
		   
	        <table>  
	          <tr>
	            <td nowrap><input type="radio" name="selectCompany" id="chk3" value="3" checked=true onclick="ClearSItem(1);">グループ登録ドライバを表示
	              <div style="margin-left:25px; padding-top:8px; padding-bottom:5px;">
	              <table border="0" cellspacing="0" cellpadding="0">
	                <tr>
	                  <td nowrap style="width:88px">グループ選択</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td> 
	                  <td nowrap>
	                    <select name="cmbGroup2" style="width:115px;" onchange="fSearch();">
				          <OPTION VALUE = 'G0' SELECTED>G0</OPTION>
				          <OPTION VALUE = 'G1'>G1</OPTION>
				          <OPTION VALUE = 'G2'>G2</OPTION>
				          <OPTION VALUE = 'G3'>G3</OPTION>
				          <OPTION VALUE = 'G4'>G4</OPTION>
				          <OPTION VALUE = 'G5'>G5</OPTION>
				          <OPTION VALUE = 'G6'>G6</OPTION>
				          <OPTION VALUE = 'G7'>G7</OPTION>
				          <OPTION VALUE = 'G8'>G8</OPTION>
				          <OPTION VALUE = 'G9'>G9</OPTION>
			            </select>
	                  </td>
	                  <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	                  <td nowrap>
	                    <%If Trim(v_GroupName2) <> "" Then %>
	                      <%=v_GroupName2%>
	                    <%else%>
	                      グループ名なし
	                    <%end if%>
	                  </td>
	                </tr>
	              </table>
	              </div>
	            </td><td></td>
	          </tr>
	          <tr><td nowrap><input type="radio" name="selectCompany" id="chk1" value="1" onclick="ClearSItem(1);">自社承認ドライバを全員表示</td><td></td></tr>
	          <tr><td nowrap><input type="radio" name="selectCompany" id="chk2" value="2" onclick="ClearSItem(1);">グループ登録済他社承認ドライバを表示</td><td></td></tr>
	          <tr><td nowrap><br /></td></tr>
	          <tr><td nowrap><b>（表示条件の追加）</b>&nbsp;※部分一致可</td></tr>
	        </table>
	        
	        <div style="margin-left:25px;">
	          <table>
	            <tr><td nowrap style="width:100px">名前検索</td><td nowrap><input type="text" name="SDriverName" value="" onfocus="this.select();"></td></tr>
	            <tr><td nowrap style="width:100px">会社名検索</td><td nowrap><input type="text" name="SDriverCompany" value="" onfocus="this.select();"></td></tr>
	            <tr>
	              <td nowrap>ドライバＩＤ検索</td><td nowrap><input type="text" name="SDriverID" value="" onfocus="this.select();"></td>
	              <td width=90 align=right nowrap><input type="button" name="Button" value="表示更新" onClick="fGDriver();"></td>
	              
	            </tr>
	          </table>
	        </div>
	      </DIV>
        </td>
      </tr>
      <!--Search Condition End-->
      <!--Search Result Pagination Start-->
      <tr align=right nowrap>
        <td width="100%" height="30" align=right nowrap>
          <table border="0" cellpadding="0" cellspacing="0">
            <tr>
		      <td width="100%" align="center" nowrap>
		      <!--Page Pagination Start-->
		        <%					
				  If Num > 0 Then						
					abspage = ObjRS.AbsolutePage
					pagecnt = ObjRS.PageCount
					call LfPutPage(Num,abspage,pagecnt,"pagenum","table1")
				  End If									
			    %>
		      <!--Page Pagination End-->
		      </td>
		    </tr>
		  </table>
        </td>
      </tr>
      <!--Search Result Pagination End-->
      <!--Search Result Start-->
      <tr>
        <td nowrap>
	      <!--Y.TAKAKUWA Add-S 2014-07-19-->
		  <!-- ARVEE DIVISION PATTERN-S -->
          <div style="overflow: hidden; background-color:#fff; margin-left: 25px; " id="DivHeaderRow"></div>
		  <!--Y.TAKAKUWA Add-E 2014-07-19-->
		  <!-- ARVEE DIVISION PATTERN-E -->
		  <div id="BDIV1" style="margin-left:25px;">
			<% If Num>0 Then%>
			<!--Driver List Start-->
            <!--Y.TAKAKUWA Add-S 2014-07-19-->
            <!-- ARVEE PATTERN - S -->			
            <table border="1" cellpadding="1" cellspacing="0" width=100% id="TBInOut2">	
            <thead>
			   <!--HEADER INFORMATION START-->
			   <tr>
				 <th id="H2Col01" class="hlist" nowrap>氏名</th>
				 <th id="H2Col02" class="hlist" nowrap>ドライバID</th>								
				 <th id="H2Col03" class="hlist" nowrap>会社名</th>
				 <th id="H2Col04" class="hlist" align="center" nowrap >G0<br />
				     <input type="checkbox" name="G0_AllCheck" id="G0_AllCheck" onclick="fCheckAllGroup('0',this);">
				 </th>
				 <th id="H2Col05" class="hlist" align="center" nowrap>G1<br />
				      <input type="checkbox" name="G1_AllCheck" id="G1_AllCheck" onclick="fCheckAllGroup('1',this);">
				 </th>
				 <th id="H2Col06" class="hlist" align="center" nowrap>G2<br />
				      <input type="checkbox" name="G2_AllCheck" id="G2_AllCheck" onclick="fCheckAllGroup('2',this);">
				 </th>
				 <th id="H2Col07" class="hlist" align="center" nowrap>G3<br />
				      <input type="checkbox" name="G3_AllCheck" id="G3_AllCheck" onclick="fCheckAllGroup('3',this);">
				 </th>
				 <th id="H2Col08" class="hlist" align="center" nowrap>G4<br />
				      <input type="checkbox" name="G4_AllCheck" id="G4_AllCheck" onclick="fCheckAllGroup('4',this);">
				 </th>
				 <th id="H2Col10" class="hlist" align="center" nowrap>G5<br /><input type="checkbox" name="G5_AllCheck" id="G5_AllCheck" onclick="fCheckAllGroup('5',this);"></th>
				 <th id="H2Col11" class="hlist" align="center" nowrap>G6<br /><input type="checkbox" name="G6_AllCheck" id="G6_AllCheck" onclick="fCheckAllGroup('6',this);"></th>
				 <th id="H2Col12" class="hlist" align="center" nowrap>G7<br /><input type="checkbox" name="G7_AllCheck" id="G7_AllCheck" onclick="fCheckAllGroup('7',this);"></th>
				 <th id="H2Col13" class="hlist" align="center" nowrap>G8<br /><input type="checkbox" name="G8_AllCheck" id="G8_AllCheck" onclick="fCheckAllGroup('8',this);"></th>
				 <th id="H2Col14" class="hlist" align="center" nowrap>G9<br /><input type="checkbox" name="G9_AllCheck" id="G9_AllCheck" onclick="fCheckAllGroup('9',this);"></th>																																		
			   </tr>
			   </thead>
			   <tbody>
			   <tr bgcolor="#CCFFFF">	
			     <td id="D2Col01" align="center" valign="middle" nowrap style="border-bottom:0px;width:120px;"></td>
			     <td id="D2Col02" align="center" valign="middle" nowrap style="border-bottom:0px;width:120px;"></td>
			     <td id="D2Col03" align="center" valign="middle" nowrap style="border-bottom:0px;width:200px;"></td>
			     <td id="D2Col04" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col05" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col06" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col07" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col08" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col09" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col10" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col11" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col12" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			     <td id="D2Col13" align="center" width="30px" align="center" nowrap style="border-bottom:0px;"></td>
			   </tr>		        									
			   </tbody>	
			</table>
			<!-- ARVEE PATTERN-E -->
			<!--Y.TAKAKUWA Add-E 2014-07-19-->
			<!--Y.TAKAKUWA Upd-S 2014-07-19-->
                        <table border="1" cellpadding="1" cellspacing="0" width=100% id="TBInOut">			   
			   <tbody>
			   <!--DETAIL INFORMATION START-->
                            <% 
								x = 1 							
								For i=1 To ObjRS.PageSize
								 	If Not ObjRS.EOF Then
									x = x + 1
							%>
							<tr bgcolor="#CCFFFF">	
							
							  <td id="D2Col01" align="center" valign="middle" nowrap style="width:120px;border-top:0px;">
                                <%=Trim(ObjRS("LoDriverName"))%><BR />
                              </td>
							  <td id="D2Col02" align="center" valign="middle" nowrap style="width:120px;border-top:0px;">
                                <%=Trim(ObjRS("LoDriverID"))%><BR />
                              </td>
							  
							  <td id="D2Col03" align="center" valign="middle" nowrap style="width:200px;border-top:0px;">
                                <%=Trim(ObjRS("LoDriverCompany"))%><BR />
                              </td>
							  <% v_ItemName = "chkInOutG0_" + cstr(i) %>
							  <td id="D2Col04" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group0")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group0_cnt = Group0_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  
							  <% v_ItemName = "chkInOutG1_" + cstr(i) %>
							  <td id="D2Col05" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group1")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group1_cnt = Group1_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG2_" + cstr(i) %>
							  <td id="D2Col06" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group2")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group2_cnt = Group2_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG3_" + cstr(i) %>
							  <td id="D2Col07" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group3")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group3_cnt = Group3_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  
							  <% v_ItemName = "chkInOutG4_" + cstr(i) %>
							  <td id="D2Col08" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group4")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group4_cnt = Group4_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG5_" + cstr(i) %>
							  <td id="D2Col09" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group5")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group5_cnt = Group5_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG6_" + cstr(i) %>
							  <td id="D2Col10" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group6")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group6_cnt = Group6_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG7_" + cstr(i) %>
							  <td id="D2Col11" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group7")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group7_cnt = Group7_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG8_" + cstr(i) %>
							  <td id="D2Col12" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group8")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group8_cnt = Group8_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
							  <% v_ItemName = "chkInOutG9_" + cstr(i) %>
							  <td id="D2Col13" align="center" width="30px" align="center" nowrap style="border-top:0px;">
							    <% If Trim(ObjRS("Group9")) <> "" Then %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>" checked=true><BR>
							      <%Group9_cnt = Group9_cnt + 1%>
							    <% Else %>
							      <input type="checkbox" name="<%= v_ItemName %>" id="<%= v_ItemName %>"><BR>
							    <%End If %>
							  </td>
							  
                              <% v_ItemName = "LODriverID" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS("LODriverID"))%>">
							  <% v_ItemName = "HiTSUserID" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS("HiTSUserID"))%>">
							</tr>
						    <% 
									ObjRS.MoveNext 		
									End If
								Next	
							  ObjRS.close    
						      DisConnDBH ObjConn, ObjRS
						    %>  
						    
						    <script type="text/javascript">
						     <% Dim strGrpCnt
						        Dim strGrpNameChk
						        Dim remainder
						        Dim wholenumber
						        Dim roundnumber
						        Dim maxCnt
						        For j = 0 to 9
						          Select Case j
						            Case 0 
						              strGrpCnt = Group0_cnt
						            Case 1
						              strGrpCnt = Group1_cnt
						            Case 2
						              strGrpCnt = Group2_cnt
						            Case 3
						              strGrpCnt = Group3_cnt
						            Case 4 
						              strGrpCnt = Group4_cnt
						            Case 5
						              strGrpCnt = Group5_cnt
						            Case 6
						              strGrpCnt = Group6_cnt
						            Case 7
						              strGrpCnt = Group7_cnt
						            Case 8
						              strGrpCnt = Group8_cnt
						            Case 9
						              strGrpCnt = Group9_cnt
						          End Select
						          strGrpNameChk = "G" & j & "_AllCheck"
						          
						          If Num > 100 then
						            remainder = Num mod 100
						            If remainder > 0 then
						              wholenumber = CInt(Num/100)
						              roundnumber = Num/100
						              If roundnumber > wholenumber then
						                wholenumber = wholenumber + 1
						              End If
						              If CStr(wholenumber) = CStr(abspage) then
						                maxCnt = remainder * 10
						              Else
						                maxCnt = 100
						              End If
						            Else
						              maxCnt = Num
						            end If
						          Else
						            maxCnt = Num
						          End If
						          
						     %>
						       if('<%=maxCnt%>' == '<%=strGrpCnt%>'){
						         document.getElementById('<%=strGrpNameChk%>').checked=true;
						       }
						     <% Next %>
						     
						     
						     <%
                               'dim icnt
                               'dim icnt2
                               'for icnt = 0 to 9    
                                 'if Trim(Arr_OwnGroup(icnt)) = "" then
                                    'response.write("document.frm.G" & icnt & "_AllCheck.disabled=true;" & vbCrlf)
                                    'for icnt2 = 1 to (x-1)
                                       'response.write("document.frm.chkInOutG" & icnt & "_" & icnt2 & ".disabled=true;" & vbCrlf)
                                    'Next
                                 'End If 
                               'Next
                               
                             %>  
					
						    </script>
						    <!--DETAIL INFORMATION END-->	    									
			   </tbody>								
			 </table>
			 <!--Y.TAKAKUWA Upd-E 2014-07-19-->
			 <!--Driver List End-->
			 <INPUT type=hidden name="DataCnt" value="<%=x%>">
			 <!--NO DATA START-->
			 <% Else %>
			   <div style="margin-left:25px;">	    
			   <table border="1" cellPadding="2" cellSpacing="0" id="NODATA">						
				 <TR class=bgw><TD nowrap>ドライバーの登録がありません</TD></TR>
			   </table>
			   </div>		
			 <% End If %>	
			 <!--NO DATA END-->
		  </div>
		<script>
		$('#BDIV1').on('scroll', function () {
                    $('#DivHeaderRow').scrollLeft($(this).scrollLeft());
                });
		</script>

		</td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td><div style="margin-left:60px">※チェックをはずすことでグループから除外できます</div></td></tr>
      <!--Search Result End-->
      <tr><td>&nbsp;</td></tr>
      <!--Register Buttons Start-->
      <tr>		
		<td>
		  <div style="margin-left:25px;">
		  <table border="0" cellpadding="2" cellspacing="0">
		    <tr>
		      <% If Num>0 Then%>
			   <!-- Y.TAKAKUWA Upd-S -->
			   <!--
		       <td><input type="button" name="RegisterOwn" value="上の内容で更新"></td>
		       -->
		       <td><input type="button" name="RegisterOwn" onclick="RegisterOwnG()" value="上の内容で更新"></td>
			   <!-- Y.TAKAKUWA Upd-E -->
		       <!--
		       <!--
			   <td><input type="button" name="Delete"  value="選択したドライバを除外"></td>
			   <td><input type="button" name="DeleteGroup"  value="このグループを削除"></td>
			   -->
			  <% Else %>
			   <!-- Y.TAKAKUWA Upd-S -->
			   <!--
			   <td><input type="button" name="RegisterOwn"  value="上の内容で更新" disabled></td>
			   -->
			   <td><input type="button" name="RegisterOwn" onclick="RegisterOwnG()" value="上の内容で更新" disabled></td>
			   <!-- Y.TAKAKUWA Upd-E -->
			   <!--
			   <td><input type="button" name="Delete"  value="選択したドライバを除外" disabled></td>
			   <td><input type="button" name="DeleteGroup"  value="このグループを削除" disabled></td>
			   -->
			  <%End If %>
			</tr>
		  </table>
		  </div>		
		</td>
	  </tr> 
      <!--Register Buttons End-->
      <tr>
        <td>
           <BR />
           <BR />
           <hr />
        </td>
      </tr>
      <!--OTHER COMPANIES APPROVAL-S-->
      <tr>
        <td>
          <DIV style="width:230px; padding:10px;background-color:#FFCCFF; text-align:center;">他社承認ドライバをグループに追加</DIV>
          <BR />
          <div style="margin-left:30px;">
	          <table>
	            <tr><td nowrap style="width:85px">ドライバID検索</td><td nowrap><input type="text" name="SDriverID2" value="" onfocus="this.select();"></td><td><input type="button" name="Select" OnClick="fSDriver();" value="検索"></td><td>※完全一致のみ</td></tr>
	          </table>
	      </div> 
        </td>
      </tr>
      <tr>
        <td>
            <!--LIST HERE-->
        </td>
      </tr>
      
      
      <!--Search Result Pagination Start-->
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
					call LfPutPage(Num2,abspage,pagecnt,"pagenum2","table2")
				  End If									
			    %>
		      <!--Page Pagination End-->
		      </td>
		    </tr>
		  </table>
        </td>
      </tr>
      <!--Search Result Pagination End-->
      <!--Search Result Start-->
      <tr>
        <td nowrap>
          
		  <div id="BDIV2" style="margin-left:25px;">
		     <a name="table2"></a>
			 <% If Num2>0 Then%>
			 <!--Driver List Start-->	
			 <table border="1" cellpadding="1" cellspacing="0" width=100% id="TBInOut3">	<!--Y.TAKAKUWA Upd-S 2014-07-19-->			
			   <thead>
			   <!--HEADER INFORMATION START-->
			   <tr>
				 <th id="Th1" class="hlist" nowrap style="width:120px;">氏名</th> 
				 <!--							
				 <th id="Th2" class="hlist" nowrap>携帯番号</th>
				 -->
				 <th id="Th3" class="hlist" nowrap style="width:320px;">会社名</th>
				 <!--
				 <th id="Th4" class="hlist" nowrap>メールアドレス</th>	
				 -->
				 <th id="Th5" class="hlist" align="center" nowrap>G0<br /></th>
				 <th id="Th6" class="hlist" align="center" nowrap>G1<br /></th>																																	
                                 <th id="Th7" class="hlist" align="center" nowrap>G2<br /></th>																																	
                                 <th id="Th8" class="hlist" align="center" nowrap>G3<br /></th>																																	
                                 <th id="Th9" class="hlist" align="center" nowrap>G4<br /></th>																																	
                                 <th id="Th10" class="hlist" align="center" nowrap>G5<br /></th>																																	
                                 <th id="Th11" class="hlist" align="center" nowrap>G6<br /></th>																																	
                                 <th id="Th12" class="hlist" align="center" nowrap>G7<br /></th>																																	
                                 <th id="Th13" class="hlist" align="center" nowrap>G8<br /></th>																																
                                 <th id="Th14" class="hlist" align="center" nowrap>G9<br /></th>
			   </tr>
			   <!--HEADER INFORMATION END-->
			   </thead>																
			   <tbody>
			   <!-- Y.TAKAKUWA Upd-S 2014-07-22 -->
			   <!--DETAIL INFORMATION START-->
                            <% 
								x = 1 
															
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
								 	
									x = x + 1
							%>
							<tr bgcolor="#CCFFFF">	
							<td id="Td1" align="center" valign="middle" nowrap style="border-top:0px;">
                              <%=Trim(ObjRS2("LoDriverName"))%><BR />
                            </td>
                            <!--
							<td id="Td2" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("PhoneNum"))%><BR />
                            </td>
                            -->
							<td id="Td3" align="center" valign="middle" nowrap style="border-top:0px;">
                              <%=Trim(ObjRS2("LoDriverCompany"))%><BR />
                            </td>
                            <!--
							<td id="Td4" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("MailAddress"))%><BR />
                            </td>
                            -->
                            <% v_ItemName = "chkInOut2G0_" + cstr(i) %>
							<td id="Td5" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group0")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G1_" + cstr(i) %>
							<td id="Td6" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group1")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G2_" + cstr(i) %>
							<td id="Td7" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group2")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G3_" + cstr(i) %>
							<td id="Td8" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group3")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G4_" + cstr(i) %>
							<td id="Td9" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group4")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G5_" + cstr(i) %>
							<td id="Td10" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group5")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G6_" + cstr(i) %>
							<td id="Td11" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group6")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G7_" + cstr(i) %>
							<td id="Td12" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group7")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G8_" + cstr(i) %>
							<td id="Td13" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group8")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
							<% v_ItemName = "chkInOut2G9_" + cstr(i) %>
							<td id="Td14" align="center" width="30" align="center" nowrap style="border-top:0px;">
							  <% If Trim(ObjRS2("Group9")) = "" Then %>
							    <input type="checkbox" name="<%= v_ItemName %>"><BR>
							  <% Else %>
							    <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR>
							  <% End If %>
							</td>
							
                            
                            <% v_ItemName = "LODriverID2_" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("LODriverID"))%>">
							
							<% v_ItemName = "HiTSUserID2_" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("HiTSUserID"))%>">
							
							</tr>
						    <% 
									ObjRS2.MoveNext 		
									End If
								Next	
							  ObjRS2.close    
						      DisConnDBH ObjConn2, ObjRS2
						    %>  
						    <!--DETAIL INFORMATION END-->
						    
						    
						    
						    <script type="text/javascript">
						     <%
                               'dim i2cnt
                               'dim i2cnt2
                               'for i2cnt = 0 to 9    
                                 'if Trim(Arr_OwnGroup(i2cnt)) = "" then
                                    'for i2cnt2 = 1 to (x-1)
                                       'response.write("document.frm.chkInOut2G" & i2cnt & "_" & i2cnt2 & ".disabled=true;" & vbCrlf)
                                    'Next
                                 'End If 
                               'Next
                             %>  
					
						    </script>	    									
			   </tbody>	
                           <!-- Y.TAKAKUWA Upd-E 2014-07-22 -->			   
			 </table>
			 <!--Driver List End-->
			 <INPUT type=hidden name="DataCnt2" value="<%=x%>">
			 <!--NO DATA START-->
			 <% Else %>	   
			   <div style="margin-left:10px"> 
                 <%If v_GamenMode = "SD2" AND Trim(SDriverID2) <> "" then %>
			     <table border="1" cellPadding="2" cellSpacing="0" id="Table2">						
				   <TR class=bgw>
				     <TD nowrap>
				       ドライバーの登録がありません
				     </TD>
				   </TR>
			     </table>	
				 <%End If%>
			   </div>
			 <% End If %>	
			 <!--NO DATA END-->
			
		  </div>
		</td>
      </tr>
      <!--Search Result End--> 
      <tr><td>&nbsp;</td></tr>
 
      <!--Group Add Button Start-->
      <tr>		
		<td>
		  <div style="margin-left:25px">
		  <table border="0" cellpadding="2" cellspacing="0">
		    <tr>		  
		    <% If Num2>0 Then%>
			    <!-- Y.TAKAKUWA Upd-S 2014-07-08 -->
			    <td><input type="button" name="Register"  onClick="RegisterDriverToGroup();" value="上の内容で更新"></td>
		        <!--
			    <td><input type="button" name="Register"  value="上の内容で更新"></td>
			    -->
			    <!-- Y.TAKAKUWA Upd-E 2014-07-08 -->
		    <% Else %>
			   <!-- Y.TAKAKUWA Upd-S 2014-07-08 -->
			   <!--
			   <td><input type="button" name="Register"  value="上の内容で更新" disabled></td>
			   -->
			   <td><input type="button" name="Register"  onClick="RegisterDriverToGroup();" value="上の内容で更新" disabled></td>
			   <!-- Y.TAKAKUWA Upd-E 2014-07-08 -->
		    <%End If %>
			</tr>
		  </table>		
		  </div>
		</td>
	  </tr> 
      <!--Group Add Button End-->
      
      <tr><td>&nbsp;</td></tr>
      <!--OTHER COMPANIES APPROVAL-E-->
      
      <tr>
        <td>
           <BR />
           <BR />
           <hr />
        </td>
      </tr>
      <!--GROUP NAME UPDATE START-->
      <tr>
        <td>
          <DIV style="width:230px; padding:10px;background-color:#FFCCFF; text-align:center;">グループ名の変更</DIV>
        </td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td>
          <div style="margin-left:25px">
          <a name="groupName"></a>
          <table>
            <tr>
              <td nowrap style="width:88px">グループ選択</td> 
	          <td nowrap>
	            <select name="cmbGroup" style="width:115px;" onchange="fSearchGrpName();">
				  <OPTION VALUE = 'G0' SELECTED>G0</OPTION>
				  <OPTION VALUE = 'G1'>G1</OPTION>
				  <OPTION VALUE = 'G2'>G2</OPTION>
				  <OPTION VALUE = 'G3'>G3</OPTION>
				  <OPTION VALUE = 'G4'>G4</OPTION>
				  <OPTION VALUE = 'G5'>G5</OPTION>
				  <OPTION VALUE = 'G6'>G6</OPTION>
				  <OPTION VALUE = 'G7'>G7</OPTION>
				  <OPTION VALUE = 'G8'>G8</OPTION>
				  <OPTION VALUE = 'G9'>G9</OPTION>
			    </select>
	          </td>
              <td>
                <input type="text" name="SGroupName" onfocus="this.select();" style="width:240px;" maxlength="20">
              </td>
              
              <td>
                <%If ErrMsg<>"" then%>
                <div style="color:Red;">&nbsp;<%=ErrMsg %></div>
                <%end if %>
              </td>
            </tr>
            <tr><td></td><td></td><td>※全角１０文字、半角２０文字まで</td><td></td><td></td><td></td></tr>
            
            <!-- Y.TAKAKUWA Upd-S 2014-07-08 -->
	　　　    <!--
            <tr><td></td><td></td><td><input type="button" name="Update"  value="更新"></td><td></td><td></td><td></td></tr> 
            -->	
            <tr><td></td><td></td><td><input type="button" name="Update" onClick="UpdateGroup();" value="更新"></td><td></td><td></td><td></td></tr>  			
            <!-- Y.TAKAKUWA Upd-S 2014-07-08 -->	
          </table>
        </td>
        </div>
      </tr>
      <!--GROUP NAME UPDATE END-->
      <tr>
        <td>
           <BR />
           <BR />
        </td>
      </tr>
      </table> 
    </td>
    <td width="120" nowrap>&nbsp;</td>
  </tr>
</table>
<!--DETAIL SECTION END-->

<!--FOOTER SECTION START-->  
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
<!--FOOTER SECTION END-->
</form>

</BODY>
</HTML>
