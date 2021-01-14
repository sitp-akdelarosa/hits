<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi115.asp				_/
'_/	Function	:空搬入入力情報取得			_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/07/29	備考欄追加	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="Common.inc"-->
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'セッションの有効性をチェック
  CheckLoginH

'データ取得
  dim CONnum,Mord
  dim hCd,sUN,Utype,User
  dim HedId,HTo,Rmon,Rday,TrhkSen,MrSk
  dim CONsize,CONtype,CONhite,CONsitu,CONtear,CMPcd,MaxW
  dim strNum,dummy, UpFlag,UpUser,ret1,ret2
  dim TruckerSubName,TruckerName
  CONnum = Request("CONnum")
  hCd    = Session.Contents("COMPcd")
  sUN    = Session.Contents("sUN")
  Utype  = Session.Contents("UType")
  User   = Session.Contents("userid")
  ret1   = true
  ret2   = true
'CW-036  UpFlag = 0
  UpFlag = 1
  dim Comment1		'C-002

'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
  If err <> 0 then
    response.write "エラー" & err & ": " & error 
    Num=0
  End If

'遷移元振分け
  If Request("InfoFlag") = "" Then
   '初期登録
    Mord=0
   '事前登録チェック
    strNum="'"& CONnum &"'"
   '輸入コンテナテーブルのコンテナ存在チェック
    checkImportCont ObjConn, ObjRS,strNum, "1", ret1
   'IT共通テーブルの登録重複チェック
    checkComInfo  ObjConn, ObjRS,strNum, "2","1", dummy, ret2

    dim tmpStr
    If ret1 AND ret2 Then
      tmpStr=",入力内容の正誤:0(正しい)"
    Else
      tmpStr=",入力内容の正誤:1(誤り)"
    End If
    WriteLogH "b202", "空搬入事前情報入力", "01",strNum&tmpStr

    If ret1 AND ret2 Then
      '指示先デフォルト表示
      dim rettmp,COMPcdR
      COMPcdR = ""
      StrSQL = "SELECT  Count(ITC.WkContrlNo) AS num "&_
               "FROM ImportCont AS INC INNER JOIN hITCommonInfo AS ITC ON INC.BLNo = ITC.BLNo "&_
               "WHERE INC.ContNo='"& CONnum &"' AND ITC.WorkCompleteDate Is Not Null "&_
               "AND ITC.Process='R' AND ITC.WkType='1' AND (ITC.FullOutType='2' OR ITC.FullOutType='4') "
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","115E-01","SQL:<BR>"&strSQL
      end if
      If ObjRS("num")<>0 Then
        ObjRS.close
        StrSQL = "SELECT ITC.TruckerSubCode1, ITC.WorkCompleteDate "&_
                 "FROM ImportCont AS INC INNER JOIN hITCommonInfo AS ITC ON INC.BLNo = ITC.BLNo "&_
                 "WHERE INC.ContNo='"& CONnum &"' AND ITC.WorkCompleteDate Is Not Null "&_
                 "AND ITC.Process='R' AND ITC.WkType='1' AND (ITC.FullOutType='2' OR ITC.FullOutType='4') "
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","115E-01","SQL:<BR>"&strSQL
        end if
        If DateDiff("d",ObjRS("WorkCompleteDate"),Now) <31 Then
          COMPcdR = Trim(ObjRS("TruckerSubCode1"))
        End If
        ObjRS.close
      Else
        ObjRS.close
        StrSQL = "SELECT  Count(ITC.WkContrlNo) AS num "&_
                 "FROM hITCommonInfo AS ITC LEFT JOIN hITFullOutSelect AS ITF ON ITC.WkContrlNo = ITF.WkContrlNo "&_
                 "WHERE (ITC.ContNo='"& CONnum &"' OR ITF.ContNo='"& CONnum &"' ) "&_
                 "AND ITC.WorkCompleteDate Is Not Null AND ITC.Process='R' AND ITC.WkType='1'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","115E-01","SQL:<BR>"&strSQL
        end if
        If ObjRS("num")<>0 Then
          ObjRS.close
          StrSQL = "SELECT ITC.TruckerSubCode1, ITC.WorkCompleteDate "&_
                 "FROM hITCommonInfo AS ITC LEFT JOIN hITFullOutSelect AS ITF ON ITC.WkContrlNo = ITF.WkContrlNo "&_
                 "WHERE (ITC.ContNo='"& CONnum &"' OR ITF.ContNo='"& CONnum &"' ) "&_
                 "AND ITC.WorkCompleteDate Is Not Null AND ITC.Process='R' AND ITC.WkType='1'"
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","115E-01","SQL:<BR>"&strSQL
          end if
          If DateDiff("d",ObjRS("WorkCompleteDate"),Now) <31 Then
            COMPcdR = Trim(ObjRS("TruckerSubCode1"))
          End If
        End If
        ObjRS.close
      End If
    'コンテナデータ所得
    StrSQL = "SELECT CN.ContType, CN.ContSize, CN.ContHeight,CN.TareWeight, CN.ContWeight, CN.Material, "&_
             "IPC.ReturnPlace, SL.FullName "&_
             "FROM (Container AS CN LEFT JOIN mShipLine AS SL ON CN.ShipLine = SL.ShipLine) "&_
             "INNER JOIN ImportCont AS IPC ON CN.ContNo = IPC.ContNo AND CN.VoyCtrl = IPC.VoyCtrl "&_
             "AND CN.VslCode = IPC.VslCode WHERE CN.ContNo='"& CONnum &"'"
     ObjRS.Open StrSQL, ObjConn
     if err <> 0 then
       DisConnDBH ObjConn, ObjRS	'DB切断
       jampErrerP "1","115E-01","SQL:<BR>"&strSQL
     end if
'CW-035     CMPcd   =Array(User,"","","","")
     CMPcd   =Array(UCase(User),COMPcdR,"","","")
'3th cahge start
'     Rmon=0
'     Rday=0
     Rmon=Null
     Rday=Null
'3th cahge end
     CONsitu  = ""
     Comment1 = ""		'C-002
    End If
  Else
    '更新
    Mord=1
    dim WkCNo,TruckerFlag,compFlag
    WkCNo = Request("WkconNo")		'ADD20030911
    StrSQL = "SELECT ITC.WkContrlNo, ITC.HeadID, ITC.ContSize, ITC.ContType, ITC.ContHeight, ITC.Material,"&_
             "ITC.TareWeight, ITC.CustOK, ITC.MaxWght, ITC.UpdtUserCode, ITC.WorkDate, ITC.RegisterCode, "&_
             "ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4, ITC.Comment1, "&_
             "ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
             "IPC.ReturnPlace, SL.FullName, ITC.WorkCompleteDate, "&_
             "ITC.TruckerSubName1, ITC.TruckerSubName2, ITC.TruckerSubName3, ITC.TruckerSubName4, ITC.TruckerSubName5, "&_
             "T1.Trucked AS Trucked1, T2.Trucked AS Trucked2, T3.Trucked AS Trucked3, T4.Trucked AS Trucked4 "&_
             "FROM (hITCommonInfo AS ITC INNER JOIN ((Container AS CN LEFT JOIN mShipLine AS SL "&_
             "ON CN.ShipLine = SL.ShipLine) INNER JOIN ImportCont AS IPC ON (CN.ContNo = IPC.ContNo) "&_
             "AND (CN.VoyCtrl = IPC.VoyCtrl) AND (CN.VslCode = IPC.VslCode)) ON ITC.ContNo = IPC.ContNo) "&_
             "INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
             "LEFT JOIN mTrucker T1 ON (ITC.TruckerSubCode1 = T1.HeadCompanyCode) "&_
             "LEFT JOIN mTrucker T2 ON (ITC.TruckerSubCode2 = T2.HeadCompanyCode) "&_
             "LEFT JOIN mTrucker T3 ON (ITC.TruckerSubCode3 = T3.HeadCompanyCode) "&_
             "LEFT JOIN mTrucker T4 ON (ITC.TruckerSubCode4 = T4.HeadCompanyCode) "&_
             "WHERE ITC.ContNo='"& CONnum &"' AND ITC.WkContrlNo='"& WkCNo &"' AND ITC.Process='R' AND ITC.WkType='2'"
'CW-048"FROM (hITCommonInfo AS ITC INNER JOIN ((Container AS CN INNER JOIN mShipLine AS SL "&_
'20030911 ADD this Item:"ITC.WkContrlNo='"& WkCNo &"' AND "&_
'C-002 ADD : ITC.Comment1,
     ObjRS.Open StrSQL, ObjConn
     if err <> 0 then
       DisConnDBH ObjConn, ObjRS	'DB切断
       jampErrerP "1","115E-01","SQL:<BR>"&strSQL
     end if
'20030911 Dell     WkCNo    =Trim(ObjRS("WkContrlNo"))
     CMPcd  = Array("","","","","")
     CMPcd(0)  = Trim(ObjRS("RegisterCode"))
     CMPcd(1)  = Trim(ObjRS("TruckerSubCode1"))
     CMPcd(2)  = Trim(ObjRS("TruckerSubCode2"))
     CMPcd(3)  = Trim(ObjRS("TruckerSubCode3"))
     CMPcd(4)  = Trim(ObjRS("TruckerSubCode4"))
'CW-018    dim TmpA
'CW-018    TmpA   = Split(ObjRS("WorkDate"), "/", -1, 1)
'CW-018    If ObjRS("WorkDate") = "1900/01/01" Then	'日付がNullであった場合
    Dim TmpA
    If ObjRS("WorkDate") = "1900/01/01" Or IsNull(ObjRS("WorkDate")) Then	'日付がNullであった場合	'CW-018
       Rmon   = Null
       Rday   = Null
    Else
'3th chage       dim TmpA						'CW-018
'3th chage       TmpA   = Split(ObjRS("WorkDate"), "/", -1, 1)	'CW-018
'3th chage       Rmon   = TmpA(1)
'3th chage       Rday   = TmpA(2)
      TmpA   = Split(Left(ObjRS("WorkDate"),10), "/", -1, 1)
      Rmon   = TmpA(1)
      Rday   = TmpA(2)
    End If
     MrSk    =Trim(ObjRS("CustOK"))
     MaxW    =Trim(ObjRS("MaxWght"))
     UpUser  =Trim(ObjRS("UpdtUserCode"))
     compFlag  = isNull(ObjRS("WorkCompleteDate"))
     Comment1  = Trim(ObjRS("Comment1"))		'C-002

    'ログインユーザによって会社コード表示制御
     chengeCompCd CMPcd, UpFlag
     If UpFlag <> 5 Then
       TruckerFlag= Trim(ObjRS("TruckerFlag"&UpFlag))
     Else
       TruckerFlag = 0
     End If

    'ログインユーザによってヘッドID表示制御
     HedId  = Trim(ObjRS("HeadID"))
     IF TruckerFlag = 1 Then 
       HedId  = "*****"
     End If
     
'2009/08/04 Tanaka Upd-S    
'' 2009/03/10 R.Shibuta Add-S
''ログインユーザを元に担当者名を選択
'	Select Case User
'		Case Trim(ObjRS("RegisterCode"))
'			TruckerSubName = ObjRS("TruckerSubName1")
'		Case Trim(ObjRS("Trucked1"))
'			TruckerSubName = ObjRS("TruckerSubName2")
'		Case Trim(ObjRS("Trucked2"))
'			TruckerSubName = ObjRS("TruckerSubName3")
'		Case Trim(ObjRS("Trucked3"))
'			TruckerSubName = ObjRS("TruckerSubName4")
'		Case Trim(ObjRS("Trucked4"))
'			TruckerSubName = ObjRS("TruckerSubName5")
'		Case Else
'			TruckerSubName = ""
'	End Select 
'' 2009/03/10 R.Shibuta Add-E
	Select Case User
		Case Trim(ObjRS("RegisterCode"))
			TruckerSubName = ObjRS("TruckerSubName1")
			TruckerName = ObjRS("TruckerSubName1")
		Case Trim(ObjRS("Trucked1"))
			TruckerSubName = ObjRS("TruckerSubName1")
			TruckerName = ObjRS("TruckerSubName2")
		Case Trim(ObjRS("Trucked2"))
			TruckerSubName = ObjRS("TruckerSubName2")
			TruckerName = ObjRS("TruckerSubName3")
		Case Trim(ObjRS("Trucked3"))
			TruckerSubName = ObjRS("TruckerSubName3")
			TruckerName = ObjRS("TruckerSubName4")
		Case Trim(ObjRS("Trucked4"))
			TruckerSubName = ObjRS("TruckerSubName4")
			TruckerName = ObjRS("TruckerSubName5")
		Case Else
			TruckerSubName = ""
	End Select 
'2009/08/04 Tanaka Upd-E    
  End If


'データ設定
  HTo     =Trim(ObjRS("ReturnPlace"))
  TrhkSen =Trim(ObjRS("FullName"))
  CONsize =Trim(ObjRS("ContSize"))
  CONtype =Trim(ObjRS("ContType"))
  CONhite =Trim(ObjRS("ContHeight"))
  '2016/10/24 H.Yoshikawa Upd Start
  'CONsitu =Trim(ObjRS("Material"))
  CONsitu =""
  '2016/10/24 H.Yoshikawa Upd End
'Modified 2003.7.26
'  CONtear =Trim(ObjRS("TareWeight"))
  'If Request("InfoFlag") = "" Then				'2016/10/24 H.Yoshikawa Del
  '  CONtear =ObjRS("TareWeight")*100			'2016/10/24 H.Yoshikawa Del
  'Else											'2016/10/24 H.Yoshikawa Del
    CONtear =ObjRS("TareWeight")
  'End if										'2016/10/24 H.Yoshikawa Del
'Modification END 2003.7.26

'DB接続解除
  ObjRS.close
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>データ検索中</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function GoNext(){
<% IF ret1 AND ret2 Then %>
  mord=<%=Mord%>;
  target=document.dmi115F;
  if(mord==0){
    target.action="./dmi120.asp";
  } else {
    target.action="./dmo120.asp";
  }
  target.submit();
<%Else%>
  window.resizeTo(500,500);
<%End If%>
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="GoNext();">
<% IF ret1 AND ret2 Then %>
<!-------------DB検索用画面--------------------------->
<FORM name="dmi115F" method="POST">
<P>データ検索中<BR>
しばらくお待ちください</P>
<INPUT type=hidden name="Mord"    value="<%=Mord%>">
<INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
<INPUT type=hidden name="UpUser"  value="<%=UpUser%>">
<INPUT type=hidden name="CONnum"  value="<%=CONnum%>">
<INPUT type=hidden name="CMPcd0"  value="<%=CMPcd(0)%>">
<INPUT type=hidden name="CMPcd1"  value="<%=CMPcd(1)%>">
<INPUT type=hidden name="CMPcd2"  value="<%=CMPcd(2)%>">
<INPUT type=hidden name="CMPcd3"  value="<%=CMPcd(3)%>">
<INPUT type=hidden name="CMPcd4"  value="<%=CMPcd(4)%>">
<INPUT type=hidden name="TruckerSubName" value="<%=Trim(TruckerSubName)%>">

 <%' 2009/08/04 Tanaka Add-S %>
  <INPUT type=hidden name="TruckerName" value="<%=Trim(TruckerName)%>">
 <%' 2009/08/04 Tanaka Add-E %>
<INPUT type=hidden name="HTo"     value="<%=HTo%>">
<INPUT type=hidden name="CONsize" value="<%=CONsize%>">
<INPUT type=hidden name="CONtype" value="<%=CONtype%>">
<INPUT type=hidden name="CONhite" value="<%=CONhite%>">
<INPUT type=hidden name="CONsitu" value="<%=CONsitu%>">
<INPUT type=hidden name="CONtear" value="<%=CONtear%>">
<INPUT type=hidden name="TrhkSen" value="<%=TrhkSen%>">
<INPUT type=hidden name="Rmon"    value="<%=Rmon%>">
<INPUT type=hidden name="Rday"    value="<%=Rday%>">
<INPUT type=hidden name="Comment1" value="<%=Comment1%>" ><%'C-002 ADD START%>
<% IF Mord = 1 Then %>
<INPUT type=hidden name="HedId"   value="<%=HedId%>">
<INPUT type=hidden name="MrSk"    value="<%=MrSk%>">
<INPUT type=hidden name="MaxW"    value="<%=MaxW%>">
<INPUT type=hidden name="WkCNo"     value="<%=WkCNo%>">
<INPUT type=hidden name="TruckerFlag" value="<%=TruckerFlag%>">
<INPUT type=hidden name="compFlag" value=<%=compFlag%>>
<% Else 'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige%>
  <INPUT type=hidden name="compFlag" value="false">
<% End If%>
</TABLE>
 <INPUT type=submit OnClick="GoNext()">
</FORM>
<!-------------画面終わり--------------------------->
<%Else%>
<!-------------エラー画面--------------------------->
<CENTER>
<FORM name="dmi015F" method="POST">
<DIV class=alert>
  <%If ret1=false Then%>
    <P>指定されたコンテナ「<%=strNum%>」は<BR>
       システムに登録されていません。<BR>
       入力の間違いがないか番号を確認してください。</P>
  <%Else%>
    <P>指定されたコンテナ「<%=strNum%>」は<BR>
       既に登録されています。</P>
  <%End If%>
</DIV>
<P><INPUT type=submit value="閉じる" onClick="window.close()"></P>
</FORM>
</CENTER>
<%End If%>
</BODY></HTML>
