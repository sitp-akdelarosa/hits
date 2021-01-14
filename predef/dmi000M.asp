<%
	@LANGUAGE = VBScript
	@CODEPAGE = 932
%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi000M.asp				_/
'_/	Function	:事前情報一覧画面メニュー		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th   2004/01/31	3次対応		_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
	
	if Request("logout") = "logout" then
		Session.Contents("userid") = ""
		response.redirect("../userchk.asp?link=predef/dmi000F.asp")
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
//遷移
function opnewin(i){
  Fname=document.dmi000MF;
  Fname.target="List";
  switch(i){
    case 1:
        Fname.action="./dmo010F.asp";
        break;
    case 2:
        Fname.action="./dmo110F.asp";
        break;
    case 3:
        Fname.action="./dmo210F.asp";
        break;
    case 4:
        Fname.action="./dmo310F.asp";
        break;
    case 5:
      	Win = window.open('dmi010.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
    case 6:
      	Win = window.open('dmi110.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
    case 7:
      	Win = window.open('dmi210.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
    case 8:
      	Win = window.open('dmi310.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
	case 9:
        w=900;
        h=600;
        if(screen.width){
            l=(screen.width-w)/2;
        }
        if(screen.availWidth){
            l=(screen.availWidth-w)/2;
        }
        if(screen.height){
            t=(screen.height-h)/2;
        }
        if(screen.availHeight){
            t=(screen.availHeight-h)/2;
        }
    	Win = window.open("dmi410.asp", "FConIn", "width="+w+",height=" + h +",top="+t+",left="+l+",resizable=yes,scrollbars=no");
    	break;
	case 10:
        Fname.action="./top.asp";
        break;
    //Y.TAKAKUWA Upd-S 2013-02-14
    case 11:
        Fname.action="./dml000A.asp";
        break;
    //Y.TAKAKUWA Upd-E 2013-02-14
  }
  //Y.TAKAKUWA Upd-S 2013-02-18
  if(i<5 || i == 10 || i == 11){
    Fname.submit();
  }
  //Y.TAKAKUWA Upd-E 2013-02-18
}
function flogout(){
  Fname=document.dmi000MF;
  Fname.target="_top";
  Fname.logout.value = "logout";
  Fname.action = "./dmi000M.asp";
  Fname.submit();
}
-->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY class="menu">
<!-------------事前入力初期画面Menu--------------------------->
<CENTER>
<P><B><Font color="#000066">事前情報</FONT></B></P>
<P><A HREF="JavaScript:opnewin(10)">作業<BR>テーブル</A></P>
<P><BR></P>
<P><B><Font color="#000066">各種一覧画面</FONT></B></P>
<P><A HREF="JavaScript:opnewin(1)">実搬出</A></P>
<P><A HREF="JavaScript:opnewin(2)">空搬入</A></P>
<P><A HREF="JavaScript:opnewin(3)">空バンピック</A></P>
<P><A HREF="JavaScript:opnewin(4)">搬入票作成</A></P>
<P><B><Font color="#000066">各種入力画面</FONT></B></P>
<% If Session.Contents("UType") = 3 Then 
     Response.Write "<P>実搬出</P>"
   Else
     Response.Write "<P><A HREF='JavaScript:opnewin(5)'>実搬出</A></P>"
   End If %>
<P><A HREF="JavaScript:opnewin(6)">空搬入</A></P>
<P><A HREF="JavaScript:opnewin(7)">空バンピック</A></P>
<P><A HREF="JavaScript:opnewin(8)">搬入票作成</A></P>
<P><A HREF="JavaScript:opnewin(9)">作業発生<BR>mail設定</A></P>
<!--Y.TAKAKUWA Add-S 2013-02-14-->
<%
'コンテナロックINIファイル取込。
Function getContainerLockINI(strUser)
  dim ObjFSO, ObjTS, tmpStr
  dim tmpUser
  dim icnt
  dim iFlag
  getContainerLockINI = false
  Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists(Server.Mappath("./INI/CONTAINERLOCK.INI")) = true then
    Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./INI/CONTAINERLOCK.INI"),1,false)
    Do Until ObjTS.AtEndofStream
    tmpStr = Split(ObjTS.ReadLine,"=", 3, 1)
    if UBound(tmpStr) < 0 then
      getContainerLockINI = true
      exit function
    else
      Select Case tmpStr(0)
        Case "ALLOWEDUSER"
          getContainerLockINI = false
          tmpUser = Split(UCase(tmpStr(1)),",")
          If Ubound(tmpUser) < 0 then
            getContainerLockINI = true
          else
           if Ubound(tmpUser) = 0 then
              if Trim(tmpUser(0)) = "" Then
                getContainerLockINI = true
              end if 
           end if
           For icnt = 0 to Ubound(tmpUser)
            If Trim(tmpUser(icnt)) = Trim(UCase(strUser)) Then
              getContainerLockINI = true
              Exit For
            End If
           Next
          end if
          Exit Function
        Case Else
           getContainerLockINI = true
      End Select
    end if
    Loop
    ObjTS.Close
    Set ObjTS = Nothing
  else
    getContainerLockINI = true
  end if
  Set ObjFSO = Nothing
End Function
%>
<%if getContainerLockINI(UCase(Session.Contents("userid"))) = true then %>
<P><A HREF="JavaScript:opnewin(11)">コンテナ<BR/>ロック</A></P>
<%end if%>
<P><A HREF="JavaScript:flogout();">ログアウト</A></P>

<!--Y.TAKAKUWA Add-E 2013-02-14-->
<FORM name="dmi000MF">
<input type=hidden name="logout" value="" >
</FORM>
</CENTER>
<!-------------画面終わり--------------------------->
</BODY></HTML>
