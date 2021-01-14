<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "IMPORT", "impentry.asp"

    ' 指定引数の取得
    Dim iLineNo
    iLineNo = Request.QueryString("line")

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "expentry.asp"             '輸出コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 詳細表示行のデータの取得
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLineNo=LineNo Then
           Exit Do
        End If
    Loop
    ti.Close

    ' 輸入コンテナ照会
    WriteLog fs, "2007","輸入コンテナ照会-単独コンテナ保税輸送期間","00", anyTmp(1) & ","
%>

<html>
<head>
<title>保税輸送期間</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body bgcolor="#E6E8FF" text="#000000">
  <center>
<!-----保税輸送期間---------------->
  <table>
    <tr>
      <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
      <td nowrap><b>保税輸送期間</b></td>
    </tr>
  </table>
  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
    <tr align="center" bgcolor="#FFCC33"> 
      <td nowrap>FROM</td>
      <td nowrap>TO</td>
    </tr>
    <tr align="center"> 
      <td align=center>
<% ' 保税輸送(From)
    Response.Write DispDateTimeCell(anyTmp(28),5)
%>
      </td>
      <td align=center>
<% ' 保税輸送(To)
    Response.Write DispDateTimeCell(anyTmp(29),5)
%>
      </td>
    </tr>
  </table>
  <FORM>
    <INPUT type="button" value=" Close " onClick="opener.winfl=0;window.close()">
  </FORM>
  </center>
</body>
</html>
