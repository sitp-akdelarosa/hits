<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "EXPORT", "expentry.asp"

	Dim strBookingNo
	strBookingNo = Request.QueryString("line")

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

	' 表示ファイルのレコードがある間繰り返す
    LineNo=0
    Do While Not ti.AtEndOfStream
        LineNo=LineNo+1
		If LineNo=CInt(strBookingNo) Then
			anyTmp=Split(ti.ReadLine,",")
		Else
			strDam = ti.ReadLine
		End If
	Loop

	ti.close()

    ' 輸出コンテナ照会リスト表示
'    WriteLog fs, "1011","ブッキング情報照会-搬出済コンテナ情報","00", anyTmp(1) & "/" & anyTmp(12) & ","	'D20040223
    WriteLog fs, "1011","ブッキング情報照会-搬出済コンテナ情報","00", anyTmp(1) & "/" & anyTmp(13) & ","	'I20040223

%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="JavaScript:this.focus()">
<!-------------ここから一覧画面--------------------------->
      <center>
      <table>
        <tr>
          <td align=center> 
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>ピックアップ済コンテナ情報</b></td>
              </tr>
            </table>
            <br>

<%
	If anyTmp(11)<>"0" Then
%>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td nowrap bgcolor="#000099"><font color="#ffffff"><b>Booking No.</td>
				</td>
                <td nowrap bgcolor="#ffffff"><%=anyTmp(1)%></td>
				</td>
			  </tr>
			</table>
			<BR>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>コンテナNo.</td>
                <td nowrap>TW(Kg)</td>
              </tr>
<!-- ここからデータ繰り返し -->
<%
'		contTmp=Split(anyTmp(12),"/")	'D20040223
		contTmp=Split(anyTmp(13),"/")	'I20040223
		MaxCont=UBound(contTmp)
		For k=0 To MaxCont
			lineTmp=Split(contTmp(k),"!")			' Ins 2005/03/28
%>
              <tr bgcolor="#FFFFFF"> 
				<td nowrap align=center valign=middle>
<% ' コンテナNo.
'	        If contTmp(k)<>"" Then								' Del 2005/03/28
'			    Response.Write contTmp(k)						' Del 2005/03/28
	        If lineTmp(0)<>"" Then								' Ins 2005/03/28
			    Response.Write lineTmp(0)						' Ins 2005/03/28
	        Else
	            Response.Write "<br>"
	        End If
%>
                </td>
				<td nowrap align=center valign=middle>
<% ' TW															' Ins 2005/03/28
	        If lineTmp(1)<>"" Then								' Ins 2005/03/28
			If lineTmp(1)="0" Then
			    Response.Write "－"
			ElseIf Len(lineTmp(1))<=2 Then
			    Response.Write lineTmp(1) & "00"
			Else
			    Response.Write lineTmp(1)
			End If
	        Else												' Ins 2005/03/28
	            Response.Write "－"
	        End If												' Ins 2005/03/28
%>
                </td>
              </tr>
<%
    	Next
%>
<!-- ここまで -->
            </table>
<% End If %>

      <form>
		<input type=button value="Close" onClick="JavaScript:window.close()">
      </form>
      </center>
    </td>
  </tr>
</table>

</body>
</html>

