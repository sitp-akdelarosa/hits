<%@Language="VBScript" %><!--#include file="Common.inc"--><%    ' Tempファイル属性のチェック    CheckTempFile "IMPORT", "impentry.asp"    ' 表示モードの取得    Dim bDispMode          ' true=コンテナ検索 / false=BL検索    If Session.Contents("findkind")="Cntnr" Then        bDispMode = true    Else        bDispMode = false    End If    ' File System Object の生成    Set fs=Server.CreateObject("Scripting.FileSystemobject")    ' 表示ファイルの取得    Dim strFileName    strFileName = Session.Contents("tempfile")    If strFileName="" Then        ' セッションが切れているとき        Response.Redirect "impentry.asp"             '輸入コンテナ照会トップ        Response.End    End If    strFileName="./temp/" & strFileName    ' 輸入コンテナ照会リスト表示    WriteLog fs, "2004","輸入コンテナ照会-搬入までの位置情報","00", ","    ' 表示ファイルのOpen    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)    '戻り画面種別を記憶    Session.Contents("dispreturn")=1%><html><head><title></title><meta http-equiv="Pragma" content="no-cache"><meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS"><SCRIPT Language="JavaScript"><%    DispMenuJava%></SCRIPT></head><body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><!-------------ここから一覧画面---------------------------><table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>  <tr>    <td valign=top>      <table width="100%" border="0" cellspacing="0" cellpadding="0">        <tr>          <td rowspan=2><img src="../gif/implistt.gif" width="506" height="73"></td>          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>        </tr>        <tr>          <td align="right" width="100%" height="48"> <%' Added and Commented by seiko-denki 2003.07.18	DisplayCodeListButton'    DispMenu'	Dim strScriptName,strRoute'	strScriptName = Request.ServerVariables("SCRIPT_NAME")'	strRoute = SetRoute(strScriptName)'	Session.Contents("route") = strRoute' End of Addition by seiko-denki 2003.07.18%>          </td>        </tr>      </table>      <center><!-- commented by seiko-denki 2003.07.18		<table width=95% cellpadding="0" cellspacing="0" border="0">		  <tr>			<td align="right">			  <font color="#333333" size="-1">				<%=strRoute%>			  </font>			</td>		  </tr>		</table>End of comment by seiko-denki 2003.07.18 -->		<table width=95% cellpadding=3>			<tr>				<td align=right>					<font color="#224599">					&nbsp;&nbsp;<%=GetUpdateTime(fs)%>					</font>				</td>			</tr>		</table>      <table>        <tr>          <td>             <table>              <tr>                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>                <td nowrap><b>ターミナル搬入までの位置情報&nbsp;</b></td>                <td><img src="../gif/hr.gif"></td>              </tr>            </table>            <br>        <table border="0" cellspacing="2" cellpadding="1">          <tr>             <td width="15"><BR></td>            <td><font color="#000000" size="-1">（※1) クリックで単独コンテナ情報を表示</font></td>            <td width="15"><BR></td>            <td><font color="#000000" size="-1">（※2）仕出港の時刻は、現地時間です。</font></td>          </tr>        </table>            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">              <tr align="center" bgcolor="#FFCC33"> <%    If Not bDispMode Then        Response.Write "<td nowrap rowspan='3'>BL "        Response.Write "No.</td>"    End If%>                <td nowrap rowspan="3">コンテナNo.<font size="-1"><sup>(※1)</sup></font></td><!-- mod by nics 2009.02.24 --><!--                <td nowrap colspan="2">本船</td>-->                <td nowrap colspan="1">本船</td><!-- end of mod by nics 2009.02.24 -->                <td nowrap bgcolor="#FFCC33">仕出港</td><!-- mod by nics 2009.02.24 --><!--                <td nowrap colspan="7">ターミナル</td>-->                <td nowrap colspan="8">ターミナル</td><!-- end of mod by nics 2009.02.24 -->              </tr>              <tr bgcolor="#FFCC33" align="center">                 <td nowrap rowspan="2" bgcolor="#FFFF99">船名</td><!-- commented by nics 2009.02.24                <td nowrap rowspan="2" bgcolor="#FFFF99">仕出港名</td>end of comment by nics 2009.02.24 -->                <td nowrap rowspan="2" bgcolor="#FFFF99">離岸完了<br>                  時刻<font size="-1"><sup>(※2)</sup></font></td><!-- mod by nics 2009.02.24 --><!--                <td nowrap colspan="3" bgcolor="#FFFF99">着岸時刻</td>-->                <td nowrap colspan="2" bgcolor="#FFFF99">着岸時刻</td><!-- end of mod by nics 2009.02.24 -->                <td nowrap colspan="2" bgcolor="#FFFF99">搬入確認時刻 </td>                <td nowrap rowspan="2" bgcolor="#FFFF99">搬出可否</td>                <td nowrap rowspan="2" bgcolor="#FFFF99">ヤード搬出<br>完了時刻</td><!-- add by nics 2009.02.24 -->                <td nowrap rowspan="2" bgcolor="#FFFF99"><font color="#000000">搬出ターミナル<br>(蔵置場所コード)</font></td>                <td nowrap rowspan="2" bgcolor="#FFFF99"><font color="#000000">本船担当<br>オペレータ</font></td><!-- end of add by nics 2009.02.24 -->              </tr>              <tr bgcolor="#FFCC33" align="center"> <!-- commented by nics 2009.02.24                <td nowrap bgcolor="#FFFF99">計画</td>end of comment by nics 2009.02.24 -->                <td nowrap bgcolor="#FFFF99">予定</td>                <td nowrap bgcolor="#FFFF99">完了</td>                <td nowrap bgcolor="#FFFF99">予定</td>                <td nowrap bgcolor="#FFFF99">完了</td>              </tr><!-- ここからデータ繰り返し --><% ' 表示ファイルのレコードがある間繰り返す    LineNo=0    Do While Not ti.AtEndOfStream        anyTmp=Split(ti.ReadLine,",")        LineNo=LineNo+1%>              <tr bgcolor="#FFFFFF"><% ' BL No    If Not bDispMode Then        Response.Write "<td nowrap align=center valign=middle>"        If strBooking<>anyTmp(0) Then            Response.Write anyTmp(0)            strBooking=anyTmp(0)        Else            Response.Write "<br>"        End If        Response.Write "</td>"    End If%>                <td nowrap align=center valign=middle><% ' コンテナNo.    Response.Write "<a href='impdetail.asp?line=" & LineNo & "&return=1'>" & anyTmp(1) & "</a>"%>                </td>                <td nowrap align=center valign=middle><% ' 船名    If anyTmp(7)<>"" Then        Response.Write anyTmp(7)    Else        Response.Write "<br>"    End If%>                </td><!-- commented by nics 2009.02.24                <td nowrap align=center valign=middle><% ' 仕出港    If anyTmp(9)<>"" Then        Response.Write anyTmp(9)    Else        Response.Write "<br>"    End If%>                </td>end of comment by nics 2009.02.24 -->                <td nowrap align=center valign=middle><% ' 仕出港 - 離岸完了    Response.Write DispDateTimeCell(anyTmp(11),10)%>                </td><!-- commented by nics 2009.02.24                <td nowrap align=center valign=middle><% ' ターミナル − 着岸スケジュール    If anyTmp(31)<>"" Then        Response.Write "<font color='#0000FF'>"    End If    Response.Write DispDateTimeCell(anyTmp(31),10)    If anyTmp(31)<>"" Then        Response.Write "</font>"    End If%>                </td>end of comment by nics 2009.02.24 -->                <td nowrap align=center valign=middle><% ' ターミナル - 着岸予定    If anyTmp(2)<>"" Then        bLate = false        If anyTmp(3)<>"" Then            If anyTmp(2)<anyTmp(3) Then                bLate = true            End If        End If        If anyTmp(31)<>"" Then            If anyTmp(31)<anyTmp(2) Then                bLate = true            End If        End If        If bLate Then            Response.Write "<font color='#FF0000'>"        Else            Response.Write "<font color='#0000FF'>"        End If        Response.Write DispDateTimeCell(anyTmp(2),10)        Response.Write "</font>"    Else        Response.Write DispDateTimeCell(anyTmp(2),10)    End If%>                </td>                <td nowrap align=center valign=middle><% ' ターミナル - 着岸完了    Response.Write DispDateTimeCell(anyTmp(3),10)%>                </td>                <td nowrap align=center valign=middle><% ' ターミナル - 搬入確認予定    If anyTmp(32)<>"" Then        If anyTmp(18)<>"" Then            If anyTmp(32)<anyTmp(18) Then                Response.Write "<font color='#FF0000'>"            Else                Response.Write "<font color='#0000FF'>"            End If        Else            Response.Write "<font color='#0000FF'>"        End If        Response.Write DispDateTimeCell(anyTmp(32),10)        Response.Write "</font>"    Else        Response.Write DispDateTimeCell(anyTmp(32),10)    End If%>                </td>                <td nowrap align=center valign=middle><% ' ターミナル - ヤード搬入(確認)完了    Response.Write DispDateTimeCell(anyTmp(18),5)%>                </td>                <td nowrap align=center valign=middle><% ' ターミナル搬出可否    If anyTmp(4)="Y" Then        Response.Write "○"    ElseIf anyTmp(4)="S" Then        Response.Write "済"    Else        Response.Write "×"    End If%>                </td>                <td nowrap align=center valign=middle><% ' ターミナル - ヤード搬出完了    Response.Write DispDateTimeCell(anyTmp(13),10)%>                </td><!-- add by nics 2009.02.24 -->                     <td nowrap align=center valign=middle><% ' 搬出ターミナル(蔵置場所コード)    strDisp = "<br>"    If anyTmp(5) <> "" Then        strDisp = anyTmp(5)        If anyTmp(43) <> "" Then            strDisp = strDisp & "<br>(" & anyTmp(43) & ")"        End If    End If    Response.Write strDisp%>                     </td>                     <td nowrap align=center valign=middle><% ' 担当オペレータ    If anyTmp(45)<>"" Then        Response.Write anyTmp(45)    Else        Response.Write "<br>"    End If%>                     </td><!-- end of add by nics 2009.02.24 -->              </tr><%    Loop%><!-- ここまで -->            </table><form>      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='impreload.asp?request=implist1.asp'"></form>          </td>        </tr>      </table>      </center>    </td>  </tr>  <tr>    <td valign="bottom"><%    DispMenuBar%>    </td>  </tr></table><!-------------一覧画面終わり---------------------------><%    DispMenuBarBack "implist.asp"%></body></html>