<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ms-ExpCom.inc"-->

<%
    ' ���O�C����ʂ̎擾�Ƃ��̏���
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"         '���j���[
        Response.End
    End If

    ' �Z�b�V�����̃`�F�b�N
    If strUserKind="�C��" Then
        CheckLogin "ms-expentry.asp?kind=1"
    ElseIf strUserKind="���^" Then
        CheckLogin "ms-expentry.asp?kind=2"
    Else
        CheckLogin "ms-expentry.asp?kind=3"
    End If

    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' �L�����Ă��錟�����������[�h
    strShipper=Session.Contents("findkey1")       '�׎�R�[�h
    strForwader=Session.Contents("findkey2")      '�C�݃R�[�h
    strTrucker=Session.Contents("findkey3")       '���^�R�[�h
    strVslCode=Session.Contents("findkey4")       '�D���R�[�h
    strVoyCtrl=Session.Contents("findkey5")       'Voyage No.

    ' �w������̎擾
    Dim strRequest
    strRequest = Request.QueryString("request")  ' �X�V���N�G�X�g���ID
    Dim strSortKey
    strSortKey = Request.QueryString("sort")     ' �\�[�g���[�h�̎擾
    If strSortKey="" Then
        strSortKey=Session.Contents("sortkey")
    End If
    Session.Contents("sortkey")=strSortKey

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"         '���j���[
        Response.End
    End If

    ' �f�[�^�x�[�X�̐ڑ�
    ConnectSvr conn, rsd

    ' ���������̍쐬
    sWhere = ""
    sSort = ""

    '�׎�R�[�h
    If strShipper<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Shipper='" & strShipper & "'"
            strOption = strOption & ",�׎�R�[�h," & strShipper
        Else
            sWhere = "ExportCargoInfo.Shipper='" & strShipper & "'"
            strOption = "�׎�R�[�h," & strShipper
        End If
    End If
    '�C�݃R�[�h
    If strForwader<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Forwarder='" & strForwader & "'"
            strOption = strOption & ",�C�݃R�[�h," & strForwader
        Else
            sWhere = "ExportCargoInfo.Forwarder='" & strForwader & "'"
            strOption = "�C�݃R�[�h," & strForwader
        End If
    End If
    '���^�R�[�h
    If strTrucker<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Trucker='" & strTrucker & "'"
            strOption = strOption & ",���^�R�[�h," & strTrucker
        Else
            sWhere = "ExportCargoInfo.Trucker='" & strTrucker & "'"
            strOption = "���^�R�[�h," & strTrucker
        End If
    End If
    '�D���R�[�h
    If strVslCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.VslCode='" & strVslCode & "'"
            strOption = strOption & ",�D���R�[�h," & strVslCode
        Else
            sWhere = sWhere & "ExportCargoInfo.VslCode='" & strVslCode & "'"
            strOption = "�D���R�[�h," & strVslCode
        End If
    End If
    'Voyage No.
    If strVoyCtrl<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
            strOption = strOption & ",Voyage No.," & strVoyCtrl
        Else
            sWhere = sWhere & "ExportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
            strOption = "Voyage No.," & strVoyCtrl
        End If
    End If

    ' Sort�����̍쐬
    strSortKey=Session.Contents("sortkey")
    If strSortKey="�׎喼" Then
        sSort="ExportCargoInfo.Shipper, ExportCargoInfo.ShipCtrl"
    ElseIf strSortKey="�C��" Then
        sSort="ExportCargoInfo.Forwarder"
    ElseIf strSortKey="�׎�Ǘ��ԍ�" Then
        sSort="ExportCargoInfo.ShipCtrl"
    ElseIf strSortKey="�q�ɓ���" Then
        sSort="ExportCargoInfo.WHArTime"
    ElseIf strSortKey="CY����" Then
        sSort="ExportCargoInfo.CYRecDate"
    ElseIf strSortKey="���^�Ǝ�" Then
'        sSort="mTrucker.FullName"
        sSort="ExportCargoInfo.Trucker"
    End If

    ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
    strFileName="./temp/" & strFileName
    ' �]���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSExpCntnr(conn, rsd, ti, sWhere, sSort)

    ' �t�@�C����DB�̃N���[�Y
    ti.Close
    conn.Close

    ' �ڍ׉�ʂ���̂Ƃ��A�Y���R���e�i�f�[�^�̍s����������
    If strRequest="ms-expdetail.asp" Then
        ' �L�����Ă��錟�����������[�h
        strFindCntnr=Session.Contents("dispexpctrl")     ' �\���׎�Ǘ��ԍ�

        ' �\���t�@�C����Open
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

        ' �ڍו\���s�̃f�[�^�̎擾
        bWriteFile = 0                    '�������ʃt���O
        LineNo=0
        Do While Not ti.AtEndOfStream
            anyTmp=Split(ti.ReadLine,",")
            LineNo=LineNo+1
            If anyTmp(14)=strFindCntnr Then
               bWriteFile=1
               Exit Do
            End If
        Loop

        ti.Close
    End If

    If bWriteFile = 0 Then
        ' �Y�����R�[�h�Ȃ��Ƃ�
        bError = true
        strError = "�w������ɊY������R���e�i�͂Ȃ��Ȃ�܂����B"
    End If

    ' �A�o�R���e�i�Ɖ�
'    WriteLog fs, "�A�o���Ɩ��x��-�A�o�R���e�i�Ɖ�", "��ʍX�V:SortKey," & strSortKey

    If bError Then
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������G���[���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2>
<%
    If strUserKind="�C��" Then
        Response.Write "<img src='gif/expkaika.gif' width='506' height='73'>"
    ElseIf strUserKind="���^" Then
        Response.Write "<img src='gif/exprikuun.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/expninushi.gif' width='506' height='73'>"
    End If
%>
          </td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
      <table>
        <tr> 
          <td nowrap>
            <dl> 
            <dt><font color="#000066" size="+1">�y�A�o�R���e�i���
<%
    If strUserKind="�C��" Then
        Response.Write "(�C�ݗp)"
    ElseIf strUserKind="���^" Then
        Response.Write "(���^�p)"
    Else
        Response.Write "(�׎�p)"
    End If
%>
               ��ʁz</font><br>
            <dd>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError
%>
            </dl>
          </td>
        </tr>
      </table>
      <form action="ms-expentry.asp">
        <br><br>
        <input type="submit" value=" ��  �� ">
      </form>
      </center>
      <br>
    </td>
  </tr>
  <tr>
    <td valign="bottom"> 
<%
    DispMenuBar
%>
    </td>
  </tr>
</table>
<!-------------�G���[��ʏI���--------------------------->
<%
    DispMenuBarBack "ms-expentry.asp"
%>
</body>
</html>

<%
    Else
        If strRequest="ms-expdetail.asp" Then
            ' �ڍ׉�ʂփ��_�C���N�g
            Response.Redirect "ms-expdetail.asp?line=" & LineNo  '�A�o�R���e�i�ڍ�
        Else
            ' �ꗗ��ʂփ��_�C���N�g
            Response.Redirect strRequest                         '�A�o�R���e�i�ꗗ
        End If
    End If
%>
