<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
''    CheckLogin "expentry.asp"

    '���͉�ʂ��L��
    Session.Contents("findcsv")="yes"    ' CSV�t�@�C�����͂ł��邱�Ƃ��L��

    ' �w������̎擾
    Dim strKind
    strKind = Request.QueryString("kind")

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' �Q�ƃ��[�h���Z�b�V�����ϐ��ɐݒ�
    If strKind="cntnr" Then
        Session.Contents("findkind")="Cntnr"
    Else
        Session.Contents("findkind")="Booking"
    End If

    ' �]���t�@�C���̎擾
    tb=Request.TotalBytes      :' �u���E�U����̃g�[�^���T�C�Y
    br=Request.BinaryRead(tb)  :' �u���E�U����̐��f�[�^

    ' BASP21 �R���|�[�l���g�̍쐬
    Set bsp=Server.CreateObject("basp21")

    filesize=bsp.FormFileSize(br,"csvfile")
    filename=bsp.FormFileName(br,"csvfile")

'    fpath=fs.GetFileName(filename)
    fpath=GetNumStr(Session.SessionID, 8) & "c.csv"
    fpath=fs.BuildPath(Server.MapPath("./temp"),fpath)

    lng=bsp.FormSaveAs(br,"csvfile",fpath)

    ' �t�@�C���]���Ɏ��s�����Ƃ�
    If lng<=0 Then
        bError=true
        strError = "'" & filename & "'�t�@�C���̓]���Ɏ��s���܂����B"
    Else
        Dim strCntnrNo()

        ' �]���t�@�C����Open
        Set ti=fs.OpenTextFile(fpath,1,True)

        iRecCount=0
        strFindKey=""
        ' �]���t�@�C���̃��R�[�h������ԌJ��Ԃ�
        Do While Not ti.AtEndOfStream
            cntnrNo = Trim(ti.ReadLine)
            If cntnrNo<>"" Then
                ReDim Preserve strCntnrNo(iRecCount)
                strCntnrNo(iRecCount) = UCase(cntnrNo)
                If strFindKey<>"" Then
                    strFindKey=strFindKey & "," & strCntnrNo(iRecCount)
                Else
                    strFindKey=strCntnrNo(iRecCount)
                End If
                iRecCount=iRecCount + 1
            End If
        Loop
        ti.Close
        Session.Contents("findkey")=strFindKey     ' �Q��Key���L��
        ' �]���t�@�C���̍폜
        fs.DeleteFile fpath

        ' �R���e�i��񃌃R�[�h�̎擾
        ConnectSvr conn, rsd

        ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
        strFileName="./temp/" & strFileName
        ' �e���|�����t�@�C����Open
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
        bWriteFile = 0

        For iCount=0 To iRecCount - 1
            If strKind="cntnr" Then
                sWhere = "ExportCont.ContNo='" & strCntnrNo(iCount) & "'"
            Else
                sWhere = "ExportCont.BookNo='" & strCntnrNo(iCount) & "'"
            End If

            bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
        Next

        ti.Close
        conn.Close

        If bWriteFile = 0 Then
            ' �Y�����R�[�h�Ȃ��Ƃ�
            bError = true
            strError = "�w������ɊY������R���e�i�͂���܂���ł����B"
        End If
    End If

    ' Temp�t�@�C�������ݒ�
    SetTempFile "EXPORT"

    strOption = filename

    If bError Then
        strOption = strOption & "," & "���͓��e�̐���:1(���)"
    Else
        strOption = strOption & "," & "���͓��e�̐���:0(������)"
    End If

	Dim iWrkNum
    If strKind="cntnr" Then
		iWrkNum = 21
	Else
		iWrkNum = 22
	End If
    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "1003","�A�o�R���e�i�Ɖ�-CSV�t�@�C���]��",iWrkNum, strOption

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
          <td rowspan=2><img src="../gif/expentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
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
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>CSV�t�@�C���]��</b></td>
          <td><img src="../gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr> 
          <td nowrap>
            <font color="#000066" size="+1">�y�R���e�i���Ɖ�p�t�@�C���]����ʁz</font>
			<BR><br>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError
%>
          </td>
        </tr>
      </table>
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
    DispMenuBarBack "expcsv.asp"
%>
</body>
</html>

<%
    Else
        If bWriteFile = 1 Then
            '�߂��ʎ�ʂ��L��
            Session.Contents("dispreturn")=0
            ' �ڍ׉�ʂփ��_�C���N�g
            Response.Redirect "expdetail.asp?line=1"    '�A�o�R���e�i�ڍ�
        Else
            '�߂��ʎ�ʂ��L��
            Session.Contents("dispreturn")=0
            ' �ꗗ��ʂփ��_�C���N�g
            Response.Redirect "explist.asp"             '�A�o�R���e�i�ꗗ
        End If
    End If
%>
