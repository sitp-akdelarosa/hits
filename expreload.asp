<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
''    CheckLogin "expentry.asp"

    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "expentry.asp"

    ' �L�����Ă��錟�����������[�h
    strFindKind=Session.Contents("findkind")     ' ��������
    strFindCSV=Session.Contents("findcsv")       ' �������
    strFindKey=Session.Contents("findkey")       ' �����L�[

    ' �w������̎擾
    Dim strRequest
    strRequest = Request.QueryString("request")  ' �X�V���N�G�X�g���ID

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "expentry.asp"         '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' �f�[�^�x�[�X�̐ڑ�
    ConnectSvr conn, rsd

    ' �]���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

'������ Del_S  by nics 200902����
'    ' ������ʂ̔���
'    If strFindCSV="no" Then
'        ' ��ʓ��͂̂Ƃ�
'        sWhere = ""
'        If strFindKind="Booking" Then            ' Booking�ԍ������̂Ƃ�
'            iCanma = InStr(strFindKey,",")
'            Do While iCanma>0
'                sTemp = Trim(Left(strFindKey,iCanma-1))
'                strFindKey = Right(strFindKey,Len(strFindKey)-iCanma)
'                If sWhere<>"" Then
'                    sWhere = sWhere & " Or ExportCont.BookNo='" & sTemp & "'"
'                Else
'                    sWhere = "ExportCont.BookNo='" & sTemp & "'"
'                End If
'                iCanma = InStr(strFindKey,",")
'            Loop
'            If sWhere<>"" Then
'                sWhere = sWhere & " Or ExportCont.BookNo='" & Trim(strFindKey) & "'"
'            Else
'                sWhere = "ExportCont.BookNo='" & Trim(strFindKey) & "'"
'            End If
'        Else                                     ' Container�ԍ������̂Ƃ�
'            iCanma = InStr(strFindKey,",")
'            Do While iCanma>0
'                sTemp = Trim(Left(strFindKey,iCanma-1))
'                strFindKey = Right(strFindKey,Len(strFindKey)-iCanma)
'                If sWhere<>"" Then
'                    sWhere = sWhere & " Or ExportCont.ContNo='" & sTemp & "'"
'                Else
'                    sWhere = "ExportCont.ContNo='" & sTemp & "'"
'                End If
'                iCanma = InStr(strFindKey,",")
'            Loop
'            If sWhere<>"" Then
'                sWhere = sWhere & " Or ExportCont.ContNo='" & Trim(strFindKey) & "'"
'            Else
'                sWhere = "ExportCont.ContNo='" & Trim(strFindKey) & "'"
'            End If
'        End If
'
'        bWriteFile = SerchExpCntnr(conn, rsd, ti, sWhere)
'
'    Else
'������ Del_E  by nics 200902����
        ' �����L�[�̕���
        strCntnrNo=Split(strFindKey, ",")
        iRecCount=Ubound(strCntnrNo)+1

        bWriteFile = 0

        For iCount=0 To iRecCount - 1
            If strFindKind="Cntnr" Then
                sWhere = "ExportCont.ContNo='" & strCntnrNo(iCount) & "'"
            Else
                sWhere = "ExportCont.BookNo='" & strCntnrNo(iCount) & "'"
            End If

            bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
        Next

'������ Del_S  by nics 200902����
'    End If
'������ Del_E  by nics 200902����

    ' �t�@�C����DB�̃N���[�Y
    ti.Close
    conn.Close

    ' �ڍ׉�ʂ���̂Ƃ��A�Y���R���e�i�f�[�^�̍s����������
    If strRequest="expdetail.asp" Then
        ' �L�����Ă��錟�����������[�h
        strFindCntnr=Session.Contents("dispcntnr")     ' �\���R���e�iNo.

        ' �\���t�@�C����Open
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

        ' �ڍו\���s�̃f�[�^�̎擾
        bWriteFile = 0                    '�������ʃt���O
        LineNo=0
        Do While Not ti.AtEndOfStream
            anyTmp=Split(ti.ReadLine,",")
            LineNo=LineNo+1
            If anyTmp(1)=strFindCntnr Then
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
'    WriteLog fs, "�A�o�R���e�i�Ɖ�", "��ʍX�V"

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
          <td rowspan=2><img src="gif/csvt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
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
          <td nowrap>
            <dl> 
            <dt><font color="#000066" size="+1">�y�R���e�i���Ɖ�p�t�@�C���]����ʁz</font><br>
            <dd>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError
%>
            </dl>
          </td>
        </tr>
      </table>
      <form action="expentry.asp">
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
    DispMenuBarBack "expentry.asp"
%>
</body>
</html>

<%
    Else
        If strRequest="expdetail.asp" Then
            ' �ڍ׉�ʂփ��_�C���N�g
            Response.Redirect "expdetail.asp?line=" & LineNo  '�A�o�R���e�i�ڍ�
        Else
            ' �ꗗ��ʂփ��_�C���N�g
            Response.Redirect strRequest                      '�A�o�R���e�i�ꗗ
        End If
    End If
%>