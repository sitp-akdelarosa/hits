<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "sokuji.asp"

    ' ���[�U��ނ��擾����
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"             '�g�b�v
        Response.End
    End If

    ' ���[�U��ނɂ���ʂ�I��
    If strUserKind="�C��" Then
        Response.Redirect "sokuji-kaika-updtchk.asp"
    Else
        Response.Redirect "sokuji-koun-updtchk.asp"
    End If
    Response.End
%>
