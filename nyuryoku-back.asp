<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-te.asp"

	' userid�̃Z�b�V��������ɂ���
	Session.Contents("userid") = ""

    ' ���[�U�[ID���͉�ʂ�
    Response.Redirect "userchk.asp?link=nyuryoku-te.asp"
%>
