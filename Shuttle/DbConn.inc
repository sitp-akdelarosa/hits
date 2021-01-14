<%

'データベースのコネクション作成用インクルードファイル


Dim conn					'コネクションオブジェクトを格納する


Set conn = Server.CreateObject( "ADODB.Connection" )
conn.Open "HakataDB", "sa", ""


%>
