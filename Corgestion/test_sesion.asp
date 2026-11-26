<%@ Language=VBScript %>
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
Response.Write "SessionID=" & Session.SessionID & "<br>"
Session("contador") = CLng("0" & Session("contador")) + 1
Response.Write "Contador de sesion=" & Session("contador") & "<br>"
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.Open scon1
Response.Write "Base=" & Cnn.Execute("SELECT DB_NAME()")(0) & "<br>"
Response.Write "ANTES: session(usuario)=[" & Session("usuario") & "] session(func)=[" & Session("func") & "]<br>"
Set rs = Cnn.Execute("sp_cg_Responsables @Modo='seguridad', @IdSeg=N'" & Replace(UserLogon(), "'", "''") & "'")
If Not rs.EOF Then
    Session("func") = rs(0): Session("usuario") = rs(1)
Else
    Session("func") = "FUNC": Session("usuario") = "0"
End If
Response.Write "DESPUES: session(usuario)=[" & Session("usuario") & "] session(func)=[" & Session("func") & "]<br>"
For Each t In Array("ALTA", "MODIF", "ELIM", "CONS", "EST")
    v = getPermisosAtrm("CG_INCIDENTES", t)
    Response.Write "getPermisosAtrm(CG_INCIDENTES, " & t & ") = [" & v & "] tipo=" & TypeName(v) & "<br>"
Next
Response.Write "gAtrmOrigen=[" & gAtrmOrigen & "] CG_ATRM_APLI=[" & CG_ATRM_APLI & "]<br>"
%>