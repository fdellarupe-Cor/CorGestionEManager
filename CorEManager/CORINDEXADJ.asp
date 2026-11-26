<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%
session("usuario")="GOOGLE"
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


Const adPersistXML = 1
Const adStateOpen = 1
Const adOpenStatic = 3
Const adLockBatchOptimistic = 4
Const adCmdText = 1

  set cn = ofv.conectar(ofv.strconn2)

  Set oRst = Server.CreateObject("ADODB.Recordset")




'  sSQL="select c.codobjeto,c.tipobjeto,c.directorio,b.descri COLLATE DATABASE_DEFAULT + ' / ' + c.descri COLLATE DATABASE_DEFAULT 	as descri,c.archivo from csrelpathdescriex a "

'  sSQL = sSQL & " inner join objetos b on (b.codobjeto=a.cobj) "
'  sSQL = sSQL & " inner join objprty b1 on (b1.codobjeto=b.codobjeto and b1.estado=3) "
'  sSQL = sSQL & " inner join objhtm c on (c.codobjeto=b1.codobjeto)"



ssql = " select  c.codobjeto,c.tipobjeto,c.directorio,b.descri COLLATE DATABASE_DEFAULT + ' / ' + c.descri COLLATE DATABASE_DEFAULT 	as descri,c.archivo   from csrelpathdescriex a "

 sSQL = sSQL & " inner join objetos b on (b.codobjeto=a.cobj)  "
 sSQL = sSQL & "inner join objprty b1 on (b1.codobjeto=b.codobjeto and b1.estado=3) "
 sSQL = sSQL & "inner join objhtm c on (c.codobjeto=b1.codobjeto)"
 sSQL = sSQL & "inner join  Tipobjeto d on( d.tipobjeto=b.tipobjeto)"
 sSQL = sSQL & "inner join  Objperfil e on( e.tipoperfil = d.perfil )"
 sSQL = sSQL & "left join  Objperfildoc f on (  b.codobjeto= f.codobjeto)"


 sSQL = sSQL & " and (cast( isnull(f.objobj3,e.objobj3)as int) <>0  or  cast(isnull(f.objhtm,e.objhtm)as int) <> 0)"


 set  oRst =cn.execute(ssql)

'response.write sSQL
 orst.Save response, adPersistXML

    oRst.Close
    Set oRst=nothing
   call ofv.cerrarconn(cn)

 END IF





%>

