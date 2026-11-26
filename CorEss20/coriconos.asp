<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


   response.expires=0
   sHTML = CorSetFHeaderEx()
   ' si estas en otra aplicacion como por ejemplo seguridad cambia
   ' siempre tiene que apuntar a EM
   Session("CorApli") = "COREM20"

   'response.write session("CMIBTTBI")
  'sHTML = sHTML & "<body bgcolor=" & session("CMFMENU")& ">"
   fTam="14"

   'sInput="<font color=#990000 size=" &Ftam &"> <i title='Inicio' class=""ttx la la-home""></i></font>"
   sInput="<img   title='Inicio' data-op='../coremanager/novedmenu.asp?TAB=view1' class=ttx style='width:60px ;border-radius: 10px;' src='../coremanager/IMGNOVEDADES/IconosMenu/inicio.jpg'>"
   Celda =ofv.GenCelda("","xx style=' padding-top: 80px;'","","","","","",sInput,"si")
   Fila = ofv.genrow("","","","","","","",Celda,"si")

   'sInput="<font color= #fff81f size=" &Ftam &"> <i title='Favoritos' class=""ttx la la-star""></i></font>"
   sInput="<img  data-op='../coremanager/novedmenu.asp?TAB=view2' title='Favoritos' class=ttx style='width:60px ;border-radius: 10px;' src='../coremanager/IMGNOVEDADES/IconosMenu/favoritos.jpg'>"
   Celda = ofv.GenCelda("","xxx  style=' padding-top: 30px;'","","","","","",sInput,"si")

   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

  ' sInput="<font size=" &Ftam &"> <i title='Suscripciones' class=""ttx la la-edit""></i></font>"
   sInput="<img data-op='../coremanager/novedmenu.asp?TAB=view3' title='Suscripciones' class=ttx  style='width:60px ;border-radius: 10px;' src='../coremanager/IMGNOVEDADES/IconosMenu/suscripcion.jpg'>"

   Celda =ofv.GenCelda("","vt style=' padding-top: 30px;'","","","","","",sInput,"si")
   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

 '  sInput="<font size=" &Ftam &"> <i title='Indice Norma' class=""ttx la la-list""></i></font>"
 '  sInput="<img  data-op='../coremanager/subconsulta.asp?MENU=PORTEMA&NOMBRE=TEMA&NDAT=ESTADO&NVAL=2&seg=CLIMS154172171124031051096140&MNN=POR_INDICE&base=D' title='Indice Norma' class=ttx  style='width:60px ;border-radius: 10px;' src='../coremanager/IMGNOVEDADES/IconosMenu/indice.jpg'>"

   sInput="<img  data-op='../coremanager/CONSTEMAIND.asp?FORMATO=GAL' title='Indice Norma' class=ttx  style='width:60px ;border-radius: 10px;' src='../coremanager/IMGNOVEDADES/IconosMenu/indice.jpg'>"

   Celda =ofv.GenCelda("","vt  style=' padding-top: 30px;'","","","","","",sInput,"si")
   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




  ' sInput="<font size=" &Ftam &"> <i title='Autorización' class=""ttx la la-files-o""></i></font>"
   sInput="<img  data-op='../coremanager/novedmenu.asp?TAB=view4' title='Autorización' class=ttx style='width:60px ;border-radius: 10px;' src='../coremanager/IMGNOVEDADES/IconosMenu/autorizaciones.jpg'>"

   Celda = ofv.GenCelda("","vt  style=' padding-top: 30px;'","","","","","",sInput,"si")

   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
   tabla = ofv.GenTabla("","","","","","0","0","0","0","","",Fila,"si")



  ' SDiv ="<div style=""position: absolute; -webkit-transform: translate(40%, 40%); transform: translate(40%, 40%);left:-40%"">" & tabla & "</div>"
  SDiv ="<center xxstyle=""position: absolute; -webkit-transform: translate(40%, 40%); transform: translate(40%, 40%);left:-40%"">" & tabla & "</center>"

  sAccion=""
  Formulario =  ofv.GenForm("dspnoved","",sAccion,"info","post","","","si")
  sHTML = sHTML & SDiv  & Formulario & "</body></html>"

 end if




Response.Write sHTML

%>