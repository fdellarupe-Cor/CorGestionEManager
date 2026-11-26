<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sdocum = request.querystring("sdocum")
  sKIND = request.querystring("SKIND")

  IF SKIND <> "" THEN
     SESSION("SKIND") = SKIND
  ELSE
     SKIND = SESSION("SKIND")
  END IF 


  ''   SKIND = SESSION("SKIND")

  sft   = "&nbsp;<b>"
  sfta  = "&nbsp;<b>"
  sfidc = "<font face=corflaticon size=2 >0</font>" 
  sfida = "<font face=corflaticon size=2 >1</font>" 
  sfitc = "<font face=corflaticon size=2 >0</font>" 
  sfita = "<font face=corflaticon size=2 >1</font>" 
  sftc  = "</b>" 
  
  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_sdocum=" & sdocum

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)


  set cn = ofv.conectar(ofv.strconn2)


  strsqlx = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1) 
  end if
  call ofv.cerrarconsulta(rs)


    sarch = "NO"

    Strsql = "SELECT descri, archivo, objid, directorio "
    Strsql = Strsql & " from ObjHtm Where codobjeto = " & nCodigo & " AND TIPO = " & SKIND 
    Strsql = Strsql & " order by secuen "
  Formulario = ""
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof

    sdirec = rs(3) & ""

    if sdirec = "" then sdirec = "Archivos"
        


    if sdocum = cstr(rs(2)) then
       sarch = sdirec & "/" & rs(1)
       clase = "<font color=#993333><u>" 
       stexto = sfida & sfta & replace(rs(0)," ","&nbsp;") & sftc
    else 
       clase = "" 
       stexto = sfidc & sft & replace(rs(0)," ","&nbsp;")
    end if 
    stexto1 = "<A class=ttt  href=CoArchCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(2)
    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"

    if (rs(1) & "") = "" then
       stexto1 = "<A class=ttt  href='javascript:void(0);' "
       stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"
    end if

       Celda = ""


      ' estado observacion
      
      Celdaobs  = ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")

  

        
      if Nestado=5  then 

      	     '  sAccion = "genObs.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	     '  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(2)  & "&ID2=" & rs(2)
      
      	     '  xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      	     '  sHref   = sAccion
      	     '  sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	     '  Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Observacion de " & BtipP & bpespec & "' ","","1%","1","","",sInput,"si")
             '  Celdaobs  = ofv.GenLink("",xclas & " title='Obs de " & BtipP & bpespec & "' ",sHref,"","","",Celdaobs,"si")


               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(2)  & "&ID2=" & rs(2)
      
      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")



      end if


      Celda = Celda & Celdaobs


       Celda = Celda & ofv.GenCelda("","ut  ","","95%","","","",stexto1,"si")
       Formulario = Celda
       Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")

    rs.movenext
    loop

  sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ","","90%","","0","0","4","0","","",Fila,"si")
 '' if not sdocum = "" then  
  if sarch <> "NO" then
     sAccion = "CoArchivos3.asp?Archivo=" & sarch
     sHTML = sHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
  end if

  else

    sdocus = "NO"


  end if

'''SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp

    fila = ""

    sarch = 0
    strsql = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
    strsql = strsql & "COMPLEMENTOS.descri, CSACOMPLE.nombre FROM COMPLEMENTOS, "
    strsql = strsql & " OBJACCION, CSACOMPLE where COMPLEMENTOS.codcomp = OBJACCION.codcomp "
    strsql = strsql & " and COMPLEMENTOS.codcomp = CSACOMPLE.codcomp "
    strsql = strsql & " and OBJACCION.codobjeto =" & ncodigo
    strsql = strsql & " ORDER BY COMPLEMENTOS.codcomp "
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
  Formulario = ""
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof


   if sdocum = cstr(rs(0)) then
       sarch = rs(0)
       clase = "<font color=#993333><u>" 
       stexto = sfida & sfta & replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2)," ","&nbsp;") & sftc
    else 
       clase = "" 
       stexto = sfidc & sft & replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2)," ","&nbsp;")
    end if 
    stexto1 = "<A class=ttt  href=CoArchCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(0)
    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"


       Celda = ofv.GenCelda("","ut  ","","1%","","","",stexto1,"si")
       Formulario = Celda

       rs.movenext

       Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")
    loop

    sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ","","90%","","0","0","4","0","","",Fila,"si")

    if  sarch > 0 then  
        sAccion = "EditarImagen2.asp?ID=" & sarch
        sHTML   = sHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
    end if

  else
    scomple = "NO"
  end if 
    Call OFV.cerrarconsulta(rs)


  if sdocum <> "" then
     sHTML   = sHTML & "<script language=javascript >"
     sHTML   = sHTML & " document.cons.submit(); "
     sHTML   = sHTML & "</script>"
  end if




  if scomple = "NO" and sdocus = "NO" then
    sestr = "<h5>Sin&nbsp;datos&nbsp;disponibles"
    sestr = sestr & "</h5>" 
    Formulario = ofv.GenCelda("","ttt","center","80%","","","",sestr,"si")
    Fila = ofv.GenRow("","","","","","","",Formulario,"si")
    sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-COLOR:#dddddd;' ","","90%","","0","0","4","0","","",Fila,"si")
  end if

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & GTCOBS()


  sHTML = sHTML & "</body></html>"
  response.write sHTML

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




end if



%>