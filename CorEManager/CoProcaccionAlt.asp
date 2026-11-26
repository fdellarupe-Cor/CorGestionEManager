<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  nAccion = request.querystring("nAccion")
  nProced = request.querystring("NID")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  bVersion = request.querystring("bVersion")
  baccion = request.querystring("baccion")
  bsec = request.querystring("bsec")

  set cn = ofv.conectar(ofv.strconn2)




' response.write "datos" & request.querystring & "<br>"




  wntitulo=" onload=" & chr(34) & "document.title='COR&nbsp;E-Manager&nbsp;Consulta&nbsp;de&nbsp;Procedimientos';"  & chr(34) & " "

  sHTML = ofv.menuheader("",wntitulo)

 shtml = shtml & CorPnlFnc("N","") & "<br><hr color='#000000' size=2 width=100% ><br>"




  strsql = "SELECT a.*,b.secuen from csaccaltsel A inner join objaccion B on (A.idaccion = B.codaccion) "
  strsql = strsql & "Where a.idaccion ='" & nAccion & "' and a.objid = " & baccion
  strsql = strsql & " and (a.modi = 0 or "
  strsql = strsql & " (a.modi <> 0 and a.tipmodi <> 'E')) "
  strsql = strsql & " order by a.secuencia"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  'response.write strsql
  strsql = "SELECT * FROM TIPOACCALT WHERE TIPOACCALT.modo='OP' Or TIPOACCALT.modo='RF' "
  strsql = strsql & "ORDER BY descri"

 ' sHTML = ofv.FormBody

'  sHTML = sHTML & ofv.GenTabla("panel","","","100% border=0 ","","0","0","-1","2","#ffffff","","","no")

sHTML = shtml & "<center>" & ofv.GenTabla("","","","90% style='border-style:solid;border-width:1;border-color:#dddddd;' ","","0","0","0","0","","","","no")

  if not rs.eof then
  do until rs.eof
    Valor = rs(0)

    ID = rs(2)
      sImage = "&nbsp;"
      if bVersion = "SI" and rs(9) then
      if rs(10) = "M" then
        sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
      elseif rs(10) = "A" then
        sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
      end if
      end if

    t1 = ofv.GenCelda("","dt","","20","","","",sImage,"si")

  simagen2 = ""

  strsql = "SELECT ArchivoCur from modo where tipomodo = '" & rs(18) & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs2.eof then simagen2 = ofv.GenImage("","","","","iconos/" & rs2(0),"1","")
  call ofv.cerrarconsulta(rs2)

    t1 = t1 & ofv.GenCelda("","dt","","","","","",sImagen2,"si")

    t1 = t1 & ofv.GenCelda("","dt","","","","","","<font color=#550000 face=verdana size=2><B>&nbsp;" & rs("secuen") & "." & rs(6) & "&nbsp;</b></font>","si")

    Vcont = Vcont + 1

     'Texto Accion Alt
      if rs(4) <> "" then
        Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs(4)
       ' response.write strsql
        set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsAcciones.eof then
           sAccion = rsAcciones(0)
        end if
        call ofv.cerrarconsulta(rsAcciones)
      end if

      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccion,i-1)
        sAcciont = mid(sAccion,i+1)
      else
        sAcciont=""
      end if
      sInput = "<i>" & sVerbo  & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont

    t2 = ofv.GenCelda("","vt","","","","","",sInput,"si")

    sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"si")

    'Texto de Referencia
    If BrowserName() = "IE" OR BrowserName() = "CHROME" then
      nCol = 60
    else
      nCol = 60
    end if

    t3 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t3 = t3 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t3 = t3 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sArma = rs(7)
    if len(sArma) > 0 then
       sArma = replace(sArma,chr(13),"<br>")
    end if
    SArma = "" & sArma
    t3 = t3 & ofv.GenCelda("","vt","","","","","",SArma,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t3,"si")
    rs.movenext
  loop


    'Encabezado Continuador



    t4 = ofv.GenCelda("","","","2%","","","","&nbsp;","si")
    t4 = t4 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t4 = t4 & ofv.GenCelda("","","","","","","","&nbsp;","si")
      'Texto Continuador
    if VCDACC <> "" then
      sVerbo = VCDACC
      sAcciont = VCDACC
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(VCDACC,i-1)
        sAcciont = mid(VCDACC,i+1)
      else
      sAcciont=""
      end if
      sInput = "<font color=#336699 face=verdana size=2><B><i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont & "</b></font>"
      sContinuador = VCSACC & " - " & sInput
    end if
    t4 = t4 & ofv.GenCelda("","dt","","","","","","Continua en: " & sContinuador,"si")
    if VCDACC <> "" then
       sHTML = sHTML & ofv.GenRow("","","","","","","",t4,"si")
    end if

else

    t6 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t6 = t6 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t6 = t6 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    SArma = "<font color=#225588 face=verdana size=2><h5><B>Sin&nbsp;datos&nbsp;disponibles</B></h5></font>"
    t6 = t6 & ofv.GenCelda("","dt","","","","","",SArma,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t6,"si")


end if

    t5 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t5,"si")


    sHTML = sHTML & ofv.GenFTabla()



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


end if


 ' sHTML = shtml & "</center>"
  sHTML = sHTML & "</body></html>"

response.write sHTML

%>