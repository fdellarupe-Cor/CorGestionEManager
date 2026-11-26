<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%
session("usuario")="GOOGLE"
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

response.expires = 0

sModal  = request.querystring("modal")
sTipo   = request.querystring("Tipo")
sNombre = request.querystring("Nombre")
sNombre = ofv.convertircar(sNombre,"_"," ")
sMenu   = request.querystring("Menu")
sMETODO   = request.querystring("METODO")




session("ModMain") = "ED"


xapli = 1

  on error resume next



''''''''''''''''''''''''''''

  session("xorigen") = "tema"

  XXXX = NEWEMPRE("NO")

  sEMPRE = request.querystring("EMPRE")

  IF SEMPRE = "" THEN
     SEMPRE = ATREMP(0,0)
  ELSE
     recarga = NEWEMPRE(SEMPRE)
  END IF


  sUN = SEMPRE

  if sUN = "" then sUN = 1
  session("UN") = sUN




 ' sHTML = ofv.MENUheader("",funcion)


  'sHTML = "<html><body><meta name='robots' content='noarchive,noindex'>"
  'sHTML = "<html><head><meta name='robots' content='all'>" & chr(13)
  shtml= "<html><head><title>Indice de documentos vigentes</title><META NAME='robots' CONTENT='index, follow'> </head><body> "

'''  sHTML = sHTML & ofv.CorFormHeader






  stitem = "Tema"
  sIcono = "ICorTema.jpg"
  sNuevoTipo = "SUBTEMAS"
  nNivel = 0

  sID = 0
  nv = 0
  cborderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("APLSHW") & ";"
  cborderl = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("APLSHW") & ";"
  borderB = ""
  borderl = ""


'''''''''''''''''''''''''''''''''''''''''''
 ''' session("menu") = 3
  scodmaster = request.querystring("codmaster") & ""
  stipmaster = request.querystring("tipmaster") & ""
  sCODNIVEL  = request.querystring("CODNIVEL")


     scodmaster = "000"
     stipmaster = "PROCESO"
     sCODNIVEL  = 1

     SESSION("OROBJ") = 0




   wkproccod = "emdetaproc.asp?Nombre=" & request.querystring("Nombre")
   wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & stipmaster
   wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL
   wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

   Session("ASP")     = "vuelveprocesos.asp"
   Session("ASPSTR")  = wkproccod
   Session("ASPSTRD") = Session("ASPSTR")






''''''''''''''''''''''''''''''''''''''''''''




    WKPROCDESC =""' "LISTA DE DOCUMENTOS VIGENTES AL " & NOW()



   Celda = ""
  ' Celda = Celda & ofv.GenCelda("","Vt style='Xbackground-color:#993333;Xcolor:#ffffff;color:#990000;' ","","85%","","","",WKPROCDESC,"si")

  ' Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  ' TablaH = ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")




   Fila = ""

          wspcolor1  = "#990000"
          wspcolor2  = "#009900"
          sImageO1   = "&#0253;"
          sImageO2   = "&#0254;"
          xTrColor = "#FFFFFF"

      set cn = ofv.conectar(ofv.strconn2)




      Strsql = "Select A.*,b.descri as tobjdsc,c.estado From OBJETOS A"
      Strsql = Strsql & " INNER JOIN tipobjeto B ON ( B.tipobjeto = A.tipobjeto )"
      Strsql = Strsql & " INNER JOIN objprty c ON ( c.codobjeto = A.CODOBJETO )"
      Strsql = Strsql & " where estado = 3 and a.aplic = 'EM' and b.tipocpt = 'D' and b.FILTRO = 'S' "
      Strsql = Strsql & " ORDER BY tobjdsc, a.descri "

' response.write strsql & "<br>" & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO HAY PROCESOS, ACTIVOS O ACTIVIDADES ASOCIADAS","si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
     TablaN = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

     Tabla = "<CENTER><BR><BR>" & Tabla & "<BR><BR>" & TablaN & "</CENTER>"

  else

    filaN
    do until rs.eof



''''''''''''''''''''''''''''''


       wkproccod = "emdetaproc.asp?Nombre=" & request.querystring("Nombre")
       wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster
       wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL & "&PROPIEDIN=S"
       wkproccod = wkproccod & "&CODPROCESOIN=" & rs(0) & "&TIPPROCESOIN=" & rs(2)
       wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

       sHref     = wkproccod
       sImage    = "<font face='corflaticon' size=2 color='#FFFFFF'>4</font>"
       sLink     = ofv.GenLink("","nt",sHref,"","","",sImage,"si")

       'wkproccod = "<a class=nt TARGET='_BLANK' HREF='CORINDEXLISTASOL.ASP?CODIGO=" & RS(0) & "&TIPO=" & RS(2) & "' "
       wkproccod = "<a  HREF='CORINDEXLISTASOLEX.ASP?CODIGO=" & RS(0) & "&TIPO=" & RS(2) & "'"
       wkproccod = wkproccod & ">"
       wkproccod = wkproccod & rs("tobjdsc") & ":&nbsp;" & rs(3)
       wkproccod = wkproccod & "</a>"



       Celda     = ""
     '  Celda = Celda & ofv.GenCelda("","Ut","","90%","","","",wkproccod,"si")

       'FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

       filaN = filaN & wkproccod & "<br>" 'ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF











'''''''''''''''''''''''''''''''''



       if wkcant = 4 then
          wkcant = 1
          wkleft = wkleft + 10 + 130
          wktop  = 5
       else
          wkcant = wkcant + 1
          wktop  = wktop + 5 + 90
       end if


      rs.movenext
    loop

  end if

       Tabla = "<BR>" & Tabla & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")


      Tabla = filan '"<BR>" & Tabla & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")











  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = TablaH
  sHTML = sHTML & Celda

  Celda = Tabla
  sHTML = sHTML & Celda


  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML



%>

