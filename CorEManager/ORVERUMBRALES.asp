<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  SCODINDI=REQUEST.QUERYSTRING("ID")
  'RESPONSE.WRITE SCODINDI

  sHTML = ofv.FormHeader(session("FormN"))

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

 SHTML=SHTML & UMBRALES(SCODINDI)
END IF

 Response.Write sHTML




 FUNCTION UMBRALES(SCODINDI)
  DIM RS
  DIM CELDA
  DIM SHTML
  DIM Formulario
  DIM sForm

    WKSETTABLA = "ORUMBRALES"

    FILA=""

    Strsql = "Select * From " & WKSETTABLA & " WHERE IDINDICADOR="&SCODINDI&" Order By ORDEN"
   ' RESPONSE.WRITE STRSQL
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then rs.move(CantRegMoverx)
    posI = 1
  Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)
      sForm = "Uform" & posi


        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")



      sInput = RS(1)
      Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")


      sInput = rs(3) 'ofv.GenerarInput(rs(3).name,rs(3),"text"& actualizar,sForm,"30","dt","")
      Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

       nCol = 40


             wdescri=""
  	       if  not  isnull(rs(6)) then
  	           wdescri = TRIM("" &rs(6))
  	       end if
  '
  	       if  len(wdescri) = 0 then wdescri = " "
  '
  	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
  	       t3b = t3b & "name=" & rs(6).name & " cols=" & nCol & " rows=2 "
  	       if  bModificar and modificaR then
  	           t3b = t3b & " onblur=" & sForm & ".submit() "
  	           t3b = t3b & " onkeypress=" & chr(34)
  	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
  	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
  	       else
  	           t3b = t3b & " readonly >"
  	       end if
  	       xdatos =  wdescri
  	       t3b    = t3b & xdatos & "</textarea></font>"
  	       SINPUT = rs(6) 't3b
  '
  '
  	    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")
       'desde

           sInput = RS(4) 'ofv.GenerarInput(rs(4).name,   rs(4),"textsb" & actualizar ,sForm,"10","dtN  readonly","")

          Celda = Celda & ofv.gencelda("","cc","","50 align=right","","","",sInput,"si")

         sInput = RS(5) 'ofv.GenerarInput(rs(5).name,rs(5),"textsb"& actualizar,sForm,"10","dtN","")
         Celda = Celda & ofv.gencelda("","cc","","50 align=right","","","",sInput,"si")


         sInput = "&nbsp;"
         Celda = Celda & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & RS(7) & ";' ","CENTER","5%","","","",sInput,"si")



      Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

     ' sAccion = ""
     ' if bModificar then
     '   sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
    '    sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta &"_XFUNCION=" & formula & "_CANTVAR=" & xcantvar
   '   end if
   '   Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

      posi = posi + 1
      rs.Movenext
    Loop
    'sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

      Celda = "<BR>" & ofv.gencelda("","VTI  COLSPAN=9 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>LISTA	 DE UMBRALES</BIG>","si")
      FILAC=ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
      Celda = ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
      Celda = CELDA & ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","1%","","","","<big>ORDEN</BIG>","si")
      Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","30%","","","","<big>DESCRIPCION</BIG>","si")
      Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","30%","","","","<big>OBSERVACION</BIG>","si")
      Celda = CELDA &ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","10%","","","","<big>DESDE</BIG>","si")
      CELDA = CELDA &  ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","10%","","","","<big>HASTA</BIG>","si")
      CELDA = CELDA &  ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","","","","","<big>COLOR</BIG>","si")
      FilaC  = FILAC & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    UMBRALES = ofv.GenTabla("","","","80% ","","0","","0","0","","",FILAC & FILA,"si")

  END FUNCTION



%>