<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)




  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  Fila = ""
  xTrColor = "#F9F9F9"
  WKFINSELECT = ""

  STRSQL = "SELECT TIPOBJETO  FROM TIPOBJETO WHERE APLIC IN ('EM','GLB')"
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   DELIM=""
   STIPOS=""
   DO WHILE NOT RS.EOF

   STIPOS=STIPOS & DELIM & &"'" & RS(0) & "'"
   DELIM=","
   RS.MOVENEXT
   LOOP

   STIPOS= " IN(" & STIPOS & ")"





  sImagevMa    = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
  sImageva    = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
  sImagevm    = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
  sImagevb    = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
  sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

  sImageSI    = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
  sImageNO    = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

  sImagevn    = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

  wkstylof = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
  wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "



      set cnx = ofv.conectar(ofv.strconn5)

      Strsql = ""


      Strsql = Strsql & " select a.* from ( "

      Strsql = Strsql & " SELECT b.*, cast('12/12/2050' as datetime) - (cast(substring(b.segcktpsc,10,5) as integer)  ) as FECCOMIENZO,  "
      Strsql = Strsql & " cast( ( getdate() - ( cast('12/12/2050' as datetime) - (cast(substring(b.segcktpsc,10,5) as integer)  ) )  ) as integer) as DIFER "
      Strsql = Strsql & " FROM CKTVIEW b "
      Strsql = Strsql & " ) A "

      Strsql = Strsql & " WHERE a.OBJTIPO " & STIPOS & "   AND "











      Strsql = Strsql & " a.FIN = 'NO' AND "
      Strsql = Strsql & " a.PSCFIN = 'NO' "

      Strsql = Strsql & " ORDER BY a.SEGCKT DESC, a.PSCRESPONSABLE, a.PSCCOMIENZO "

      response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
  if rs.eof  or WKFINSELECT="" then

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTG colspan=9 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY CIRCUITOS PENDIENTES PARA<BR>LA BUSQUEDA SOLICITADA","si")
           Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


  ELSE
  posI = 0

  rs.move(CantRegMoverx)

  wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;' "


        WKIPROCc = 0
        MAINPROC = 0
        MAINACTI = ""

           Celda = ""
           Celda = Celda & ofv.gencelda("","VTI colspan=7 style='background-color:#AAAAAA;color:#ffffff' " & wkstylo1,"","","","","","LISTA DE CIRCUITOS PENDIENTES","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")

           Celda = ""
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","40%","","","","PROCESO/SUBPROCESO - (PERIODO)","si")
           Celda = Celda & ofv.gencelda("","uti " & wkstylo1,"","25%","","","","USUARIO","si")
         ''  Celda = Celda & ofv.gencelda("","uti " & wkstylo1,"","","","","","CIRCUITO","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","25%","","","","PASO","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","5%","","","","DIAS DEMORA","si")
           Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","5%","","","","INICIO","si")
       '    Celda = Celda & ofv.gencelda("","utI " & wkstylo1,"","5%","","","","REASIGNAR","si")
           Fila  = FILA & ofv.genrow("","aaa bgcolor='#ddddff'  ","","","25","","",Celda,"si")
' response.write rs.EOF & "<br>"
  Do While (Not rs.EOF)  and (posi < CantRegMostrarx)

        sInput2 = rs("feccomienzo")
        sinput3 = rs("difer")
        WKIACTIV = ""

        WKIPROCD = ""
        WKIPROCN = 0

        Celda = ""
        sInput = RS(4)

        Strsql = ""
        Strsql = Strsql & " select x.periodo, A.*, W.ESTADO "
        Strsql = Strsql & " from csrelpathdescri a   "
        Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "
        Strsql = Strsql & " inner join orrskcircuito w on (a.cobj = w.codobjeto)   "
        Strsql = Strsql & " inner join orperiodos x on (w.periodo = x.objid )  "
        Strsql = Strsql & " where  w.OBJID = 0" & sInput & "   "


 ' response.write strsql & "<br>"


        set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT rsX1.eof THEN

          FOR WKI=1 TO 8

            WKIPROC = rsX1(WKI) & ""

            IF WKIPROC <> "" THEN
               IF WKIPROCD = "" THEN WKIPROC = "<B>" & WKIPROC & "</B>"
               WKIPROCD = WKIPROCD & WKIPROC & " / "
            END IF


          NEXT

          IF LEN(WKIPROCD) > 4 THEN
             WKIPROCD = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
          END IF


          sInput = WKIPROCD & " - (" & rsX1(0) & ")"

       END IF
       call ofv.cerrarconsulta(rsX1)

      IF sNombre = "NOVEDADES" THEN

         sInputX =  sInput

         Session("ASP")      = "novedmenu.asp?menu=4"
         session("sEstado")  = "PARA APROBAR"
         session("sEstadon") = "2"
         Session("Form2")    = "apr"
         xMAINAPL            = WKAPLIC

         aspeditar = "Editar.asp?Codigo=" & rs(4) & "&Tipo=" & rs(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RS(8) & "&MAINAPL=" & xMAINAPL

         aspeditar = cstr(aspeditar)
         sHref     = "&nbsp;"
         sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

         sInput =  "<A HREF='JAVASCRIPT:VOID(0);' "
         sInput =  sInput & "CLASS='NT' TITLE='ABRIR NOVEDAD.' " & sFuncion & "><U>" & sInputX & "</U></A>"

      END IF



       Celda = Celda & ofv.gencelda("","UTi BGCOLOR='#F7EEDD' STYLE='COLOR:#000000;'  " & wkstylo1,"","","","","",sInput,"si")


    sInputX = RS(14)

    xxxx = getInfoUser(RS(14))
    if session("SGU_nombre") <> "" then
       sInputX =  session("SGU_nombre")
    end if



      IF sNombre = "NOVEDADES" THEN
         sInput =  sInputX
      ELSE

         sInput =  "<A HREF='ORWKFREASIGNADOR.ASP?WKFCOD=" & RS(8) & "&CantRegMoverx=" & CantRegMoverx &"' "
         sInput =  sInput & "CLASS='NT' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>" & sInputX & "</U></A>"

      END IF


    Celda = Celda & ofv.gencelda("","UTi  " & wkstylo1 & "  ","","","","","",sInput,"si")

    sInput = RS(7) & "/<BR>" & RS(16)
    Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")

 '   sInput = RS(16)
 '   Celda = Celda & ofv.gencelda("","UTi " & wkstylo1,"","","","","",sInput,"si")





    sinput = sinput3 & "&nbsp;"

    Celda = Celda & ofv.gencelda("", "UT " & wkstylo1,"RIGHT","","","","",sInput,"si")

    sInput = GTFechadorSh(CLNG(sInput2) - CLNG(CDATE("01-01-1900"))) & "&nbsp;"
    Celda = Celda & ofv.gencelda("", "UT " & wkstylo1,"RIGHT","","","","",sInput,"si")


    sInput = "<INPUT TYPE='BUTTON' VALUE='ASIGNAR' CLASS=BT >"
  '  Celda = Celda & ofv.gencelda("", "UT " & wkstylo1,"RIGHT","","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")



       if xTrColor = "#FFFFFF" then
          xTrColor = "#F9F9F9"
       ELSE
          xTrColor = "#FFFFFF"
       END IF






    posi = posi + 1
    rs.Movenext
  Loop

           Celda = ""
           Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","25%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Celda = Celda & ofv.gencelda("","vt","","10%","","","","&nbsp;","si")
           Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")


  END IF
  sHTML = sHTML & ofv.GenTabla("","aaa ","","95% ","","0","","0","-1","","",Fila,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""

  sAccion = "ORSUBCONSULTA.asp?Nombre=" & sNombre & "&Menu=" & sMenu
  Celda = ofv.BotonVolver(sAccion)


'  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cnx)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>