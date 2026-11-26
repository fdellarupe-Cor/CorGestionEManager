<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)
  set cn4 = ofv.conectar(ofv.strconn6)


  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKFPSC = request.querystring("WKFPSC")

 if WKFPSC = "" then
    WKFPSC = session("WKFPSC")
 else
    session("WKFPSC") = WKFPSC
 end if 


    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenF = "<B><FONT FACE=webDINGS COLOR=#993333 SIZE=5>q</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WINGDINGS COLOR=#cc0000 SIZE=3>&#0232;</FONT></B>"
    sImagenSI = "<B><FONT FACE=WINGDINGS COLOR=#0000aa SIZE=4>&#0236;</FONT></B>"
    sImagenSF = "<B><FONT FACE=WINGDINGS COLOR=#000000 SIZE=4>&#0237;</FONT></B>"
    sImagenSA = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=5>&#0252;</FONT></B>"
    sImagenSR = "<B><FONT FACE=WINGDINGS COLOR=#aa0000 SIZE=5>&#0251;</FONT></B>"
    sImagenPA = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=4>&#0231;</FONT></B>"
    sImagenPS = "<B><FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=4>&#0232;</FONT></B>"


    formularioA = ""
    formularioB = ""
    formularioC = ""


  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & session("WKF") & "' "
  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","TTT","","10%","","","",sInput,"si")

    sInput = RS(1) & " - " & RS(2) & ""
    celdas = celdas & ofv.GenCelda("","TTT","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    OBJTIPO = RS(6)

    SELECT CASE OBJTIPO

           CASE "REQ"
                ESTTABLA = "GPESTADOS"
                ESTCLAVE = "ID"
                ESTORDER = "   Order By DESCRI"
                SET CNX  = CN4

           CASE "COR"
                ESTTABLA = "ESTADOS"
                ESTCLAVE = "OBJID"
                ESTORDER = "  Order By ESTADO"
                SET CNX  = CN3

           CASE ELSE
                ESTTABLA = "ESTADOS"
                ESTCLAVE = "ID"
                ESTORDER = "   Order By DESCRI"
                SET CNX  = CN

    END SELECT


  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","-1","-1","","",filas,"si")


  END IF

    celdas = ""
    xrespo =  ""




  StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

       celdas = ""

       sInput = "#"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

       sInput = "DESCRIPCION DEL PASO"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

       sInput = "MODO"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

       sInput = "NIVEL RESPONSABLE"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

       sInput = "CICLO"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")


       filas = ofv.GenRow("","","","","","","",celdas,"si")



    sform = "WKF" & rs1(0)

       celdas = ""

       sc = "UTI"


       sInput = rs1(3) & ".&nbsp;"
       celdas = celdas & ofv.GenCelda("",sc,"right","","","","",sInput,"si")



       if bModificar then
          sInput = ofv.GenerarInput(rs1(4).name,rs1(4),"text",sform,"60","dt","")
       else
          sInput = ofv.GenerarInput("",rs1(4),"readonly","","60","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       select case rs1(20)

           case "U"

                 sInput = "UNITARIO"

           case "L"

                 sInput = "LISTA"

           case "LD"

                 sInput = "LISTA Y DEFAULT"

           case else 
           
                 sInput = "sin definir"


       end select 



       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



        if bModificar then
           StrSql = "select * from WKFPSCRESPNV WHERE WKFCOD = '" & session("WKF") & "' ORDER BY NIVELRSP"
       ' RESPONSE.WRITE STRSQL & "<BR>"
           sInput = ofv.GenerarCombo(sform,rs1(21).name,rs1(21),strsql,cn,4,3,"")
        else
           StrSql = "select DESCRIPCION from WKFPSCRESPNV WHERE NIVELRSP = 0" & rs1(21)
           set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsX1.eof then
              sInput = rsX1(0)
           ELSE
              sInput = rs1(21)
           END IF
        end if


    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       if bModificar then
          sInput = ofv.GenerarInput(rs1(22).name,rs1(22),"text",sform,"5","dtN","")
       else
          sInput = ofv.GenerarInput("",rs1(22),"readonly","","5","dtN","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs1(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOSNV"
      formulario = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = ofv.GenForm("","","","","post","",filas,"si")
    end if

    sHTML = sHTML  & "<CENTER>"

  sHTML = sHTML  & ofv.GenTabla("","aaa style='border-color:#AAAAAA;border-style:solid;border-width:1;' ","","80%","","0","","1","1","","",formulario,"si")


       celdas = ""

       sInput = "TIEMPO ESTIMADO"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

       sInput = "UNIDAD DE TIEMPO"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

       sInput = "ACCION ADICIONAL"
       celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")


       filas = ofv.GenRow("","","","","","","",celdas,"si")



    sform = "WKFUT" & rs1(0)

       celdas = ""

       sc = "UTI"


       if bModificar then
          sInput = ofv.GenerarInput(rs1(23).name,rs1(23),"text",sform,"5","dtN","")
       else
          sInput = ofv.GenerarInput("",rs1(23),"readonly","","5","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","25%","","","",sInput,"si")

       scomK = "4" & chr(9) & "1" & chr(9) & "2"
       scomD = "MINUTOS" & chr(9) & "HORAS" & chr(9) & "DIAS"


       if bModificar then
         sInput = ofv.gencomboFX(sForm,rs1(24).name,rs1(24),scomK,scomD," ")
       else

       select case CSTR(rs1(24) & "")

           case "1"

                 sInput = "HORAS"

           case "2"

                 sInput = "DIAS"

           case "3"

                 sInput = "SEMANAS"

           case "4"

                 sInput = "MINUTOS"

           case else 
           
                 sInput = "sin definir"


       end select 



       end if
       celdas = celdas & ofv.GenCelda("",sc,"","25%","","","",sInput,"si")



       scomK = "D" & chr(9) & "AR" & chr(9) & "IR" & chr(9) & "FR" & chr(9) & "SU" & chr(9) & "EEP"
       scomD = "NO" & chr(9) & "ASIGNAR RESPONSABLE" & chr(9) & "INICIAR REQUERIMIENTO" & chr(9)
       scomD = scomD & "FINALIZAR REQUERIMIENTO" & chr(9) & "SOLICITAR USUARIO" & chr(9) & "ENVIAR AL EMISOR DEL PASO"


       if bModificar AND ( RS1(6) = "N" AND RS1(7) = "N" ) then
         sInput = ofv.gencomboFX(sForm,rs1(25).name,rs1(25),scomK,scomD," ")
       else
       select case CSTR(rs1(25) & "")

           case "D"

                 sInput = "NO"

           case "AR"

                 sInput = "ASIGNAR RESPONSABLE"

           case "IR"

                 sInput = "INICIAR REQUERIMIENTO"

           case "FR"

                 sInput = "FINALIZAR REQUERIMIENTO"

           case "SU"

                 sInput = "SOLICITAR USUARIO"

           case "EEP"

                 sInput = "ENVIAR AL EMISOR DEL PASO"

           case else 
           
                 sInput = "sin definir"


       end select 
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")



    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs1(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOSNV"
      formulario = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = ofv.GenForm("","","","","post","",filas,"si")
    end if

    sHTML = sHTML  & "<CENTER>"

  sHTML = sHTML  & ofv.GenTabla("","aaa style='border-color:#AAAAAA;border-style:solid;border-width:1;' ","","80%","","0","","1","1","","",formulario,"si")




''''''''''''''''''''''''

''' ANTERIOR
    STAB = "&nbsp;"
  StrSql = "Select * From WKFPASOSPRTYNV WHERE PASOSIG = '" & session("WKFPSC") & "' "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS3.EOF THEN
    STAB = ""
     DO WHILE NOT RS3.EOF

        sInput = sImagenPA & "&nbsp;&nbsp;" & GENWKFPASO(RS3(2))
        
        STABC = ofv.GenCelda("","ut  ","","","","","",sInput,"si")
        STAB = STAB & ofv.GenRow("","","","","","","",STABC,"si")
     RS3.MOVENEXT
     LOOP
    STAB = ofv.GenTabla("","","","80%","","0","","0","0","","",STAB,"si")
  END IF

''' ANTERIOR
    celdas = ofv.GenCelda("","vtI style='background-color:#888888;color:#ffffff;' ","","","","","","ANTERIORES","si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")


    celdas = ofv.GenCelda("","ut  ","CENTER","","","","",STAB,"si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")
    STAB = ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    celdas = ofv.GenCelda("","","CENTER","","","","",STAB,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")
 

    sHTML = sHTML  & ofv.GenTabla("","aaa style='border-color:#AAAAAA;border-style:solid;border-width:1;' ","","80%","","0","","0","0","","",filas,"si") & "<BR>"




'''''''''''''''''''''''

    sInput = "OPCIONES DE BIFURCACION PARA EL PASO"
    celdas = ofv.GenCelda("","VTI style='background-color:#888888;color:#ffffff;' colspan=3","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","25","","",celdas,"si")








''' CONTROLA INICIAL '''''''''''''

    IF RS1(6) = "S" THEN

          WKINDF = 6
          WKINDB = "INI"
          WKINDBTIP = "1"

          sform = "WKF" & WKINDB
          sform2 = "WKF" & WKINDB & "PRTY"


      ARMAOPCION()

    END IF

''' CONTROLA FINAL '''''''''''''

    IF RS1(7) = "S"  THEN

          WKINDF = 7
          WKINDB = "FIN"
          WKINDBTIP = "2"

          sform = "WKF" & WKINDB
          sform2 = "WKF" & WKINDB & "PRTY"


      ARMAOPCION()

    END IF

''' CONTROLA NODOS '''''''''''''

    IF RS1(6) = "N" AND RS1(7) = "N" THEN


'''''' CONTROLA DERIVACION '''''''''''''


          WKIND  = 0


          WKINDF = 8
          WKINDB = "SAL"
          WKINDBTIP = "5"

          sform = "WKF" & WKINDB
          sform2 = "WKF" & WKINDB & "PRTY"


          ARMAOPCION()






       FOR WKIND=1 TO 5

          WKINDF = 9 + WKIND

          WKINDB = "N" & CSTR(100 + WKIND)

          WKINDB = REPLACE(WKINDB,"N1","N")   
          WKINDBTIP = "3"

          sform = "WKF" & WKINDB
          sform2 = "WKF" & WKINDB & "PRTY"

          ARMAOPCION()

       NEXT

    END IF

  sHTML = sHTML  & "<center>" & ofv.GenTabla("","aaa style='border-color:#993333;border-style:solid;border-width:2;' ","","80%","","0","","-2","-1","","",filas,"si")  & "</center>"



 END IF



  call ofv.cerrarconsulta(rs1)
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  Agregar = ""
  Buscar  = ""

  sAccion = "WKFPASOSNV.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKFPSCTPC(XTIPO,XITEMCOD)


    WKFCODPRTY = GTWKFPSC() & ITEMCOD
    WKFPASOSIG = ""
    WKFACCIONAD = "D"


    IF XTIPO = "" THEN XTIPO = "3"
    IF XITEMCOD = "" THEN XITEMCOD = "N01"

    IF XITEMCOD = "SAL" THEN 
       WKFPASOSIG  = session("WKFPSC")
       WKFACCIONAD = "SU"
    END IF

    StrSql = "Select * From WKFPSCTIPOS WHERE TIPO = '" & XTIPO & "' "
    set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs3.eof then 
       XTIPODESC = RS3(2)
    ELSE

       XTIPODESC = "NO ENCONTRADO"
    END IF


    strsql = "insert into WKFPASOSPRTYNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,WKFPASPRTYCOD,DESCRIPCION,"
    strsql = strsql & "TIPOPSC,ITEMCOD,PASOSIG,PERMITECLC,"
    strsql = strsql & "CAMEST,ESTADO,WKFACCIONAD"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & session("WKFPSC") & "','" & WKFCODPRTY & "','" & XTIPODESC & "',"
    strsql = strsql & "'" & XTIPO & "','" & XITEMCOD & "','" & WKFPASOSIG & "','S',"
    strsql = strsql & "'N','0','" & WKFACCIONAD & "'"
    strsql = strsql & ")"
 '   response.write strsql & "<br>" & chr(13)
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)





END FUNCTION

FUNCTION GENWKFPASO(XPASO)

  GENWKFPASO = "SIN ASOCIAR"

  StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & XPASO & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sinput = "<a CLASS=CC href=WKFPASOSPROPNV.asp?WKFPSC=" & RS2(2)
     sinput = sinput  & " ><b><U>" & replace(RS2(3) & ". " & RS2(4)," ","&nbsp;") & "</U></b></a>" 
     GENWKFPASO = sinput
  END IF
  call ofv.cerrarconsulta(rs2)

END FUNCTION

FUNCTION ARMAOPCION()

         WKFDESCRI = "<B>OPCION&nbsp;" & WKIND & "</B>"

         WKFFONDO = " bgcolor='#f7eedd' "


         IF WKINDF = 6 THEN

            WKFDESCRI = "<B><FONT COLOR='#336699'>INICIO</FONT></B>"
            WKFFONDO = " bgcolor='#ddddff' "


         ELSEIF WKINDF = 7 THEN

            WKFDESCRI = "<B><FONT COLOR='#008800'>FINALIZACION</FONT></B>"
            WKFFONDO = " bgcolor='#ddddff' "

         ELSEIF WKINDF = 8 THEN

            WKFDESCRI = "<B><FONT COLOR='#993333'>DERIVACION</FONT></B>"
            WKFFONDO = " bgcolor='#ddddff' "

         END IF





      IF rs1(WKINDF) = "N" THEN

         StrSql = "delete from WKFPASOSPRTYNV WHERE WKFCOD = '" & session("WKF") & "' and WKFPASCOD = '" & session("WKFPSC") & "' AND ITEMCOD = '" & WKINDB & "'"
         call ofv.crearconsultaEx(StrSql,cn,1,parametros)

         celdas = ""

         sc = "UT"

         sInput = WKIND
         celdas = celdas & ofv.GenCelda("","UTI  ","","5%","","","",sInput,"si")
 
         scomK = "N" & chr(9) & "S"
         scomD = "INACTIVO" & chr(9) & "ACTIVO"
         if bModificar then
            sInput = ofv.gencomboFX(sForm,rs1(WKINDF).name,rs1(WKINDF),scomK,scomD," ")
         else
            sInput = ofv.GenerarInput("",rs1(WKINDF),"readonly","","70","dt","")
         end if
         celdas = celdas & ofv.GenCelda("",sc & "  ","","10%","","","",sInput,"si")


         sInput = WKFDESCRI
         celdas = celdas & ofv.GenCelda("",sc,"","80%","","","",sInput,"si")

         if bModificar then
            sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs1(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOSNV"
            formulario = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
         else
            formulario = ofv.GenForm(sform,"","","","post","",celdas,"si")
         end if

         filas = filas & ofv.GenRow("","aaa bgcolor='#F5F5F5'  ","","","30","","",formulario,"si")

         celdas = ofv.GenCelda("","ut colspan=3 style='border-bottom-color:#DDDDDD;border-bottom-style:solid;border-bottom-width:1;' ","CENTER","","","","","&nbsp;","si")
         filas = filas & ofv.GenRow("","aaa bgcolor='#F5F5F5'  ","","","10","","",celdas,"si")


      ELSE

         StrSql = "SELECT * from WKFPASOSPRTYNV WHERE WKFCOD = '" & session("WKF") & "' and WKFPASCOD = '" & session("WKFPSC") & "' AND ITEMCOD = '" & WKINDB & "'"
         set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if rsX1.eof then

            xxxx = GENWKFPSCTPC(WKINDBTIP,WKINDB)

            StrSql = "SELECT * from WKFPASOSPRTYNV WHERE WKFCOD = '" & session("WKF") & "' and WKFPASCOD = '" & session("WKFPSC") & "' AND ITEMCOD = '" & WKINDB & "'"
            set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         END IF

         celdas = ""

         sc = "UT"

         sInput = WKIND
         celdas = celdas & ofv.GenCelda("","UTI  ","","5%","","","",sInput,"si")

         IF WKINDF = 6 OR WKINDF = 7 THEN

            IF rs1(WKINDF) = "N" THEN
               sInput = "INACTIVO"
            ELSE  
               sInput = "ACTIVO"
            END IF  

         ELSE 

 
            scomK = "N" & chr(9) & "S"
            scomD = "INACTIVO" & chr(9) & "ACTIVO"
            if bModificar then
               sInput = ofv.gencomboFX(sForm,rs1(WKINDF).name,rs1(WKINDF),scomK,scomD," ")
            else
               sInput = ofv.GenerarInput("",rs1(WKINDF),"readonly","","70","dt","")
            end if

         END IF

         celdas = celdas & ofv.GenCelda("",sc & " ","","10%","","","",sInput,"si")





         sInput = WKFDESCRI
         celdas = celdas & ofv.GenCelda("",sc,"","80%","","","",sInput,"si")



         if bModificar then
            sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs1(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOSNV"
            formulario = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
         else
            formulario = ofv.GenForm(sform,"","","","post","",celdas,"si")
         end if

         filas = filas & ofv.GenRow("","aaa " & WKFFONDO,"","","30","","",formulario,"si")




         filasw = ""
         celdas = ""

         sInput = "OPCION:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","15%","","","",sInput,"si")


         if bModificar then
            sInput = ofv.GenerarInput(rsX1(4).name,rsX1(4),"text",sform2,"60","dt","")
         else
            sInput = ofv.GenerarInput("",rsX1(4),"readonly","","60","dt","")
         end if
         celdas = celdas & ofv.GenCelda("","uti","","35%","","","",sInput,"si")

         sInput = "TIPO:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","15%","","","",sInput,"si")


         if bModificar then
            StrSql = "Select * From WKFPSCTIPOS order by TIPO"
       ' RESPONSE.WRITE STRSQL & "<BR>"
            sInput = ofv.GenerarCombo(sform2,rsX1(5).name,rsX1(5),strsql,cn,1,2,"")
         else
            StrSql = "Select descripcion From WKFPSCTIPOS WHERE TIPO = '" & rsX1(5) & "'"
            set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rsX2.eof then
               sInput = rsX2(0)
            ELSE
               sInput = rsX1(5)
            END IF
         end if
         celdas = celdas & ofv.GenCelda("","uti","","35%","","","",sInput,"si")



         filasw = filasw & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","30","","",celdas,"si")

         celdas = ""

         sInput = "PASO SIGUIENTE:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

         IF WKINDF = 7 THEN

            sInput = "---------&nbsp;"
            sInputk = "&nbsp;"

         ELSEIF WKINDF = 8 THEN

            WKPASO = rsX1(7) & ""
            if WKPASO <> "" then 
               sInput = GENWKFPASO(WKPASO)
            ELSE
               sInput = "SIN ASIGNAR"
            END IF
            sInputk = "&nbsp;"

         ELSE  


            WKPASO = rsX1(7) & ""
            if WKPASO <> "" then 
               sInput = GENWKFPASO(WKPASO)
            ELSE
               sInput = "SIN ASIGNAR"
            END IF


            sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
            sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
            sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
            sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
            sinputK = sinputK  & "none" & chr(34) & ";' "
            sinputK = sinputK  & "  href=WKFPSCSELPASNV.asp?WKFPSCTPC=" & rsX1(0)
            sinputK = sinputK  & "  TITLE='Configurar'  >" & sImagenI3 & "</a>" 

         END IF

         celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput & "&nbsp;" & sinputK,"si")

         sInput = "CICLADO:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")



         scomK = "N" & chr(9) & "S"
         scomD = "NO" & chr(9) & "SI"

         if bModificar then
             sInput = ofv.gencomboFX(sForm2,rsX1(8).name,rsX1(8),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",rsX1(8),"readonly","","70","dt","")
         end if

         sInput = "SI"


         celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")



         filasw = filasw & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","30","","",celdas,"si")


         celdas = ""

         sInput = "CAMBIO ESTADO:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

         scomK = "N" & chr(9) & "A" & chr(9) & "M"
         scomD = "NO" & chr(9) & "AUTOMATICO" & chr(9) & "MANUAL"

         if bModificar then
             sInput = ofv.gencomboFX(sForm2,rsX1(9).name,rsX1(9),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",rsX1(9),"readonly","","70","dt","")
         end if

         celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")


         IF rsX1(9) = "A" THEN

         sInput = "NUEVO ESTADO:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")


         if bModificar then
            strsql = "Select * From " & ESTTABLA & ESTORDER
       ' RESPONSE.WRITE STRSQL & "<BR>"
            sInput = ofv.GenerarCombo(sform2,rsX1(10).name,rsX1(10),strsql,cnX,0,1,"")

         else
            strsql = "Select * From " & ESTTABLA & " where " & ESTCLAVE & " = 0" & rsX1(10)
            set rsX2 = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
            if not rsX2.eof then
               sInput = rsX2(1)
            ELSE
               sInput = rsX1(10)
            END IF
         end if
         celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")


         ELSE

         sInput = "&nbsp;"
         celdas = celdas & ofv.GenCelda("","UTI COLSPAN=2 ","","","","","",sInput,"si")

         END IF


 
         filasw = filasw & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","30","","",celdas,"si")

         celdas = ""

         sInput = "ACCION ADICIONAL:"
         celdas = celdas & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")


       scomK = "D" & chr(9) & "AR" & chr(9) & "IR" & chr(9) & "FR" & chr(9) & "SU" & chr(9) & "EEP"
       scomD = "NO" & chr(9) & "ASIGNAR RESPONSABLE" & chr(9) & "INICIAR REQUERIMIENTO" & chr(9)
       scomD = scomD & "FINALIZAR REQUERIMIENTO" & chr(9) & "SOLICITAR USUARIO" & chr(9) & "ENVIAR AL EMISOR DEL PASO"


       if bModificar AND (RS1(6) = "N" AND RS1(7) = "N" ) AND WKINDF <> 8   then
         sInput = ofv.gencomboFX(sForm2,rsX1(11).name,rsX1(11),scomK,scomD," ")
       else
       select case CSTR(rsX1(11) & "")

           case "D"

                 sInput = "NO"

           case "AR"

                 sInput = "ASIGNAR RESPONSABLE"

           case "IR"

                 sInput = "INICIAR REQUERIMIENTO"

           case "FR"

                 sInput = "FINALIZAR REQUERIMIENTO"

           case "SU"

                 sInput = "SOLICITAR USUARIO"

           case "EEP"

                 sInput = "ENVIAR AL EMISOR DEL PASO"

           case else 
           
                 sInput = "sin definir"


       end select 
       end if
       celdas = celdas & ofv.GenCelda("","UTI COLSPAN=3 ","","","","","",sInput,"si")


         filasw = filasw & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","30","","",celdas,"si")


         wktabla = ofv.GenTabla("","aaa bgcolor='#FFFFFF'  style='border-color:#f7aa22;border-style:solid;border-width:1;' ","","80%","","0","","1","1","","",filasw,"si")
         celdas = ofv.GenCelda("","ut colspan=3 ","CENTER","","","","",wktabla,"si")

         if bModificar then
            sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rsX1(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOSPRTYNV"
            formulario = ofv.GenForm(sform2,"",sAccion,"","post","",celdas,"si")
         else
            formulario = ofv.GenForm(sform2,"","","","post","",celdas,"si")
         end if




         filas = filas & ofv.GenRow("","aaa " & WKFFONDO,"","","30","","",formulario,"si")


         celdas = ofv.GenCelda("","ut colspan=3 style='border-bottom-color:#f7aa22;border-bottom-style:solid;border-bottom-width:1;' ","CENTER","","","","","&nbsp;","si")
         filas = filas & ofv.GenRow("","aaa " & WKFFONDO,"","","20","","",celdas,"si")


      END IF

END FUNCTION


%>