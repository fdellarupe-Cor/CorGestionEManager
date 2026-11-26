<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("Opcion")
  t = request.querystring("t")
  Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")
  CantRegAMover = request.querystring("CantRegAMover")
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)



  sHTML = ofv.FormHeader(Session("FormN"))

  ofv.ObtenerAtributos Session("Form") ,""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cx =  ofv.conectar(ofv.strconn1)

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "SELECT * From CLIENTEPLAN  order by ID"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  xgrupo = " "
  xgrupoDSC = "NO ENCONTRADO"
  xgrp = "S"
  xusr = "S"

  if t <> "T" then
     strsql = "SELECT CODTEMA,DESCRI From PRODSERV Where CODPRODSERV = " & ID
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsu.eof then
        xgrupo    = rsu(0)
        xgrupoDSC = rsu(1)
     END IF
     call ofv.cerrarconsulta(rsu)

     strsql = "SELECT * From TEMASTATRM Where ATRMTIPOST = 'S' AND ATRMCODST = " & ID
     strsql = strsql & "  Order By EMPRE"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rsu.eof then xusr = "N"

     XDESCT = "SUBTEMA: "

  else

      xgrupo = ID
      xusr   = "N"
      XDESCT = "TEMA: "

  end if

     strsql = "SELECT * From TEMASTATRM Where ATRMTIPOST = 'T' AND ATRMCODST = " & xgrupo
     strsql = strsql & "  Order By EMPRE"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rsg.eof then xgrp = "N"


  redim aAtributos(3)
  aAtributos(0) = "ACCEDE"
  aAtributos(1) = "NO"
  aAtributos(2) = "NO"
  aAtributos(3) = "NO"


  redim aAtrm(3)
  aAtrm(0) = 1
  aAtrm(1) = 1
  aAtrm(2) = 1
  aAtrm(3) = 1

  redim aAtrmD(3)
  aAtrmD(0) = 1
  aAtrmD(1) = 1
  aAtrmD(2) = 1
  aAtrmD(3) = 1


  XDESCT = ""

  IF  T = "S" THEN
      XDESCT = XDESCT & "<big>TEMA:&nbsp;" & xgrupo  & "&nbsp;-&nbsp;" & xgrupoDSC & "</big><br><br>"
  END IF

  XDESCT = XDESCT & "ACCESO&nbsp;A&nbsp;EMPRESAS&nbsp;DEL&nbsp;SUBTEMA:&nbsp;" & ID & "&nbsp;-&nbsp;" & Nombre

  celdas = ofv.GenCelda("","ttt colspan=5","","","","","",XDESCT,"si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdaz = ofv.GenCelda("","ttr","","","","","","Empresa&nbsp;/&nbsp;Unidad&nbsp;de&nbsp;Negocios","si")
  iz = 0
    celdaz = celdaz & ofv.GenCelda("","ttr","","","","","",aAtributos(iz),"si")

  formulario = formulario & ofv.GenRow("","","","","25","","",celdaz,"si")



  do until rs2.eof

     xdef = "D"
     xcolor = ""

     aAtrm(0) = 1
     aAtrm(1) = 1
     aAtrm(2) = 1
     aAtrm(3) = 1

     aAtrmD(0) = 1
     aAtrmD(1) = 1
     aAtrmD(2) = 1
     aAtrmD(3) = 1

      xdescri = ""

      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs2(0) & ")</font>"
      xdescri = Ucase(rs2(1)) & " - "

      xdescri = xdescri & xdesatr & xdesc2

    celdaz = ofv.GenCelda("","ttt","","","","","",xdescri,"si")

    if xgrp <> "N" then
       if clng(rsg(0)) = clng(rs2(0)) then
          if not rsg(5)  then
             aAtrm(0) = abs(cint(rsg(5)))
          '   aAtrm(1) = abs(cint(rsg(6)))


             aAtrmD(0) = abs(cint(rsg(5)))
          '   aAtrmD(1) = abs(cint(rsg(6)))


             xdef = "T"
             xcolor = "#993333"
          end if
          rsg.movenext
          if rsg.eof then xgrp = "N"
        end if
     end if

    if xusr <> "N" then
       if clng(rsu(0)) = clng(rs2(0)) then
             aAtrm(0) = abs(cint(rsu(5)))
            ' aAtrm(1) = abs(cint(rsu(6)))

             xcolor = "#6699cc"
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if



    fila = ""
    celdas = ""

    sForm = "AtributosGrupos" & rs2(0)

      if bModificar then
        sAccion = "asp/ActualizarATRMTEMAST.asp?volver=" & vuelta & "&ID=" & rs2(0)
        sAccion = sAccion & "&grupousuario=" & ID & "&t=" & t  & "&Nombre="
        sAccion = sAccion & ofv.convertircar(Nombre," ","_") & "&CantRegAMover=" & CantRegAMover


      end if


 '   For iz = 0 to 1
    For iz = 0 to 0
      if aAtrm(iz) then
         sSelect = ""
         sSel    = 1
      else
         sSelect = "Selected"
         sSel    = 0
      end if

      if bModificar then

         IF IZ = 1 THEN
            IF  aAtrm(0) = 1 THEN
                sInput = "<font class=at><select name=" & aAtributos(iz) & " onchange=submit()>"
                sInput = sInput & "<option value=1>SI  <option value=0 " & sSelect
                sInput = sInput & ">NO  </select></font>"

                sInput = ofv.combosino(sForm & "X1",sSel,aAtributos(iz))
                sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")


            ELSE
                xcolor = ""
                SINPUT = "---"
                sInput = sInput & ofv.GenerarInput(aAtributos(iz),aAtrm(iz),"HIDDEN","","2","dt","")
            END IF
         ELSE
            sInput = "<font class=at><select name=" & aAtributos(iz) & " onchange=submit()>"
            sInput = sInput & "<option value=1>SI  <option value=0 " & sSelect
            sInput = sInput & ">NO  </select></font>"

            sInput = ofv.combosino(sForm & "X1",sSel,aAtributos(iz))
         '''   sInput = replace(sInput,"document.all.","document." & sForm & ".all.")
            sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")

         END IF



      else
          if  aAtrm(iz) = 1 then
              xestado = "SI"
          else
              IF  IZ = 1 THEN
                  IF  aAtrm(0) = 1 THEN
                      xestado = "NO"
                  ELSE
                      xcolor = ""
                      xestado = "---"
                  END IF
              ELSE
                  xestado = "NO"
              END IF
          end if
          sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
      end if

      xcolor2 = session("APLBGD")
      if aAtrm(iz) = 0 then
         if (aAtrm(iz) <> aAtrmD(iz)) or xdef = "T" then
            xcolor2 = xcolor
         end if
      end if

      sinput1 = ""
      sInput2 = ofv.GenerarInput(aAtributos(iz) & "D",aAtrmD(iz),"hidden","","2","dt","")
      if iz = 0 then
         sInput1 = ofv.GenerarInput("default",xdef,"hidden","","2","dt","")
      end if

    celdas = celdas & ofv.GenCelda("","ut style='border-color:" & xcolor2 & ";border-style:solid;border-width:2;' ","center","2%","","","",sInput + sinput2 + sinput1,"si")

    next




    formulario = formulario & ofv.GenRow("","","","","","","",celdaz & celdas,"si")
    rs2.movenext
  loop
  celdas = ofv.GenCelda("","abc colspan=5","center","","","","","&nbsp;","si")
  fila = ofv.GenRow("","","","","","","",celdas,"si")



  tabla = ofv.GenTabla("","aaa style='border-color:#dddddd;border-style:solid;border-width:1;' ","","70%","","","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center>" & tabla & "</center>"

  Agregar = ""
  Buscar  = ""

  IF T = "T" THEN
     sAccion = "TEMASSUBTEMAS.asp?CantRegMover=" & CantRegAMover
  ELSE
     sAccion = "SUBTEMAS.asp?Retorno=" & CantRegAMover & "&tema=" & xgrupo
  END IF
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  sHTML = sHTML & "</body></html>"

  if t <> "T" then call ofv.cerrarconsulta(rsU)
  call ofv.cerrarconsulta(rsG)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>