<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  response.buffer=true

  dim wstipo(6)
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo      = request.querystring("Codobjeto")
  sTipo        = request.querystring("Tipobjeto")
  sAccId       = request.querystring("AccId")
  CantRegMover = request.querystring("CantRegMover")

  if sAccId  <> "" then
     Session("AccId") = sAccId
  else
     sAccId  = Session("AccId")
  end if

  vuelta  = ofv.vueltaasp("../",CantRegMover,CantRegMover) & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
  vuelta2 = ofv.vueltaasp("../",CantRegMover,CantRegMover)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  'Encabezados

  formulario = ""
  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)


  response.write sHTML
  sHTML = ""

' Acciones del procedimiento



  'Encabezados
  Celda = ofv.GenCelda("","ttt","","90%","","","","Propiedades del Paso","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")
  response.write sHTML
  sHTML = ""


  sHTML = ""
  response.flush

      Fila = ""

    strsql = "select * from objaccion "
    strsql = strsql & "WHERE objid = " &  sAccId
'    response.write strsql
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       posicion = 0




      nAccion = rs2(0)
      nAccionID = rs2(18)

      sForm = "edProcedimientos" & rs2(0)


    'Secuencia
     Celda = ofv.gencelda("","tt","","20%","","","","Paso","si")
     Celda = Celda & ofv.GenCelda("","vti","","2%","","","","&nbsp;","si")
 '    Celda = Celda & ofv.GenCelda("","vt","","1%","","","","&nbsp;","si")


      if bModificar then
        sInput = ofv.GenerarInput(rs2(10).name,rs2(10),"text",sForm,"10","dtn","SI")
      else
        sInput = rs2(10)
      end if


      Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","70%","","","",sInput,"si")

      Fila = Fila & ofv.genrow("","XXX height=30px","","","","","",Celda,"si")


        sAccion = " "
        sModo   = " "
        sArch   = "NO"
        sDscr   = "S/D"
        if rs2(4) <> "" then
          Strsql = "Select descri,modo From Tipoaccalt Where tipoaccionalt = " & rs2(4)
          set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsAcciones.eof then
             sAccion = rsAcciones(0)
             sModo   = rsAcciones(1)
          END IF
          call ofv.cerrarconsulta(rsAcciones)
        end if


      westado = "S"

      for it=1 to 6
          wstipo(it) = "N"
      next


      strsql12 = "SELECT * from modo "
      strsql12 = strsql12 & "Where tipomodo = '" & sModo & "' "
      set rs12 = ofv.crearconsultaEx(StrSql12,cn,1,volver)
      if not rs12.eof then
         sArch   = rs12(2)
         sDscr   = rs12(1)
         for it=1 to 6
             if rs12(it+3) then
                wstipo(it) = "S"
             end if
         next
      end if
      call ofv.cerrarconsulta(rs12)





    'Acciones
     Celda = ofv.gencelda("","tt","","20%","","","","accion","si")
     Celda = Celda & ofv.GenCelda("","vti","","2%","","","","&nbsp;","si")
  '   Celda = Celda & ofv.GenCelda("","vt","","1%","","","","&nbsp;","si")




     sInput1 = ofv.GenerarInput(rs2(4).name,rs2(4),"hidden",sForm,"2","dtn","SI")



      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccion,i-1)
        sAcciont = mid(sAccion,i+1)
      else
        sAcciont=""
      end if
      sInput = "<big><B><i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & replace(sAcciont," ","&nbsp;") & "</b></big>"
   ''   sInput = "<u>" & sInput & "</u>"

      sID = "EdProcedACC.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta

      sLink = ofv.GenLink("","cc title='Establecer Accion' ",sID,"","","",sInput,"si")

      Celda = Celda & ofv.GenCelda("","ttt colspan=2 ","","","","","",sLink & sInput1,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    'Icono
     Celda = ofv.gencelda("","tt","","20%","","","","tipo","si")
     Celda = Celda & ofv.GenCelda("","vti","","2%","","","","&nbsp;","si")
    ' Celda = Celda & ofv.GenCelda("","vt","","1%","","","","&nbsp;","si")


      if sArch <> "NO" then
        sImage = ofv.GenImage("","","","","Iconos/" & sArch,"0","")
      else
        sImage = ofv.GenImage("","","","","Iconos/ICorAImagen.jpg","0","")
      end if

      sImage = sImage & "&nbsp;&nbsp;&nbsp;" & sDscr
      sLink  = sImage

      if bModificar then
        sHref = "AgregarAcciones.asp?tabla=Tipoaccalt&Opcion=si&volver=" & vuelta
       '' sHref = sHref & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
        sLink = ofv.GenLink("","nt valign='middle' title='Agregar nueva accion a la lista' ",sHref,"","","",sImage,"si")
      end if

      Celda = Celda & ofv.GenCelda("","vt colspan=2  ","","","","","",sLink,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    'Relacion entre eventos (conectores)
     Celda = ofv.gencelda("","tt","","20%","","","","Conector","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

         sID = "EdProcedrel.asp?"
         sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
      ''   sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo

      if not isnull(rs2(12))  then
         if  clng(rs2(12)) > 0 then
             xclas  = "btbs"
             sInput = "<font class='btic' ><</font>"
         else
             xclas  = "btbn"
             sInput = "<font class='btic' ><</font>"
         end if
      else
         xclas  = "btbn"
         sInput = "<font class='btic' ><</font>"
      end if

      sHref = sID
      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      sInput = "<div class=" & xclas & " >" & sinput & "</div>"
      stitle = "Relacionar evento con conectores."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      sinput = "&nbsp;"
      Celda = Celda & ofv.GenCelda("","vt","","","","","",sinput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    'Texto de Referencia
     Celda = ofv.gencelda("","tt","","20%","","","","Referencia","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")


      sID = "TextoEsquema.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
  ''    sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo

      if rs2(11) <> "" or rs2(16) <> "" or rs2(17) <> "" then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID
      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar texto de referencia."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda  = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      sinput = "&nbsp;"
      Celda  = Celda & ofv.GenCelda("","vt","","","","","",sinput,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    'Link a Documento
  if wstipo(3) = "S" then
     Celda = ofv.gencelda("","tt","","20%","","","","Link a Documento","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

         sID = "EdProcedLnk.asp?"
         sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
     ''    sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo


         if not isnull(rs2(19))  then
            if  clng(rs2(19)) > 0 then
                xclas  = "btbs"
                sInput = "<font class='btic' >&#0126;</font>"
            else
                xclas  = "btbn"
                sInput = "<font class='btic' >&#0126;</font>"
            end if
         else
            xclas  = "btbn"
            sInput = "<font class='btic'  >&#0126;</font>"
         end if

         sHref = sID



      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Actualizar link a otro documento."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      sinput = "&nbsp;"
      Celda  = Celda & ofv.GenCelda("","vt","","","","","",sinput,"si")
      Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  end if

    'Circuito Alternativo
  if wstipo(1) = "S" then
     Celda = ofv.gencelda("","tt","","20%","","","","Circuito Alternativo","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

      session("volverproc") = ""

      sID = "AccionesAlt.asp?"
      sID = sID & "Accion=" & nAccionID & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
    ''  sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo


        strsql11 = "SELECT idaccion from accionmaster2 "
        strsql11 = strsql11 & "Where idaccion = '" & nAccionID & "'"
        if bVersion = "NO" then
           strsql11 = strsql11 & " and (modi = 0 or (modi <> 0 and "
           strsql11 = strsql11 & "tipmodi <> 'E')) "
        end if
        set rs11 = ofv.crearconsultaEx(StrSql11,cn,1,volver)
        if not rs11.eof then
           xclas  = "btbs"
           sInput = "<font class='btic'  >2</font>"
        else
           westado = "N"
           wntexto = "ACCION DE CONTROL: No tiene acciones alternativas informadas!."
            xclas  = "btbn"
           sInput = "<font class='btic'  >2</font>"
        end if
        call ofv.cerrarconsulta(rs11)
        sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer procedimiento alternativo (por salida negativa del control)."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      sinput = "&nbsp;"
      Celda  = Celda & ofv.GenCelda("","vt","","","","","",sinput,"si")

      Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   end if

    'Componentes
     Celda = ofv.gencelda("","tt","","20%","","","","Componentes","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")


      sID = "EdProcedCMP.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
    ''  sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
      sHref = sID

      wscompo =   rs2(6)
      if isnull(wscompo) then wscompo = "000"

      sComponente = "N"

        if rs2(6) <> "" then
          Strsql = "SELECT   nombre,archivo FROM CSACOMPLE where codcomp = " & rs2(6)
          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsComponentes.eof then
             sComponente = rsComponentes(0)
             sInput      = sComponente
             if len(sInput) > 50 then
                sInput   = left(sInput,47) & "..."
             end if
             sImage      = ofv.GenImage("","","","","Iconos/" & rsComponentes(1),"0","") & "&nbsp;&nbsp;" & sInput
             sHref2       = "EditarImagen2.asp?ID=" & rs2(6)
             sLink       = ofv.GenLink("","nt title='Visualizar el componente' ",sHref2,"_blank","","",sImage,"si")
          end if
          call ofv.cerrarconsulta(rsComponentes)
        end if

      sInput1    = ofv.GenerarInput(rs2(6).name,wscompo,"hidden",sForm,"2","dtn","SI")

      if sComponente <> "N" then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0163;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0163;</font>"
      end if

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Asociar componente."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      sinput = "&nbsp;"
      if sComponente <> "N" then
         sinput = sinput & sLink
      end if


      sinput = sinput & sInput1
      Celda  = Celda & ofv.GenCelda("","vt","","","","","",sinput,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Cantidad de copias
  if wstipo(2) = "S" and sComponente <> "N" then
     Celda = ofv.gencelda("","tt","","20%","","","","Copias","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")



        'Copias
        nCopias = 0
        if not isnull(rs2(13)) then nCopias = rs2(13)
        if cint(nCopias) = 0 then
           westado = "N"
           wntexto = "ACCION RELACIONADA A EMISION/RECEPCION DE DOCUMENTOS: No tiene informado la cantidad de copias a administrar!."
        end if
        if bModificar then

            SCOMSEP = ""
            scomK   = ""
            scomD   = ""

            For i=0 to 9
                scomK = scomK & SCOMSEP & i
                scomD = scomD & SCOMSEP & i
                SCOMSEP = chr(9)
            next

            sInput1 = ofv.gencomboFX(sForm,rs2(13).name,rs2(13),scomK,scomD,"")



          '  sInput1 = ofv.GenerarInput(rs2(13).name,nCopias,"textsb",sForm,"3","dtn","SI")
        else
          sInput1 = rs2(13)
        end if

        if nCopias > 0 then
           xclas  = "btbs"
        else
           xclas  = "btbn"
        end if

        sInput = "<font class='btic2' title='Cantidad de copias.' style='cursor:pointer;' >N</font>"
        sHref  = "javascript:void(0);"

        sInput = "<div class=" & xclas & "  ><center>" & sinput & "</center></div>"

        xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
        stitle = "Cantidad de copias."
        sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
        Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


        Celda = Celda & ofv.GenCelda("","vt","","","","","",sInput1,"si")

       Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    end if

         sInput1 = ""


    'Sector Continuador
     Celda = ofv.gencelda("","tt","","20%","","","","sector","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

      if  cstr(rs2(8)) <> "0"  then
          xclas  = "btbs"
          sInput = "<font class='btic' >G</font>"
      else
          xclas  = "btbn"
          sInput = "<font class='btic' >G</font>"
      end if

        sContinuador = ""
        if rs2(8) <> "" then
          Strsql = "Select descri From Eco Where eco = " & rs2(8)
          set rsContinuadores = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsContinuadores.eof then sContinuador = rsContinuadores(0)
          call ofv.cerrarconsulta(rsContinuadores)
        end if
        if sContinuador <> "" then
           sContinuador = Ucase("Continuador:  " & sContinuador)
           sContinuador = replace(sContinuador," ","&nbsp;")
           sContinuador = sContinuador
        end if
        sConti = sContinuador

      if rs2(8) = "" or cstr(rs2(8)) = "0" then
         sInput1 = ofv.GenerarInput(rs2(9).name,"","hidden","","40","dt","")
      end if

      sID = "EdProcedCON.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
   ''   sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
      sHref = sID
      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer Sector Continuador."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      sLink = ofv.GenLink("","nt title='" & stitle & "' ",sHref,"","","",sConti,"si") & sInput1

      Celda = Celda & ofv.GenCelda("","vt ","","","","","","&nbsp;&nbsp;" & sLink,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

         sInput1 = ""

   'Perfil Continuador
      if rs2(8) = "" or cstr(rs2(8)) = "0" then

      else



         Celda = ofv.gencelda("","tt","","20%","","","","Perfil","si")
         Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

          sPerfil = ""
          if rs2(9) <> "" then
             Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs2(9) & "'"
             set rsPerfiles = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rsPerfiles.eof then sPerfil = rsPerfiles(0)
             call ofv.cerrarconsulta(rsPerfiles)
          end if

          if sPerfil <> "" then
             sPerfil = Ucase("Perfil:  " & sPerfil)
             sPerfil = replace(sPerfil," ","&nbsp;")
             sPerfil = sPerfil
          end if
          sConti = sPerfil



          sID = "EdProcedPER.asp?"
          sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
          sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta
        ''  sID = sID & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo

        sper = rs2(9)

        if isnull(sper) then
           sImage = ofv.GenImage("","","","","Iconos/ICorIconSHV3.jpg","0","")
           xclas  = "btbn"
           sInput = "<font class='btic' >&#0128;</font>"
        elseif sper = " " then
           sImage = ofv.GenImage("","","","","Iconos/ICorIconSHV3.jpg","0","")
           xclas  = "btbn"
           sInput = "<font class='btic' >&#0128;</font>"
        else
           sImage = ofv.GenImage("","","","","Iconos/ICorIconSHL3.jpg","0","")
           xclas  = "btbs"
           sInput = "<font class='btic' >&#0128;</font>"
        end if

        sHref = sID
        sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer Perfil Continuador."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      sLink = ofv.GenLink("","nt title='" & stitle & "' ",sHref,"","","",sConti,"si")

      Celda = Celda & ofv.GenCelda("","vt ","","","","","","&nbsp;&nbsp;" & sLink,"si")

       Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
     end if

    'Version

         sInput1 = ""

      sSelect = ""
  if bVersion = "SI" then

     if rs2(14) then

        Celda = ofv.gencelda("","tt","","20%","","","","actualizacion","si")
        Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

              if rs2(15) = "A" then
                 stitle = "Paso agregado."
                 xclas  = "btba"
                 sInput = "<font class='btic'  >6</font>"
                 sInput1 = "Agregada"
              elseif rs2(15) = "E" then
                 stitle = "Paso eliminado."
                 xclas  = "btbn"
                 sInput = "<font class='btic'  >5</font>"
                 sInput1 = "Eliminada"
              elseif rs2(15) = "M" then
                 stitle = "Paso modificado."
                 xclas  = "btbs"
                 sInput = "<font class='btic'  >4</font>"
                 sInput1 = "Modificada"
              end if

         sHref  = "javascript:void(0);"
         xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

         sInput = "<div class=" & xclas & " >" & sinput & "</div>"
         sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
         Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

         if bModificar then
            scomK  = "TRUE" & chr(9) & "FALSE"
            scomD  = sInput1 & chr(9) & "No"
            sInput = ofv.gencomboFX(sForm,rs2(14).name,"TRUE",scomK,scomD," ")
         else
            sInput = sInput1
         end if

         Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")

         Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     end if

   end if

    'Observaciones
     Celda = ofv.gencelda("","tt","","20%","","","","Observaciones","si")
     Celda = Celda & ofv.GenCelda("","vti","","","","","","&nbsp;","si")

      if westado = "N" then
         xclas   = "btbn"
         sInput  = "<font class='btic'  >=</font>"
         vfondo  = "<font color=#993333>PASO CON OBSERVACIONES.</font>"

      else
          sInput  = "<font class='btic'  >=</font>"
          xclas   = "btbs"
          wntexto = "PASO CORRECTO - SIN OBSERVACIONES."
          vfondo  = wntexto
      end if

      sHref  = "javascript:void(0);"
      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Ver observaciones."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' onclick=" & chr(34) & "alert('" & wntexto & "');"  & chr(34),sHref,"","","",sInput,"si")

      Celda = Celda & ofv.GenCelda("","ut","","","","","",vfondo,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'fmto
     Celda = ofv.gencelda("","vt","","20%","","","","&nbsp;","si")
     Celda = Celda & ofv.GenCelda("","vti","","2%","","","","&nbsp;","si")
     Celda = Celda & ofv.GenCelda("","vt","","1%","","","","&nbsp;","si")
     Celda = Celda & ofv.GenCelda("","vt","","70%","","","","&nbsp;","si")
     Fila  = Fila & ofv.genrow("","","","","5","","",Celda,"si")


    Tabla = ofv.GenTabla("","","","95%","","0","","1","2","","",Fila,"si")



      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & rs2(0) & "&Tabla=Objaccion&volver=" & vuelta2
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Tabla,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",Tabla,"si")
      end if

      response.write Formulario


  end if

  Agregar = ""
  Buscar = ""

  'Boton Volver
  sAccion   = "edProcedDet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&CantRegMover=" & CantRegMover
  sRetornar = ofv.BotonVolver(sAccion)




'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0",rs2,"0","0","0","0","0","","","",Agregar,Buscar,sRetornar)

  call ofv.cerrarconsulta(rs2)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML
response.flush

%>