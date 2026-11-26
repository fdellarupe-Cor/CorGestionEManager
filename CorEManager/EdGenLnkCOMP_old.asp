<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nblk = request.querystring("nblk")
  if len(nblk) = 0 then
     nblk = "S"
     vueltaC  = request.querystring("volver")
     session("volverproc") = ""
  else
     vueltaC  = ""
  end if

  nTabla   = request.querystring("tabla")
  nid      = request.querystring("id")
  nCodigo  = request.querystring("Codobjeto")
  sTipo    = request.querystring("Tipobjeto")


  if session("volverproc") = "" then session("volverproc") = vueltaC
  if vueltaC = "" then vueltaC = session("volverproc")

  volver2 = ofv.vueltaaspEx("../","0","0") & "_Id=" & nId & "_tabla=" & Ntabla & "_nblk=S"

  esttxt = " STYLE={font-family:tahoma;color:#4466bb;font-size:10pt;font-weight:bold;background-color:#ffffff;} "

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * From " & ntabla & " where objid = " & nID
''response.write Strsql & " Strsql "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     select case ucase(ntabla)

            case "OBJACCION"

                  wproced = rs(3)
                  wID     = rs(18)
                  wlink   = rs(6)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(20)
                  wDAT    = 20
                  wlnkname = rs(6).name


            case "OBJESTDET"

                  wproced = nid
                  wID     = nid
                  wlink   = rs(9)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(33)
                  wDAT    = 33
                  wlnkname = rs(9).name

      END SELECT




  end if

  call ofv.cerrarconsulta(rs)


''''''''''''''''''  COMIENZO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  sHTML = ofv.FORMheader("")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  sHTML = sHTML & "<SCRIPT language='javascript'  >" & chr(13) & chr(13)
  sHTML = sHTML & "    var myGlos;                           " & chr(13)
  sHTML = sHTML & "</SCRIPT>" & chr(13) & chr(13)
  sHTML = sHTML & "<SCRIPT language='javascript' for='LINKTXT' event='onmouseout' >" & chr(13) & chr(13)

  sHTML = sHTML & " if (document.selection.type=='Text') {;"

  sHTML = sHTML & "     myGlos = document.selection.createRange().text;"

  sHTML = sHTML & "     document.selection.createRange().text = '||||$' + myGlos + '$||||' ;"
  sHTML = sHTML & "     document.getElementById('TDLINKTXT').style.display = 'none' ;"
  sHTML = sHTML & "     document.getElementById('TD2LINKTXT').style.display = '' ;"
  sHTML = sHTML & "     document.getElementById('TD3LINKTXT').style.display = '' ;"
  sHTML = sHTML & "     document.getElementById('TD4LINKTXT').style.display = '' ;"
  sHTML = sHTML & "     document.getElementById('LINKTXT2').innerHTML = 'El texto asociado es: ' + myGlos;"
  sHTML = sHTML & "     document.getElementById(xnomdato).value = document.getElementById('LINKTXT').innerHTML;"
'''' sHTML = sHTML & "     document.selection.createRange().text = '<@>' + myGlos + '<@>' ;"



''' sHTML = sHTML & "     alert(document.all[xnomdato].value);"
  sHTML = sHTML & " };"

' sHTML = sHTML & "     alert('entro');"


  sHTML = sHTML & "</SCRIPT>" & chr(13) & chr(13)

''  sHTML = sHTML & "<p>ing: " & request.querystring & "<br><br></p>"

  sHTML = sHTML & "<center><br>"



''''''' inserta link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if nblk = "IL" then


     wlink   = request.form("comp1") & ""
     wlinkT  = request.form("comp2") & ""
     wlinkN1 = request.form("compN1") & ""
     wlinkN2 = request.form("compN2") & ""
     wlinkN3 = request.form("compN3") & ""
     wlinkN4 = request.form("compN4") & ""
     wlinkN5 = request.form("compN5") & ""



     if wlink > 0 then

        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
        strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & nCodigo & ",'" & sTipo & "',"
        strsql = strsql & "'" & wproced & "','" & wId & "',"
        strsql = strsql & wlink & ",'" & wlinkT & "','','','KL','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
        strsql = strsql & ")"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & wId & "'  AND LTIPO = 'KL'  ORDER BY OBJID DESC "
      ''response.write Strsql & " Strsql "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then
           wlnkid = rs(0)
        else
           wlnkid = 0
           wlink = 0
        end if
        call ofv.cerrarconsulta(rs)

        if clng(wlnkid) > 0 then
           strsql = "update " & ntabla & " set " & wlnkname & " = " & wlnkid
           strsql = strsql & " where  objid = " & nID
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if
     end if

     nblk = "S"
  end if

''''''' FIN inserta link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''





''''''' agrega link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if nblk = "AL" then

     wtipoL   = 0
     wtipOBJ  = 0
     wUrl     = "NO"
     wUrlDesc = "NO"
     wtipo    = "NO"

     wNV1    = "0"
     wNV2    = "0"
     wNV3    = "0"

   ''response.write wtipoL  & " wtipoL " & "<br>"
   ''response.write wtipOBJ & " wtipOBJ " & "<br>"

     sHTML = sHTML & GenBuscaCMP(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0")
     sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 ","","60%","",0,0,0,0,"","","","no")

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")
     sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "</FORM>" & chr(13)

     if bModificar then
        sAccion = "EdGenLnkCOMP.asp?"
        sAccion = sAccion & "Id=" & nId  & "&nblk=IL&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
        sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
     else
        sAccion = ""
     end if
     sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)


     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  colspan=2 align=center  >"

     sHTML = sHTML & "<input name='comp1'   type='hidden'  size=1  value='" & XGUFCOBJ & "' class=bt >"
     sHTML = sHTML & "<input name='comp2'   type='hidden'  size=1  value='" & XGUFTOBJ & "' class=bt >"
     sHTML = sHTML & "<input name='compn1'  type='hidden'  size=1  value='" & XGUFNIV1 & "' class=bt >"
     sHTML = sHTML & "<input name='compn2'  type='hidden'  size=1  value='" & XGUFNIV2 & "' class=bt >"
     sHTML = sHTML & "<input name='compn3'  type='hidden'  size=1  value='" & XGUFNIV3 & "' class=bt >"
     sHTML = sHTML & "<input name='compn4'  type='hidden'  size=1  value='" & XGUFNIV4 & "' class=bt >"
     sHTML = sHTML & "<input name='compn5'  type='hidden'  size=1  value='" & XGUFNIV5 & "' class=bt >"

     sHTML = sHTML & "</td>"
     sHTML = sHTML & "</tr>"

     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  align=center   >"

     sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"

     sHTML = sHTML & "</td>"

     sHTML = sHTML & "<td  align=center   >"
     sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
     sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "</tr>"

     sHTML = sHTML & "</FORM>" & chr(13)
     sHTML = sHTML & "</table><br><br>"


     nblk = "NO"

  end if
''''''' FIN agrega link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






''''''' Alta 1era vez ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if clng("0" & wlink) = 0 and nblk = "S" then


  ''''Para obtener la descripcion de la Accion
    ''Primero obtiene el tipo de accion de la tabla objaccion
    Strsql = "SELECT * From OBJACCION "
    Strsql = Strsql & " WHERE codproced = '" & wProced & "' "
    Strsql = Strsql & " AND objid = '" & nID & "' "
  ''response.write Strsql & " Strsql "
    set rsz = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsz.eof then
  	 nTAccion  = rsz(4)
    end if
    call ofv.cerrarconsulta(rsz)

    ''Despues obtiene la descri de la accion de la tabla Tipoaccalt
    Strsql = "Select descri From Tipoaccalt "
    Strsql = Strsql & " Where tipoaccionalt = '" & nTAccion & "' "

    set rsy = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsy.eof then
    nTdescri  = rsy(0)
    end if
  call ofv.cerrarconsulta(rsy)

     sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=5  ","","40%","",0,0,0,0,"","","","no")
     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  colspan=2 class=ttt align=center >"
     sHTML = sHTML & "AGREGAR ASOCIACION CON COMPONENTE - "
     sHTML = sHTML & nTdescri
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "</tr>"
     sHTML = sHTML & "<tr HEIGHT=60 >"
     sHTML = sHTML & "<td  COLSPAN=2 align=center   >&nbsp;</TD>"
     sHTML = sHTML & "</tr>"





     sHTML = sHTML & "<tr>"


     sHTML = sHTML & "<td  align=RIGHT WIDTH='50%'  >"

'' -------------------------- BOTON AGREGAR ---------------------------------------------

     if bModificar then
        sAccion = "EdGenLnkCOMP.asp?"
        sAccion = sAccion & "Id=" & nId  & "&nblk=AL&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
        sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
     else
        sAccion = ""
     end if

     sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='AGREGAR ASOCIACION' class=bt >"
     sHTML = sHTML & "</FORM>" & chr(13)

     sHTML = sHTML & "</td>"


     sHTML = sHTML & "<td  align=center WIDTH='50%'  >"

'' -------------------------- BOTON CANCELAR ---------------------------------------------

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")
     sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "<input  type=submit  size=10%  value='CANCELAR' class=bt >"
     sHTML = sHTML & "</FORM>" & chr(13)

     sHTML = sHTML & "</td>"


     sHTML = sHTML & "</tr>"

     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  COLSPAN=2 align=center   >&nbsp;</TD>"
     sHTML = sHTML & "</tr>"
     sHTML = sHTML & "</table><br><br>"


     nblk = "NO"


  end if

''''''' FIN Alta 1era vez ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''' modicar link (OBTIENE CLAVE) '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ''response.write nblk & " nblk " & "<br>"


  if nblk = "ML" OR nblk = "MLD" OR nblk = "MI" OR nblk = "MID" OR nblk = "A"  then


     Strsql = "SELECT * From " & ntabla & " where objid = " & nID
   ''response.write Strsql & " Strsql " & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then

     select case ucase(ntabla)

            case "OBJACCION"


                  wlinkEY   = rs(6)



            case "OBJESTDET"


                  wlinkEY   = rs(9)


      END SELECT

    END IF


  END IF



''''''' modicar link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "ML" then

     wlinkD = request.form("comp2") & ""

        strsql = "update OBJLNK set LurlDesc = '" & wlinkD & "' "
        strsql = strsql & "where objid = " & wlinkEY
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

     nblk = "S"

  end if

''''''' FIN modicar link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''' modicar link local D '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "MLD" then

     wlink   = request.form("comp1") & ""
     wlinkT  = request.form("comp2") & ""
     wlinkN1 = request.form("compN1") & ""
     wlinkN2 = request.form("compN2") & ""
     wlinkN3 = request.form("compN3") & ""
     wlinkN4 = request.form("compN4") & ""
     wlinkN5 = request.form("compN5") & ""

     if wlink > "0" then




        strsql = "update OBJLNK set codobjetoL = " & wlink & " "
        strsql = strsql & ", tipobjetoL = '" & wlinkT & "' "
        strsql = strsql & ", IDEN3L = '" & wlinkN1 & "' "
        strsql = strsql & ", IDEN4  = '" & wlinkN2 & "' "
        strsql = strsql & ", IDEN4L = '" & wlinkN3 & "' "
        strsql = strsql & "where objid = " & wlinkEY
      ''  RESPONSE.WRITE STRSQL & "<BR>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



     end if

     nblk = "S"
  end if


''''''' FIN modicar link local D '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''' anular link ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "A" then

 '    response.write "anular link: " & "<br>"

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")

     strsql = "delete from OBJLNK "
     strsql = strsql & "where objid = " & wlinkEY
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

 ''    response.write "anular link: " & strsql & "<br>"

           strsql = "update " & ntabla & " set " & wlnkname & " = 0 "
           strsql = strsql & " where  objid = " & nID
           call ofv.crearconsultaEx(StrSql,cn,1,volver)

     response.redirect sAccion

     nblk = "NO"

  end if
''''''' FIN anular link ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''' Mostrar link '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "S" OR nblk = "SD" then

     Strsql = "SELECT * From " & ntabla & " where objid = " & nID
     response.write Strsql & " Strsql "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then

       select case ucase(ntabla)

            case "OBJACCION"

                  wproced = rs(3)
                  wID     = rs(18)
                  wlink   = rs(6)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(20)
                  wDAT    = 20
                  wlnkname = rs(6).name

            case "OBJESTDET"

                  wproced = nid
                  wID     = nid
                  wlink   = rs(9)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(33)
                  wDAT    = 33
                  wlnkname = rs(9).name

        END SELECT


        wtipoL   = 0
        wtipOBJ  = 0
        wUrl     = "NO"
        wUrlDesc = "NO"
        wtipo    = "NO"

        wNV1    = "0"
        wNV2    = "0"
        wNV3    = "0"

       Strsql = "SELECT * From objlnk where objid = " & wlink
       response.write Strsql & Strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs1.eof then
          wtipoL   = clng(rs1(6))
          wtipOBJ  = rs1(7)

          response.write rs1(7) & " rs1(7) " & "<br>"


          wUrl     = rs1(8)
          wUrlDesc = rs1(9)
          wtipo    = rs1(10)

          wNV1    = rs1(11) & ""
          wNV2    = rs1(12) & ""
          wNV3    = rs1(13) & ""

          IF  WNV1 = "" THEN WNV1 = "0"
          IF  WNV2 = "" THEN WNV2 = "0"
          IF  WNV3 = "" THEN WNV3 = "0"


          wasotxt     = rs1(14) & ""
          wasotxtstr  = rs1(15)
          wasotxtcan  = rs1(16)
          wasotxtdat  = rs1(17) & ""

       end if
       call ofv.cerrarconsulta(rs1)

 '' response.write  wtipoL & "-" & wtipOBJ  & "-"  &  wUrl & "-"  & wUrlDesc & "-"  & wtipo & "<br>"

'''----'''''' LINK A INTERNET O EXTERNO '''''''''''''''''''''''''''''''''''''''


       if wtipo = "I" then




      else

'-------- muestra link local



         if nblk = "SD" then

            response.write wtipoL & " wtipoL " & "<br>"
            response.write wtipOBJ & " wtipOBJ " & "<br>"


            sHTML = sHTML & GenBuscaCMP(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0")
            sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

            sHTML = sHTML & "<tr>"
            sHTML = sHTML & "<td  align=center   >"

            sHTML = sHTML & "<input  type=Button size=10%  value='Aceptar' class=bt "
            sHTML = sHTML & " onclick=" & chr(34) & "fcomp3.submit();" & chr(34) &  " >"
            sHTML = sHTML & "</td>"

            sHTML = sHTML & "<td  align=center   >"
            sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
            sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
            sHTML = sHTML & "</td>"
            sHTML = sHTML & "</tr>"

            sHTML = sHTML & "</FORM>" & chr(13)

            sHTML = sHTML & "</table><br><br>"

            sAccion =  replace(vueltaC,"../","")
            sAccion =  replace(sAccion,"_","&")

            sAccion = "EdGenLnkCOMP.asp?"
            sAccion = sAccion & "Id=" & nId  & "&nblk=S&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
            sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC

            sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
            sHTML = sHTML & "</FORM>" & chr(13)

            if bModificar and XGUFCOBJ > 0  then
               sAccion = "EdGenLnkCOMP.asp?"
               sAccion = sAccion & "Id=" & nId  & "&nblk=MLD&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
               sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
            else
               sAccion = ""
            end if
            sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
            sHTML = sHTML & "<input name='comp1'   type='hidden'  size=1  value='" & XGUFCOBJ & "' class=bt >"
            sHTML = sHTML & "<input name='comp2'   type='hidden'  size=1  value='" & XGUFTOBJ & "' class=bt >"
            sHTML = sHTML & "<input name='compn1'  type='hidden'  size=1  value='" & XGUFNIV1 & "' class=bt >"
            sHTML = sHTML & "<input name='compn2'  type='hidden'  size=1  value='" & XGUFNIV2 & "' class=bt >"
            sHTML = sHTML & "<input name='compn3'  type='hidden'  size=1  value='" & XGUFNIV3 & "' class=bt >"
            sHTML = sHTML & "<input name='compn4'  type='hidden'  size=1  value='" & XGUFNIV4 & "' class=bt >"
            sHTML = sHTML & "<input name='compn5'  type='hidden'  size=1  value='" & XGUFNIV5 & "' class=bt >"
            sHTML = sHTML & "</FORM>" & chr(13)



         ELSE



            sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

            sHTML = sHTML & "<tr>"
            sHTML = sHTML & "<td  colspan=2 class=ttt >"
            sHTML = sHTML & "ASOCIACION AL COMPONENTE - "
            sHTML = sHTML & nTdescri
            sHTML = sHTML & "</td>"
            sHTML = sHTML & "</tr>"


            sHTML = sHTML & "<tr>"
            sHTML = sHTML & "<td  colspan=2 align=center  >"





            XMST = "0"

            IF wtipoL <> 0 THEN

               XMST = "1"

             ''response.write rs1(0) & " rs1(0) " & "<br>"
             ''response.write wtipoL & " wtipoL " & "<br>"
             ''response.write wtipOBJ & " wtipOBJ " & "<br>"

               response.write rs1(6) & " rs1(6) " & "<br>"
               response.write rs1(7) & " rs1(7) " & "<br>"

               wtipOBJ  = rs1(7)

               response.write wtipOBJ & " wtipOBJ " & "<br>"

                CALL GETLISTCMP("E",wtipoL,wtipOBJ,wtipolist)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT   >"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"

                IF WNV1 <> "0" THEN

                   XMST = "2"


                   CALL GETLISTSLPCMP("E",wtipoL,wtipOBJ,WNV1)
                   XGUFCOBJKEY = XGUTABLISKEY
                   XGUFCOBJDAT = XGUTABLISDAT


                   XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                   XGUFBTDF1 = XGUFBTDF1 & " >"

                   XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT   >"
                   XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                   XGUFBTDF1 = XGUFBTDF1 & "</TR>"

                   IF WNV2 <> "0" THEN

                      XMST = "3"

                      CALL GETLISTNV2CMP("E",wtipoL,wtipOBJ,WNV1,WNV2)
                      XGUFCOBJKEY = XGUTABLISKEY
                      XGUFCOBJDAT = XGUTABLISDAT


                      XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                      XGUFBTDF2 = XGUFBTDF2 & " >"

                      XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT   >"
                      XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                      XGUFBTDF2 = XGUFBTDF2 & "</TR>"

                      IF WNV3 <> "0" THEN

                         XMST = "4"

                         CALL GETLISTNV3CMP("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3)
                         XGUFCOBJKEY = XGUTABLISKEY
                         XGUFCOBJDAT = XGUTABLISDAT


                         XGUFBTDF3 = XGUFBTDF3 & "<TR  "
                         XGUFBTDF3 = XGUFBTDF3 & " >"

                         XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT   >"
                         XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
                         XGUFBTDF3 = XGUFBTDF3 & "</TR>"


                      END IF
                   END IF
                END IF
             END IF



             IF XMST > "0" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


             END IF

             IF XMST > "1" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


             END IF

             IF XMST > "2" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


             END IF

             IF XMST > "3" THEN

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


             END IF

             IF XMST > "0" THEN


                sAccion = "EdGenLnkCOMP.asp?"
                sAccion = sAccion & "Id=" & nId  & "&nblk=SD&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
                sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC

                XGUFCOBJDAT = ""
                XGUFCOBJDAT = XGUFCOBJDAT & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
                XGUFCOBJDAT = XGUFCOBJDAT & "<input  type=SUBMIT  size=10%  value='MODIFICAR' class=bt >"
                XGUFCOBJDAT = XGUFCOBJDAT & "</FORM>" & chr(13)

                XGUFBTDF3 = ""
                XGUFBTDF3 = XGUFBTDF3 & "<TR  "
                XGUFBTDF3 = XGUFBTDF3 & " >"

                XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT ALIGN=RIGHT  >"
                XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF3 = XGUFBTDF3 & "</TR>"

                XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSRB""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
                XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


             END IF


             sHTML = sHTML & XGUFBTAB1

             sHTML = sHTML & "</td>"
             sHTML = sHTML & "</tr>"

             sHTML = sHTML & "<tr STYLE='display:none;' >"
             sHTML = sHTML & "<td  colspan=2 align=center  >"

             sHTML = sHTML & "<font class=at><select size=10  name=comp1  "
             sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)
             sHTML = sHTML & " >"


             Strsql = "Select * From CSACOMPLE  order by descri, TIPODESCRI"

             set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             do while not rs2.eof
                sHTML = sHTML & "<option value=" & rs2(0)
                if clng(rs2(0)) = clng(wtipoL) then sHTML = sHTML & " selected "
                sHTML = sHTML & " > " & rs2(3) & ":" & rs2(2) & "&nbsp;-&nbsp;" & rs2(6) & "</option>"

                rs2.movenext
             loop
             call ofv.cerrarconsulta(rs2)
             sHTML = sHTML & "</select></font>"

             sHTML = sHTML & "</td>"
             sHTML = sHTML & "</tr>"


'-------- FIN muestra link local

            END IF

         end if

         if nblk = "S" then



'''''''''''''''''''''''''''' FINAL : BOTONES DE ANULAR O VOLVER '''''''''''''''''''''''''''''''




            sHTML = sHTML & "<tr>"
            sHTML = sHTML & "<td  align=center   >"

            sHTML = sHTML & "<input  type=Button size=10%  value='Anular Link' class=bt "
            sHTML = sHTML & " onclick=" & chr(34) & "fcomp3.submit();" & chr(34) &  " >"
            sHTML = sHTML & "</td>"

            sHTML = sHTML & "<td  align=center   >"
            sHTML = sHTML & "<input  type=Button  size=10%  value='Volver' class=bt "
            sHTML = sHTML & " onclick=" & chr(34) & "fcomp5.submit();" & chr(34) &  " >"
            sHTML = sHTML & "</td>"
            sHTML = sHTML & "</tr>"

            sHTML = sHTML & "</FORM>" & chr(13)

            sHTML = sHTML & "</table><br><br>"

            sAccion =  replace(vueltaC,"../","")
            sAccion =  replace(sAccion,"_","&")
            sHTML = sHTML & "<FORM  name=fcomp5 ID=fcomp5 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
            sHTML = sHTML & "</FORM>" & chr(13)

            if bModificar then
               sAccion = "EdGenLnkCOMP.asp?"
               sAccion = sAccion & "Id=" & nId  & "&nblk=A&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
               sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
            else
               sAccion = ""
            end if
            sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
            sHTML = sHTML & "</FORM>" & chr(13)

         END IF

      end if
      call ofv.cerrarconsulta(rs)

   end if

''''''' FIN Mostrar link '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




   call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""




  sHTML = sHTML & "</center>"

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & "</body></html>"

end if
  if nblk <> "A" then response.write sHTML

%>