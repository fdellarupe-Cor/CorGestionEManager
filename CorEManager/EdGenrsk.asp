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
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     select case ucase(ntabla)

            case "OBJRSKREL"

                  wproced = nid
                  wID     = nid
                  wlink   = rs(6)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(7)
                  wDAT    = 7

            case "OBJACCION"

                  wproced = rs(3)
                  wID     = rs(18)
                  wlink   = rs(19)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(20)
                  wDAT    = 20



            case "OBJESQUEMA2"

                  wproced = RS(3)
                  wID     = nid
                  wlink   = rs(9)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(10)
                  wDAT    = 10


            case "OBJCONTROL2"

                  wproced = rs(3)
                  wID     = nid
                  wlink   = rs(8)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(9)
                  wDAT    = 9

            case "OBJESTDET"

                  wproced = nid
                  wID     = nid
                  wlink   = rs(32)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(33)
                  wDAT    = 33

      END SELECT




  end if

  call ofv.cerrarconsulta(rs)


''''''''''''''''''  COMIENZO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  sHTML = ofv.FORMheader("")

  sHTML = sHTML & "<SCRIPT language='javascript'  >" & chr(13) & chr(13)
  sHTML = sHTML & "    var myGlos;                           " & chr(13)
  sHTML = sHTML & "</SCRIPT>" & chr(13) & chr(13)
  sHTML = sHTML & "<SCRIPT language='javascript' for='LINKTXT' event='onmouseout' >" & chr(13) & chr(13)

  sHTML = sHTML & " if (document.selection.type=='Text') {;"

  sHTML = sHTML & "     myGlos = document.selection.createRange().text;"

  sHTML = sHTML & "     document.selection.createRange().text = '||||$' + myGlos + '$||||' ;"
  sHTML = sHTML & "     document.all.TDLINKTXT.style.display = 'none' ;"
  sHTML = sHTML & "     document.all.TD2LINKTXT.style.display = 'block' ;"
  sHTML = sHTML & "     document.all.TD3LINKTXT.style.display = 'block' ;"
  sHTML = sHTML & "     document.all.TD4LINKTXT.style.display = 'block' ;"
  sHTML = sHTML & "     document.all.LINKTXT2.innerHTML = 'El texto asociado es: ' + myGlos;"
  sHTML = sHTML & "     document.all[xnomdato].value = document.all.LINKTXT.innerHTML;"
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
        strsql = strsql & wlink & ",'" & wlinkT & "','','','LOR','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
        strsql = strsql & ")"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then
           wlnkid = rs(0)
        else
           wlnkid = 0
           wlink = 0
        end if
        call ofv.cerrarconsulta(rs)

        if clng(wlnkid) > 0 then
           strsql = "update " & ntabla & " set rellink = " & wlnkid
           strsql = strsql & " where  objid = " & nID
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if
     end if

     nblk = "S"
  end if

''''''' FIN inserta link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''' inserta link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if nblk = "II" then

     wlink = request.form("comp1")


        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
        strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',"
        strsql = strsql & "'" & wproced & "','" & wId & "',"
        strsql = strsql & "0,'','" & wlink & "','" & wlink & "','IOR')"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then
           wlnkid = rs(0)
        else
           wlnkid = 0
           wlink = 0
        end if
        call ofv.cerrarconsulta(rs)

        if clng(wlnkid) > 0 then
           strsql = "update " & ntabla & " set rellink = " & wlnkid
           strsql = strsql & " where  objid = " & nID
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if

     wlink = wlnkid
     nblk = "S"
  end if

''''''' FIN inserta link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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


     sHTML = sHTML & GenBuscaObj(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0")
     sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 ","","60%","",0,0,0,0,"","","","no")

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")
     sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "</FORM>" & chr(13)

     if bModificar then
        sAccion = "EdGenRsk.asp?"
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


''''''' agrega link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if nblk = "AI" then

     sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  colspan=2 class=ttt align=center  >"
     sHTML = sHTML & "Agregar&nbsp;Link&nbsp;a&nbsp;Internet "
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "</tr>"
     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")
     sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "</FORM>" & chr(13)

     if bModificar then
        sAccion = "EdGenRsk.asp?"
        sAccion = sAccion & "Id=" & nId  & "&nblk=II&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
        sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
     else
        sAccion = ""
     end if
     sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "<tr height=50 >"
     sHTML = sHTML & "<td  colspan=2 align=center  >"
     sHTML = sHTML & ofv.GenerarInput("comp1","www.","agregar","","100","dt","sb")
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

''''''' FIN agrega link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''' Alta 1era vez ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if clng(wlink) = 0 and nblk = "S" then

     sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=5  ","","60%","",0,0,0,0,"","","","no")
     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  colspan=3 class=ttt align=center >"
     sHTML = sHTML & "Agregar&nbsp;Link "
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "</tr>"
     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  COLSPAN=3 align=center   >&nbsp;</TD>"
     sHTML = sHTML & "</tr>"
     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  align=center   >"

     if bModificar then
        sAccion = "EdGenRsk.asp?"
        sAccion = sAccion & "Id=" & nId  & "&nblk=AL&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
        sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
     else
        sAccion = ""
     end if

     sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='AGREGAR LINK A OTRO DOCUMENTO' class=bt >"
     sHTML = sHTML & "</FORM>" & chr(13)
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "<td  align=center   >"

     if bModificar then
        sAccion = "EdGenRsk.asp?"
        sAccion = sAccion & "Id=" & nId  & "&nblk=AI&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
        sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
     else
        sAccion = ""
     end if

     sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='AGREGAR LINK A INTERNET' class=bt >"
     sHTML = sHTML & "</FORM>" & chr(13)
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "<td  align=center   >"

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")
     sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
     sHTML = sHTML & "<input  type=submit  size=10%  value='CANCELAR' class=bt >"
     sHTML = sHTML & "</FORM>" & chr(13)
     sHTML = sHTML & "</td>"
     sHTML = sHTML & "</tr>"

     sHTML = sHTML & "<tr>"
     sHTML = sHTML & "<td  COLSPAN=3 align=center   >&nbsp;</TD>"
     sHTML = sHTML & "</tr>"
     sHTML = sHTML & "</table><br><br>"


     nblk = "NO"


  end if

''''''' FIN Alta 1era vez ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''' modicar link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "ML" then

     wlinkD = request.form("comp2") & ""

        strsql = "update OBJLNK set LurlDesc = '" & wlinkD & "' "
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
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
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

     end if

     nblk = "S"
  end if


''''''' FIN modicar link local D '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''' modicar link internet ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "MI" then

     wlink = request.form("comp1")

   '''  wlinkD = request.form("comp2")



        strsql = "update OBJLNK set Lurl = '" & wlink & "' "
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



     nblk = "S"
  end if
''''''' FIN modicar link internet ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''' modicar link internet D''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "MID" then

   '''  wlink = request.form("comp1")

     wlinkD = request.form("comp2")



        strsql = "update OBJLNK set LurlDesc = '" & wlinkD & "' "
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



     nblk = "S"
  end if
''''''' FIN modicar link internet D ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''' anular link ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "A" then

 '    response.write "anular link: " & "<br>"

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")

     strsql = "delete from OBJLNK "
     strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & wId & "' "
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

 '    response.write "anular link: " & strsql & "<br>"

     strsql = "update " & ntabla & " set rellink = 0, reflink = '' "
     strsql = strsql & " where  objid = " & nID
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

     response.redirect sAccion

     nblk = "NO"

  end if
''''''' FIN anular link ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''' Mostrar link '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if nblk = "S" OR nblk = "SD" then

     Strsql = "SELECT * From " & ntabla & " where objid = " & nID
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then

       select case ucase(ntabla)


            case "OBJRSKREL"

                  wproced = nid
                  wID     = nid
                  wlink   = rs(6)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(7)
                  wDAT    = 7



            case "OBJACCION"

                  wproced = rs(3)
                  wID     = rs(18)
                  wlink   = rs(19)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(20)
                  wDAT    = 20


            case "OBJESQUEMA2"

                  wproced = RS(3)
                  wID     = nAccion
                  wlink   = rs(9)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(10)
                  wDAT    = 10


            case "OBJCONTROL2"

                  wproced = rs(3)
                  wID     = nAccion
                  wlink   = rs(8)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(9)
                  wDAT    = 9

            case "OBJESTDET"

                  wproced = nid
                  wID     = nid
                  wlink   = rs(32)
                  if isnull(wlink) then wlink = 0
                  wtxtl   = rs(33)
                  wDAT    = 33


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
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs1.eof then
          wtipoL   = clng(rs1(6))
          wtipOBJ  = rs1(7)
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

 ' response.write  wtipoL &  wUrl & wUrlDesc & wtipo & "<br>"

'''----'''''' LINK A INTERNET O EXTERNO '''''''''''''''''''''''''''''''''''''''


       if trim(wtipo) = "I" or trim(wtipo) = "IOR" then

'-------- muestra link Internet

          sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

          sHTML = sHTML & "<tr>"
          sHTML = sHTML & "<td  colspan=2 class=ttt >"
          sHTML = sHTML & "Link&nbsp;a&nbsp;direccion"
          sHTML = sHTML & "</td>"
          sHTML = sHTML & "</tr>"

         if bModificar then
            sAccion = "EdGenRsk.asp?"
            sAccion = sAccion & "Id=" & nId  & "&nblk=MI&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
            sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
         else
            sAccion = ""
         end if

         sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

         sHTML = sHTML & "<tr height=40 >"
         sHTML = sHTML & "<td  colspan=2 class=uti  >"
         sHTML = sHTML & ofv.GenerarInput("comp1",WUrl,"textsb","fcomp1","100","dt","sb")
         sHTML = sHTML & "</td>"
         sHTML = sHTML & "</tr>"



         sHTML = sHTML & "</FORM>" & chr(13)


'-------- FIN muestra link Internet


      else

'-------- muestra link local



         if nblk = "SD" then


            sHTML = sHTML & GenBuscaObj(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0")
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

            sAccion = "EdGenRsk.asp?"
            sAccion = sAccion & "Id=" & nId  & "&nblk=S&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
            sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC

            sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
            sHTML = sHTML & "</FORM>" & chr(13)

            if bModificar and XGUFCOBJ > 0  then
               sAccion = "EdGenRsk.asp?"
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
            sHTML = sHTML & "Link&nbsp;a&nbsp;documento"
            sHTML = sHTML & "</td>"
            sHTML = sHTML & "</tr>"


            sHTML = sHTML & "<tr>"
            sHTML = sHTML & "<td  colspan=2 align=center  >"





            XMST = "0"

            IF wtipoL <> 0 THEN 

               XMST = "1"

                CALL GETLISTOBJ("E",wtipoL,wtipOBJ)
                XGUFCOBJKEY = XGUTABLISKEY
                XGUFCOBJDAT = XGUTABLISDAT 


                XGUFBTDF0 = XGUFBTDF0 & "<TR  "
                XGUFBTDF0 = XGUFBTDF0 & " >"

                XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT   >"
                XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
                XGUFBTDF0 = XGUFBTDF0 & "</TR>"

                IF WNV1 <> "0" THEN 

                   XMST = "2"


                   CALL GETLISTSLP("E",wtipoL,wtipOBJ,WNV1)
                   XGUFCOBJKEY = XGUTABLISKEY
                   XGUFCOBJDAT = XGUTABLISDAT 


                   XGUFBTDF1 = XGUFBTDF1 & "<TR  "
                   XGUFBTDF1 = XGUFBTDF1 & " >"

                   XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT   >"
                   XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
                   XGUFBTDF1 = XGUFBTDF1 & "</TR>"

                   IF WNV2 <> "0" THEN 

                      XMST = "3"

                      CALL GETLISTNV2("E",wtipoL,wtipOBJ,WNV1,WNV2)
                      XGUFCOBJKEY = XGUTABLISKEY
                      XGUFCOBJDAT = XGUTABLISDAT 


                      XGUFBTDF2 = XGUFBTDF2 & "<TR  "
                      XGUFBTDF2 = XGUFBTDF2 & " >"

                      XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT   >"
                      XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
                      XGUFBTDF2 = XGUFBTDF2 & "</TR>"

                      IF WNV3 <> "0" THEN 

                         XMST = "4"

                         CALL GETLISTNV3("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3)
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


                sAccion = "EdGenRsk.asp?"
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

             Strsql = "Select * From csGenObjetos where estado = 3  OR ESTADO = 1 order by xdescri"
             set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             do while not rs2.eof
                sHTML = sHTML & "<option value=" & rs2(0)
                if clng(rs2(0)) = clng(wtipoL) then sHTML = sHTML & " selected "
                sHTML = sHTML & " > " & rs2(1) & "</option>"

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


'''''''''''''''''''''''''''''''''''''''''''''''''''''' BLOQUE ASOCIACION AL TEXTO '''''''''''''''''''''''''''''''''''''''''''


'----------' agregado: asociado a texto/sin asociar al texto '''' 

             Strsql = "SELECT * From objlnk where objid = " & wlink
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rs1.eof then

   '''  RESPONSE.WRITE STRSQL & "<BR>" & wasotxt & "<BR>"


                wasotxt     = rs1(14) & ""
                wasotxtstr  = CLNG("0" & rs1(15))
                wasotxtcan  = CLNG("0" & rs1(16))
                wasotxtdat  = rs1(17) & ""


                select case ucase(ntabla)


                       case "OBJRSKREL"


                             wtxtl2   = rs(3)
                             wDAT2    = 3


 
                       case "OBJACCION"


                             wtxtl2   = rs(11)
                             wDAT2    = 11


                       case "OBJESQUEMA2"


                             wtxtl2   = rs(6)
                             wDAT2    = 6


                       case "OBJCONTROL2"


                             wtxtl2   = rs(4)
                             wDAT2    = 4

                       case "OBJESTDET"


                             wtxtl2   = rs(8)
                             wDAT2    = 8


                 END SELECT

                 IF wasotxt = "" THEN wasotxt = "N"

            '     sHTML = sHTML & "<tr>"
            '     sHTML = sHTML & "<td  colspan=2 class=ttt  >"
            '     sHTML = sHTML & "Presentacion&nbsp;del&nbsp;link"
            '     sHTML = sHTML & "</td>"
            '     sHTML = sHTML & "</tr>"  & chr(13)

            '     if bModificar then
            '        sAccion = "asp/ActualizarEN.asp?ID=" & WLINK & "&Tabla=OBJLNK&volver=" & volver2
            '        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
            '     else
            '        sAccion = "" 
            '     end if
            '     sHTML = sHTML & "<FORM  name=fcomp9 ID=fcomp9 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

            '     sHTML = sHTML & "<tr>"
            '     sHTML = sHTML & "<td  colspan=2 class=uti  >"

            '     scomK = "S" & chr(9) & "N"
            '     scomD = "ASOCIADO AL TEXTO" & chr(9) & "SIN ASOCIAR"


            '     if bModificar then
            '        sInput = ofv.gencomboFX("fcomp9",RS1(14).NAME,wasotxt,scomK,scomD," ")
            '     else
            '        sInput = ofv.GenerarInput("",wasotxt,"readonly","","10","dt","")
            '     end if

            '     sHTML = sHTML & sInput

            '     sHTML = sHTML & ofv.GenerarInput(RS1(15).name,"0","hidden","","10","","")
            '     sHTML = sHTML & ofv.GenerarInput(RS1(16).name,"0","hidden","","10","","")
            '     sHTML = sHTML & ofv.GenerarInput(RS1(17).name,rs(WDAT2).name,"hidden","","10","","")


            '     sHTML = sHTML & "</td>"
            '     sHTML = sHTML & "</tr>"

            '     sHTML = sHTML & "</FORM>" & chr(13)


                 if wasotxt = "S" then

                    xdesde =  wasotxtstr
                    xhasta =  wasotxtcan


                    if instr(wtxtl2 & "","||||$") > 0 then

                       wtxtl2 = replace(wtxtl2,"<br>",chr(13))
                       wtxtl2 = replace(wtxtl2,"<BR>",chr(13))

                       xdesde =  instr(wtxtl2 & "","||||$")
                       xhasta =  instr(wtxtl2 & "","$||||") - xdesde - 5

                       wtxtl2 = replace(wtxtl2,"||||$","") 
                       wtxtl2 = replace(wtxtl2,"$||||","")


                       Strsql = "update " & ntabla & " set " & rs(WDAT2).name & " = '" & wtxtl2 & "'  where objid = " & nID
   ''''  response.write "<br><br><br><br>" & strsql & "<br>"
                       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

                       Strsql = "update objlnk set postxt = " & xdesde & ", canttxt = " & xhasta & "   where objid = " & wlink
   ''''  response.write "<br><br><br><br>" & strsql & "<br>"
                       call ofv.crearconsultaEx(StrSql,cn,1,parametros)


                    end if 


                    if  xdesde = 0 and xhasta = 0 then




                        if bModificar then

                           sAccion = "asp/ActualizarEN.asp?ID=" & nId & "&Tabla=" & ntabla & "&volver=" & volver2
                           sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

                        else
                           sAccion = ""
                        end if


                        sHTML = sHTML & "<tr>"
                        sHTML = sHTML & "<td  colspan=2 class=uti id='TDLINKTXT' style='display:block;' >"

                        t3b = "<DIV class='DINDC' STYLE=""height:60;overflow:'auto';"" "
                        t3b = t3b & "name='LINKTXT' id='LINKTXT' "
                        t3b = t3b & "  >"
                        xdatos = replace(wtxtl2 & "",chr(13),"<br>")
                        t3b = t3b & xdatos & "</div>"


                        sHTML = sHTML & t3b
                        sHTML = sHTML & "</td>"
                        sHTML = sHTML & "</tr>"

                        sHTML = sHTML & "<tr>"
                        sHTML = sHTML & "<td  colspan=2 class=ut id='TD2LINKTXT' style='display:none;' >"

                        t3b = "<DIV  STYLE=""height:60;overflow:'auto';"" "
                        t3b = t3b & "name='LINKTXT2' id='LINKTXT2' "
                        t3b = t3b & "  >xx"
                        t3b = t3b & "</div>"

                        sHTML = sHTML & "<script language='javascript' >"
                        sHTML = sHTML & " var xnomdato = '" & rs(WDAT2).name & "'; "
                        sHTML = sHTML & "</script>"

                        sHTML = sHTML & t3b
                        sHTML = sHTML & "<FORM style='display:none;' name=fcomp10 ID=fcomp10 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

                        sHTML = sHTML & ofv.GenerarInput(rs(WDAT2).name,rs(WDAT2),"hidden","","10","","")


                        IF UCASE(NTABLA) = "OBJACCION" THEN

                           sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
                           sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")

                        END IF


                        sHTML = sHTML & "</FORM>" & chr(13)

                        sHTML = sHTML & "</td>" 
                        sHTML = sHTML & "</tr>"


                        sHTML = sHTML & "<tr>"
                        sHTML = sHTML & "<td align='right' class=ut id='TD3LINKTXT' style='display:none;' >"
                        sHTML = sHTML & "<FORM   action=''  method=post >" & chr(13)
                        sHTML = sHTML & "<input  type=Button size=10%  value='aceptar' class=bt "
                        sHTML = sHTML & " onclick=" & chr(34) & "fcomp10.submit();" & chr(34) &  " >"
                        sHTML = sHTML & "</td>"


                        sHTML = sHTML & "</FORM>" & chr(13)
                        sHTML = sHTML & "</td>"

                        sAccion = "EdGenRsk.asp?"
                        sAccion = sAccion & "Id=" & nId  & "&nblk=" & nblk & "&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
                        sAccion = sAccion & "&Tipobjeto=" & sTipo


                        sHTML = sHTML & "<td  align='left' class=ut id='TD4LINKTXT' style='display:none;' >"
                        sHTML = sHTML & "<FORM   action='" & sAccion & "'  method=post >" & chr(13)

                        sHTML = sHTML & "<input  type=submit size=10%  value='cancelar' class=bt "
                        sHTML = sHTML & "  >"
                        sHTML = sHTML & "</td>"

                        sHTML = sHTML & "</FORM>" & chr(13)
                        sHTML = sHTML & "</td>"

                        sHTML = sHTML & "</tr>"

                     else

                        sHTML = sHTML & "<tr>"
                        sHTML = sHTML & "<td  colspan=2 class=uti id='TDLINKTXTx' style='display:block;' >"

                        t3b = "<DIV  STYLE=""height:80;overflow:'auto';border-style:solid;border-width:1;border-color:#dddddd;"" "
                        t3b = t3b & "name='LINKTXTx' id='LINKTXTx' "
                        t3b = t3b & "  >"

                        xdatos = wtxtl2
                        xdatos = left(xdatos,xdesde - 1) & "<u><b>" & mid(xdatos,xdesde,xhasta) & "</b></u>" & right(xdatos,len(xdatos) - (xdesde + xhasta - 1))
                        xdatos = replace(xdatos,chr(13),"<br>")

                        t3b = t3b & xdatos & "</div>"
                        sHTML = sHTML & t3b
 
                        sHTML = sHTML & "</td>"
                        sHTML = sHTML & "</tr>"

                        sHTML = sHTML & "<tr>"

                        if bModificar then
                           sAccion = "asp/ActualizarEN.asp?ID=" & WLINK & "&Tabla=OBJLNK&volver=" & volver2
                           sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

                        else
                           sAccion = ""
                        end if

                        sHTML = sHTML & "<td colspan=2 align='right' class=ut id='TD4LINKTXT' style='display:block;' >"
                        sHTML = sHTML & "<FORM   action='" & sAccion & "'  method=post >" & chr(13)
                        sHTML = sHTML & ofv.GenerarInput(RS1(15).name,"0","hidden","","10","","")
                        sHTML = sHTML & ofv.GenerarInput(RS1(16).name,"0","hidden","","10","","")
                        sHTML = sHTML & ofv.GenerarInput(RS1(17).name,rs(WDAT2).name,"hidden","","10","","")
                        sHTML = sHTML & "<input  type=submit size=10%  value='modificar' class=bt "
                        sHTML = sHTML & "  >"
                        sHTML = sHTML & "</td>"

                        sHTML = sHTML & "</FORM>" & chr(13)
                        sHTML = sHTML & "</td>"

                        sHTML = sHTML & "</tr>"

                     end if

                  else

                '     sHTML = sHTML & "<tr>"
                '     sHTML = sHTML & "<td  colspan=2 class=ttt  >"
                '     sHTML = sHTML & "Referencia&nbsp;para&nbsp;el&nbsp;link"
                '     sHTML = sHTML & "</td>"
                '     sHTML = sHTML & "</tr>"  & chr(13)

                '     if bModificar then
                '        sAccion = "asp/ActualizarEN.asp?ID=" & nId & "&Tabla=" & ntabla & "&volver=" & volver2
                '        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
                '     else
                '        sAccion = ""
                '     end if
                '     sHTML = sHTML & "<FORM  name=fcomp0 ID=fcomp0 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

                '     sHTML = sHTML & "<tr>"
                '     sHTML = sHTML & "<td  colspan=2 class=uti  >"

                 '    t3b = "<font class=at><textarea class=dt "
                 '    t3b = t3b & "name=" & rs(WDAT).name & " cols=90 rows=6 "
                 '    if bModificar then
                 '       t3b = t3b & " onblur=fcomp0.submit() "
                 '       t3b = t3b & " onkeypress=" & chr(34)
                 '       t3b = t3b & " if (this.value.length > 490) {alert('El texto es demasiado"
                 '       t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
                 '    else
                 '       t3b = t3b & " readonly >"
                 '    end if
                 '    xdatos = wtxtl
                 '    t3b = t3b & xdatos & "</textarea></font>"

                 '    sHTML = sHTML & t3b

                 '    IF UCASE(NTABLA) = "OBJACCION" THEN
                 '       sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
                 '       sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")
                 '    END IF

                 '    sHTML = sHTML & "</td>"
                 '    sHTML = sHTML & "</tr>"
                 '    sHTML = sHTML & "</FORM>" & chr(13)

                  end if

                  if bModificar then
                     sAccion = "EdGenRsk.asp?"
                     sAccion = sAccion & "Id=" & nId  & "&nblk="

                     if wtipo = "I" then
                        sAccion = sAccion & "MID"
                     else
                        sAccion = sAccion & "ML"
                     end if 
 

                     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
                     sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
                  else
                     sAccion = ""
                  end if
                  sHTML = sHTML & "<FORM  name=fcompn ID=fcompn action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

              '    sHTML = sHTML & "<tr>"
              '    sHTML = sHTML & "<td  colspan=2 class=ttt height=40 >"
              '    sHTML = sHTML & "Descripcion&nbsp;para&nbsp;el&nbsp;link"
              '    sHTML = sHTML & "</td>"
              '    sHTML = sHTML & "</tr>"  & chr(13)

                  sHTML = sHTML & "<tr>"
                  sHTML = sHTML & "<td  colspan=2 class=uti height=40 >"
                  sHTML = sHTML & ofv.GenerarInput("comp2",wUrlDesc,"hidden","fcompn","100","dt","sb")
                  sHTML = sHTML & "</td>"
                  sHTML = sHTML & "</tr>"

               end if
               call ofv.cerrarconsulta(rs1)

'--------------FIN ASOCIACION AL TEXTO ''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '    END IF

       '  end if



'''''''''''''''''''''''''''' FINAL : BOTONES DE ANULAR O VOLVER '''''''''''''''''''''''''''''''


     '    if nblk = "S" then


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
               sAccion = "EdGenRsk.asp?"
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


  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & "</center></body></html>"
end if
  if nblk <> "A" then response.write sHTML

%>