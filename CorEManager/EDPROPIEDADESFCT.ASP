<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 '' session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  XFEC = request.querystring("XFEC")
  XFECACT = request.querystring("XFECACT")

'' response.write xfec & " " & xfecact & "<br>"

  vuelta = "../EDPROPIEDADES.asp"

  set cn = ofv.conectar(ofv.strconn2)




  sHTML = ofv.FormHeader( session("FormN"))
  sHTML = sHTML & "<center>"


  Strsql = "SELECT OBJPRTY.*, Estados.estado FROM OBJPRTY, Estados "
  Strsql = Strsql & "where OBJPRTY.estado = Estados.objid and codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof Then

       IF XFECACT = "N" THEN
          XFECACT = CLNG(DATE()) - clng(cdate("01-01-1900"))
       END IF 

       XFSR   = XFECACT
       XFSRD  = XFSR





  set cn2 = ofv.conectar(ofv.strconn0)

'''''''''
  cniv = request.querystring("Cniv")
  cnom = request.querystring("Cnom")
  cx = request.querystring("Cx")
  cA = request.querystring("CA")

  if cniv = "" then 
     cniv = session("ENC7")
     cnom = session("ENC8")
     cx = session("ENC9")
   '  cA = session("ENCA")
  else
     session("ENC7") = cniv
     session("ENC8") = cnom
     cx = cdbl(cx)
     session("ENC9") = cx
     session("ENCA") = cA
  end if






  posi = 0
    sForm = "form" & posi
       sclase = "CC" 


      Celda = ofv.GenCelda("","VT colspan=2 ","","","","","","SELECCIONAR FECHA " & XDESC,"si")
      Celda = Celda & ofv.GenCelda("","ut","","","","","","ACTUAL:","si")
     
    if clng(XFSR) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput1)
       sInput1 = Fechador(XFSR)
    END IF

      Celda = Celda & ofv.gencelda("","ut","","","","","",sInput1,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")




  ''  filas = filas & ofv.GenTabla("","","","50%","","0","","0","0","","",Fila,"si")

    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())
    smes = datepart("m",date())

    SMESC = SMES

    if not isnull(XFSR) then

       if clng(XFSR) > 0 then

          sfecc = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
          smesc = datepart("m",sfecc)
          sanio = datepart("yyyy",sfecc)
       end if 


    end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")
       sanioc = request.form("saniox")

       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC

       if isnull(sanioc) then  sanioc = sanio
       if len(sanioc) = 0 then  sanioc = sanio

      
    else
       sanioc = sanio

    end if  

    if smesc = 0 then smesc = smes

    session("ENCA") = smesc






    dim Aanios(9)

    for ian = 0 to 9
        Aanios(9 - ian) = (sanio + 3) - ian        
    next 


    combanio = ofv.GenerarCombo3(sForm & "A", Aanios, "saniox", sanioc)



    
      Celda = ofv.GenCelda("","ut","","5%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     sInput = ofv.GenerarCombo(sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     SINPUT3 = ""




     Celda = Celda & ofv.GenCelda("","ut","","5%","","","","AÑO:" & SINPUT3,"si")
     
     sinput = combanio 
     Celda = Celda & ofv.gencelda("","ut","","45%","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "EDPROPIEDADESFCT.asp?cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc & "&XFEC=" & XFEC & "&XFECACT=" & XFECACT
    sAccion = sAccion & "&CODOBJETO=" & NCODIGO & "&TIPOBJETO=" & STIPO
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")


      Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","ut","","","","","","BLANQUEAR:","si")
     
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objprty"
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&volver=" & vuelta

           t1 = ofv.GenerarInput("bvo1","SI","submit","","10","bt","")
           t1 = t1 & ofv.GenerarInput(RS(cint(XFEC)).name,clng(cdate("01-01-1900")),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac999","",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1




    Formulario = Formulario & ofv.genrow("","","","","","","",Celda,"si")


    filas2 = ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","100%","","0","","0","0","","",Formulario,"si")


    dim afech(33,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)
    sfecd = clng(dateserial(sanioc,smesc,1))
    swd = datepart("w",dateserial(sanioc,smesc,1))
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)
    sfech = clng(dateserial(sanioc,smesc,28))


    yy = 0

    for ii = 1 to 3
        sfecha = sfech + ii
        sfecha = cdate(clng(sfecha))
        smess = datepart("m",sfecha)
        if cint(smesc) <> cint(smess) then
           exit for
        else
           yy = ii
        end if

    next  

    sfech = sfech + yy

    srest = 0
 '   srest = clng(cdate("01-01-1900"))
    topefech = sfech - sfecd 

    yy = 0



    wswd = swd

    for ii = 0 to topefech
        
        afech(ii+1,0) = wswd
        wswd = wswd + 1
        if wswd > 7 then wswd = 1        

        afech(ii+1,1) = sfecd + ii


    next


    for ii = 1 to (cint(afech(1,0)) - 1)
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    yy = 1


    for ii = ii to (ii + topefecH )
        afechs(ii,0) = yy
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next


    for ii = ii to 42
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","ttr style='border-style:outset;border-width:2;'","CENTER","10%","","","","S","si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    CELDA = ""
    yy = 0

    for ii = 1 to 43

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 42 then exit for  

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
           scvale = "N" 
        else
           sclase = "TT style='border-style:outset;border-width:2;'" 
           scvale = "S" 
        end if 

        if afechs(ii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","uT","CENTER","10%","","","",sinput,"si")
        else
           sinput1 = afechs(ii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

           if  scvale = "S" then
               sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objprty"
               sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&volver=" & vuelta

               t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10","bt","")
               t1 = t1 & ofv.GenerarInput(RS(cint(XFEC)).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
               t1 = ofv.GenCelda("","uT","center","10%","","","",t1,"si")
               t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 

           else
               saccion = ""


               t1 = ofv.GenerarInput("bvo1",sInput1,"button","","10","btr","")
               t1 = t1 & ofv.GenerarInput(RS(cint(XFEC)).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
               t1 = ofv.GenCelda("","uT","center","10%","","","",t1,"si")
               t1 = ofv.GenForm("","","","","post","",t1,"si") 

           end if  


           Celda = Celda & t1

        end if


    next

 

   call ofv.cerrarconn(cn2)

'''
   filas2 = "<br>" & filas2 & "<br>" & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",FILA,"si")

   sHTML = sHTML & filas2



  sHTML = sHTML & "</center>"

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  sHTML = sHTML & "</body></html>"

  END IF

  call ofv.cerrarconn(cn)
end if

response.write sHTML

Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>