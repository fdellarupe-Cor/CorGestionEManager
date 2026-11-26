<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nCodigo = request.querystring("Codobjeto")
sTipo   = request.querystring("Tipobjeto")
sEst    = request.querystring("Tipo")
detcon  = request.querystring("detcon")
ID      = request.querystring("ID")
'ID2 = request.querystring("ID2")
''volver  = "../edestrprtycont.asp?Volver=" & request.querystring("volver") & "&Codobjeto="
''volver  = volver & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & detcon & "&Tipo=" & sEst

volver  = "../edestrprtyMST.asp?Volver=" & request.querystring("volver") & "&Codobjeto="
volver  = volver & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & detcon & "&Tipo=" & sEst ' & "&ID2=" & ID2

set cn = ofv.conectar(ofv.strconn2)

select case UCase(sEst)
  case "TITULO"
    sDettipo = "A"
  case "SUBTITULO1"
    sDettipo = "B"
  case "SUBTITULO2"
    sDettipo = "C"
  case "SUBTITULO3"
    sDettipo = "D"
  case "SUBTITULO4"
    sDettipo = "E"
  case "SUBTITULO5"
    sDettipo = "F"
  case "TEXTO"
    sDettipo = "G"
end select

dim aFName(17)
dim aFValue(17)

aFname(0) = "detfont"
aFname(1) = "detfsize"
aFname(2) = "detfcolor"
aFname(3) = "detbgcolor"
aFname(4) = "detborde"
aFname(5) = "detftipo"
aFname(6) = "detsang"
aFname(7) = "detindent1"
aFname(8) = "detindent2"
aFname(9) = "detvin"
aFname(10) = "detvinfont"
aFname(11) = "detvintipo"
aFname(12) = "detvinindent"
aFname(13) = "detnumer"
aFname(14) = "detnumertipo"
aFname(15) = "detnumerclase"
aFname(16) = "detalign"
aFname(17) = "dettextalign"


for each yy in request.form
x = lcase(yy)
  if left(x,len(x)-1) = "detftipo" then
    sValue = request.form(x)
    aFValue(5) = aFValue(5) + sValue
  else
    if request.form(x) = "nsnc" then
      sValue = "''"
    else
      sValue = "'" & replace(Request.Form(x),"_"," ") & "'"
    end if
    for i = 0 to 17
      if aFName(i) = x then
        aFValue(i) = sValue
        exit for
      end if
    next
  end if
next

'''if aFValue(5) <> "" then aFValue(5) = "'" & aFValue(5) & "'"
aFValue(5) = "'" & aFValue(5) & "'"



if ID = "" and stipo <> "SYSTEM" then
  strsql = "Insert Into Objestprty (Codobjeto,Tipobjeto,detcon,dettipo,"
  for i = 0 to 16
    strsql = strsql & aFname(i) & ","
  next
  strsql = strsql & aFname(17) & ") Values (" & nCodigo & ",'" & sTipo & "'," & detcon & ",'"
  strsql = strsql & sDettipo & "',"
  for i = 0 to 16
    if i = 14 then 
      strsql = strsql & "'1.',"
    elseif i = 15 then 
      strsql = strsql & "'Simple',"
    else
      if aFValue(i) = "" then
        strsql = strsql & "'',"
      else
        strsql = strsql & aFvalue(i) & ","
      end if
    end if
  next
  strsql = strsql & aFvalue(17) & ")"
else
  if stipo = "SYSTEM" then
     strsql = "Update SYSESTPRTY Set " 
  else
     strsql = "Update Objestprty Set "
  end if


  xsep = "" 

  for i = 0 to 17
      if aFValue(i) <> "" then
         strsql = strsql & xsep & aFname(i) & "= " & aFvalue(i)
         xsep = ","
      end if


  next


  strsql = strsql &  " Where "

  if stipo <> "SYSTEM" then 
     strsql = strsql & " Codobjeto = " & nCodigo
     strsql = strsql & " And Tipobjeto = '" & sTipo & "' And Detcon = " & Detcon
     strsql = strsql & " And Dettipo = '" & sDettipo & "'"
  else
     strsql = strsql & " Dettipo = '" & sDettipo & "' "
  end if

end if

'response.write strsql
call ofv.crearconsultaEx(StrSql,cn,1,volver)
if cn.errors.count = 0 then
'response.write volver
 response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconn(cn)

%>