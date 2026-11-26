<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

dim pepe(15)
dim tablae(15)
dim tablas(15)

ii = 1
comodin = 0

xmax = 15
xmin = 1



for ii = 1 to xmax

  pepe(ii) = (xmax + 1) - ii

 'response.write pepe(ii) & " pepe " & "<br>"
 'response.write ii & " ii " & "<br>"

next

  for x = 1 to 15
      tablae(x) =  pepe(x)
  next

  x = 1

  response.write tablae(x) & " tablae " & "<br>"

  for x = 1 to xmax
  '' response.write tablae(x) & " tablae -36- " & "<br>"
  '' response.write tablae(x + 1) & " tablae + 1 " & "<br>"
  if x + 1 <= 15 then
  if tablae(x) < tablae(x + 1) then
     tablas(x) = tablae(x)
     response.write tablas(x) & " tablas -38- " & "<br>"

  else
     comodin = tablae(x)
     response.write comodin & " comodin -42- " & "<br>"
     if comodin > tablae(x + 1) then
         tablas(x) = tablae(x + 1)
         response.write tablas(x) & " tablas(>) " & "<br>"
     else
         tablas(x) = comodin
         response.write tablas(x) & " tablas(<) " & "<br>"
     end if

   end if

 end if

 Next

 sinput = tablas
 celdas = celdas & ofv.GenCelda("","VT","","20","","","",sInput,"si")
 filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
 sHTML = sHTML & filas

 sHTML = sHTML  & "</body></html>"

 response.write "termino"

 Response.Write sHTML

%>