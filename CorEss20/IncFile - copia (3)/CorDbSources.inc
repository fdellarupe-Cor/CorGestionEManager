<%


 session("TIPOCON")		   =""
 select case SESSION("FILETIPO")
  case "L"

   session ("DirUpload")=  Request.ServerVariables("PATH_TRANSLATED")
   session ("DirUpload")=left(session ("DirUpload"),instrRev(session ("DirUpload"),"\",-1))

  case "R"

  session ("DirUpload")="http://server-20:80/coresv3/Fileserver/"

end select


ofv.ATrusted = "N"

ofv.ADbUid = "sadmin"
ofv.ADbPwd = "781784763767752745723"


'------ SQL-Server Dbase --------

ofv.strconn0="Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=dbcordatv48_UNIF485"

'Provider="SQLNCLI11;Server=cor81;Database=gpdat;DataTypeCompatibility=80;MARS Connection=True;"

ofv.strconn1 = ofv.strconn0
ofv.strconn2 = ofv.strconn0

ofv.strconn3 = ofv.strconn0
ofv.strconn4 = "dsn=Dbcordat3sql"


ofv.strconn5 ="Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=Dbcorrqtv48_UNIF485n"


ofv.strconn6=ofv.strconn5


ofv.strconn7 =ofv.strconn2
strconn8 ="PROVIDER=SQLOLEDB;Data Source=server-40;Database=dbsearchv48_unif"


%>