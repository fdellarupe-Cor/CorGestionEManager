<%



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
ofv.strconn0="Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=GP_DBCORDAT_PROD"


ofv.strconn1 = ofv.strconn0
ofv.strconn2 = ofv.strconn0

ofv.strconn3 = ofv.strconn0
ofv.strconn4 = "dsn=Dbcordat3sql"


ofv.strconn5 ="Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=GP_RQT_PROD"
ofv.strconn6=ofv.strconn5


ofv.strconn7 =ofv.strconn2
strconn8 =ofv.strconn0' "PROVIDER=SQLOLEDB;Data Source=server-40;Database=dbsearchGp"


%>