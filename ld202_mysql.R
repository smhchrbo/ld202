setwd("~/R/rwd")
dbcon <- read.csv("db.csv",header = T,stringsAsFactors = F)


setwd("~/R/rwd/ld202")
if(require(RMySQL)&require(sqldf)&require(xlsx)&require(reshape)){
  conn <- dbConnect(RMySQL::MySQL(), host = dbcon$host,
                    user = dbcon$user, 
                    password = dbcon$password)
  
  #2016
  year <- "2016"
  stmt1<-paste("select * from v_ff where ffy like '",year,"%'",sep="")
  stmt2<-paste("select * from v_人员进出查询 where `进出日期` like '",year,"%'",sep="")
  dbGetQuery(conn,"SET NAMES UTF8")
  dbGetQuery(conn,paste("USE",dbcon$db,sep=" "))
  ld<-dbGetQuery(conn,stmt1) 
  dbGetQuery(conn,paste("USE",dbcon$db2,sep=" "))
  ld.ry<-dbGetQuery(conn,stmt2) 
  str(dbDisconnect(conn))
}

#避免sqldf与RMySQL冲突，此处先卸载MySQL包
detach("package:RMySQL", unload=TRUE)

##
if(1){

#人员类别、岗位状态、发放月份、发放项目,合计千元
ld1<-sqldf("select ry,zt,ffy,lb,round(sum(je)/1000,1) as s from ld group by ry,zt,ffy,lb ")
ld1.cast<-cast(ld1,RY+zt+LB~FFY,value = "s")
#人员类别、岗位状态、发放月份，合计千元
ld2<-sqldf("select ry,zt,ffy,round(sum(je)/1000,1) as s from ld group by ry,zt,ffy")
ld2.cast<-cast(ld2,RY+zt~FFY,value = "s")
#计算工资发放人数
ld3<-ld[ld$LB=="工资应发数",c("FFY","RY","zt","LB")]
ld3.cast<-cast(ld3,RY+zt~FFY,value = "LB",length)
#写入excel
system.time(write.xlsx2(ld1.cast,paste("ld202-",year,".xlsx",sep=""),sheetName='ld1detail',row.names=FALSE))
system.time(write.xlsx2(ld2.cast,paste("ld202-",year,".xlsx",sep=""),sheetName='ld2aggregate',append=TRUE,row.names=FALSE))
system.time(write.xlsx2(ld3.cast,paste("ld202-",year,".xlsx",sep=""),sheetName='ld3numofemp',append=TRUE,row.names=FALSE))
system.time(write.xlsx2(ld.ry,paste("ld202-",year,".xlsx",sep=""),sheetName='ry',append=TRUE,row.names=FALSE))

}
