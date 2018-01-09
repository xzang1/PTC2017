library("kml")
data("pregnandiol", package = "kml3d")
head(pregnandiol)

install.packages("C:/Users/xzang/Downloads/kml3d_2.4.tar.gz", repos = NULL, type="source")

#TEST TEST TEST
library("kml3d")
data("pregnandiol", package = "kml3d")
head(pregnandiol)
cldPreg <- cld3d(pregnandiol, timeInData = list(.........)
typeof(cldPreg)
kml3d(cldPreg,nbClusters=2:5, nbRedrawing=10)
xll()
slots()
setClass("cldPreg3d",representation(name="character", y="list"))
x<-new("cldPreg3d",name=cldPreg@idAll,y=cldPreg@c9)
typeof(x)
setAs("cldPreg3d", "data.frame", function(from) data.frame(name=from@name, y=from@y))

#method 2
choice(cldPreg)
x11()
pregnandiol$clusters <-getClusters(cldPreg,4)
pregnandiol$clusters
setClass("cldPreg3d",representation(name="character", y="integer"))
x<-new("cldPreg3d",name=cldPreg@idAll,y=pregnandiol$clusters)
df=data.frame(cldPreg@idAll,pregnandiol$clusters)

as.data.frame.integer(pregnandiol$clusters)

library(xlsx)
write.xlsx(df, "c:/Users/xzang/Desktop/cldPreg.xlsx",sheetName="Sheet1")
Sys.setenv("R_ZIPCMD"="C:/Rtools/bin/zip.exe")

rm(Hugh_edit2)


#Real start
library(openxlsx)
Workbook_Final <-read.xlsx("C:/Users/xzang/Desktop/calibrated_cumulative_seg4_percent.xlsx",1)
head(Workbook_Final)
Workbook_Final$sum<-rowSums(Workbook_Final[,c(2:258)])
Workbook_Final<-Workbook_Final[-which(Workbook_Final$sum==0),]
library(xlsx)
write.xlsx(workbook_Final,"c:/Users/xzang/Desktop/calibrated_cumulative_seg4 real.xlsx")


typeof(Workbook_Final)
as.data.frame(Workbook_Final)
is.data.frame(Workbook_Final)

Workbook_Final <-Workbook_Final[!is.na(Workbook_Final$Accounts),]
sum(is.na(Workbook_Final$Accounts))
install.packages("C:/Users/xzang/Downloads/kml3d_2.4.tar.gz", repos = NULL, type="source")
library("kml3d")
 
cldWorkbook <- cld3d(Workbook_Final,idAll = Workbook_Final$Accounts, timeInData = list(c(2:32),+c(33:64),+c(65:96),+c(97:100),+c(101:132),+c(133:164),+c(165:196),+c(197:207),+c(208:239),+c(240:271)))
kml3d(cldWorkbook,nbClusters=26)



Workbook_Final$clusters <-getClusters(cldWorkbook,26)
Workbook_Final$clusters

summary(Workbook_Final$clusters)

df.workbook=data.frame(Workbook_Final$Accounts,Workbook_Final$clusters)
head(df.workbook)

Sys.setenv("R_ZIPCMD"="C:/Rtools/bin/zip.exe")
library(xlsx)
write.xlsx(df.workbook, "c:/Users/xzang/Desktop/Cluster1 result.xlsx",sheetName="Sheet1")


library(XLConnect)
options( java.parameters = "-Xmx8000m" )
xlcFreeMemory()
write.xlsx(Workbook_Final,"c:/Users/xzang/Desktop/Workbook_Final1.xlsx",sheetName="Sheet2")
#Exame kml3d()

getAnywhere(kml3d)


