#Step 2 file: also contains some testing
library(openxlsx)
library(xlsx)
library(XLConnect)
library(dynamicTreeCut)
#for daisy
library(cluster)
library(clustMixType)
#Dynamic tree cut
library(dynamicTreeCut)

xclusters <-read.xlsx("C:/Users/xzang/Desktop/Demographic data with clusters1 include 0s.xlsx",sheetIndex = 1)
head(xclusters)
summary(xclusters)
xclusters<-xclusters[,-2]
colnames(xclusters)[1]<-"Accounts"

#Clear memory
options( java.parameters = "-Xmx8000m" )
xlcFreeMemory()

#Change A-Z to 1-26
colnames(df.workbook)[1] <-"Accounts"
colnames(df.workbook)[2]<-"cluster"

x<-as.factor(df.workbook$cluster)
levels(x) <- 1:length(levels(x))
as.numeric(df.workbook$cluster)
df.workbook$cluster<-x

#Clean/standardize the demographic data:

xclusters[,2]=(xclusters[,2]-mean(xclusters[,2]))/sd(xclusters[,2])
xclusters[,4]=(xclusters[,4]-mean(xclusters[,4]))/sd(xclusters[,4])
summary(xclusters)

#MERGE:
new_merge <-merge(x = df.workbook, y = xclusters, by = "Accounts")
#delete 0s,remuve duplicates
new_merge<-new_merge[-which(new_merge$Global.Revenue==0),]
new_merge<-new_merge[-which(new_merge$Global.Employees==0),]
new_merge<-new_merge[-which(duplicated(new_merge$Accounts)==TRUE),]

#Real start (cutree)

new_merge1<-new_merge
new_merge1$cluster2<-NA
row.names(new_merge1)<-new_merge1[,1]
new_merge1<-new_merge1[,-1]

for(i in 1:26){
  xframe<-as.data.frame(new_merge1[new_merge1$cluster==i,,drop=FALSE])
  xframe<-xframe[,-6]
  distclust <-daisy(xframe,metric=c("euclidean","gower"),stand=TRUE)
  xhclust <-hclust(distclust,method="complete")
  length<-ceiling(sum(new_merge1$cluster==i)/5)
  cluster2<-cutree(xhclust,k=length)
  cluster2<-paste(i,0,cluster2,sep="")#change cluster number into i_number
  xframe$cluster2<-cluster2
  xframe<-xframe[,6,drop=FALSE]
  new_merge1[row.names(xframe),"cluster2"]<-xframe$cluster2
  i=i+1
}


write.xlsx(new_merge1,"C:/Users/xzang/Desktop/FINA_clusters_percent5_0.2.xlsx",sheetName = "Sheet1")


#Start: dynamicTreeCut

new_merge1<-new_merge
new_merge1$cluster2<-NA
new_merge1<-new_merge[-which(duplicated(new_merge$Accounts)==TRUE),]
row.names(new_merge1)<-new_merge1[,1]
new_merge1<-new_merge1[,-1]


for(i in 1:26){
  xframe<-as.data.frame(new_merge1[new_merge1$cluster==i,,drop=FALSE])
  xframe<-xframe[,-6]
  distclust <-daisy(xframe,metric=c("euclidean","gower"),stand=TRUE)
  xhclust <-hclust(distclust,method="complete")
  distclust<-as.matrix(distclust)
  cluster2<-cutreeDynamic(xhclust,cutHeight = 0.2,minClusterSize = 2,method="hybrid",distM=distclust,4,maxCoreScatter = 96)
  cluster2<-paste(i,0,cluster2,sep="")#change cluster number into i_number
  xframe$cluster2<-cluster2
  xframe<-xframe[,6,drop=FALSE]
  new_merge1[row.names(xframe),"cluster2"]<-xframe$cluster2
  i=i+1
}


summary(new_merge1$cluster2)

