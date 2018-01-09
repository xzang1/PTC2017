
#Use PAM and daisy
library(openxlsx)
library(xlsx)
library(XLConnect)
library(WriteXLS)
#for daisy
library(cluster)
library(clustMixType)


#MERGE:
new_merge <-merge(x = df.workbook, y = xclusters, by = "Accounts", all.x = TRUE)
#delete NAs
new_merge<-new_merge[-which(is.na(new_merge$Global.Revenue)==TRUE),]
new_merge<-new_merge[-which(is.na(new_merge$Global.Employees)==TRUE),]
new_merge<-new_merge[-which(duplicated(new_merge$Accounts)==TRUE),]
summary(new_merge$cluster)

#standardize it
new_merge[,3]=(new_merge[,3]-mean(new_merge[,3]))/sd(new_merge[,3])
new_merge[,5]=(new_merge[,5]-mean(new_merge[,5]))/sd(new_merge[,5])
#Pre-processing
new_merge1<-new_merge
new_merge1$cluster2<-NA
row.names(new_merge1)<-new_merge1[,1]
new_merge1<-new_merge1[,-1]

#Use PAM, and save the dissimilarity matrix
for(i in 1:26){
  xframe<-as.data.frame(new_merge1[new_merge1$cluster==i,,drop=FALSE])
  xframe<-xframe[,-6]
  distclust <-daisy(xframe,metric=c("euclidean","gower"),stand=TRUE)
  length<-ceiling(sum(new_merge1$cluster==i)/4)
  cluster2<-pam(distclust,length,diss=TRUE,medoids = NULL,stand = TRUE,keep.diss = TRUE,keep.data = FALSE)
  #save dissimilarity matrix
  diss_matrix<-as.matrix(cluster2$diss)
  diss_matrix[upper.tri(diss_matrix)]<-NA
  write.csv(diss_matrix,paste("C:/Users/xzang/Desktop/Dissimilarity matrix for 26 clusters/csv_cluster", i, "csv", sep = "."))
  #save cluster number
  cluster2<-paste(i,0,cluster2$clustering,sep="")#change cluster number into i_number
  xframe$cluster2<-cluster2
  xframe<-xframe[,6,drop=FALSE]
  new_merge1[row.names(xframe),"cluster2"]<-xframe$cluster2
  i=i+1
}

#Without dissimilarity matrix

for(i in 1:26){
  xframe<-as.data.frame(new_merge1[new_merge1$cluster==i,,drop=FALSE])
  xframe<-xframe[,-6]
  distclust <-daisy(xframe,metric=c("euclidean","gower"),stand=TRUE)
  length<-ceiling(sum(new_merge1$cluster==i)/4)
  cluster2<-pam(distclust,length,diss=TRUE,medoids = NULL,stand = TRUE,cluster.only = TRUE)
  cluster2<-paste(i,0,cluster2,sep="")#change cluster number into i_number
  xframe$cluster2<-cluster2
  xframe<-xframe[,6,drop=FALSE]
  new_merge1[row.names(xframe),"cluster2"]<-xframe$cluster2
  i=i+1
}



write.xlsx(new_merge1,"C:/Users/xzang/Desktop/FINA_clusters_percent-PAM(1).xlsx",sheetName = "Sheet1")



