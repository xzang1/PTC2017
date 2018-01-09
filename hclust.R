library(openxlsx)
library(xlsx)
library(XLConnect)
library(dynamicTreeCut)
#for daisy
library(cluster)
library(clustMixType)
#Clear memory
options( java.parameters = "-Xmx1024m" )
xlcFreeMemory()

#Read xlsx and csv
xclusters <-read.xlsx("C:/Users/xzang/Desktop/merged.xlsx",sheetIndex = 1)
xclusters<-df.workbook
rm(df.workbook)
#or
xclusters <-read.xlsx("C:/Users/xzang/Desktop/clustertopseg_standardized.xlsx",sheetIndex = 1)

#Change A-Z to 1-26
colnames(df.clusters)[1] <-"Accounts"
colnames(df.clusters)[2]<-"cluster"

x<-as.factor(df.workbook$clusters)
levels(x) <- 1:length(levels(x))
as.numeric(df.clusters$clusters)
df.clusters$clusters<-x

head(df.workbook$clusters)

#merge data
head(demo)
names(demo)[1]<-"Accounts"
new_merge <-merge(x = xclusters, y = demo, by = "Accounts", all.x = TRUE)

#clear out NAs and 0s

new_merge[complete.cases(new_merge[,2:5]),]
nrow(new_merge)
head(new_merge)
#Revised
sum(is.na(xclusters$Country__c))



#Run the clustering algorithm:

for(i in 1:26){
  xframe<-as.data.frame(xclusters[xclusters[,2]==i,,drop=FALSE])
  distclust <-daisy(xframe,metric=c("euclidean","manhattan","gower"),stand=TRUE)
  xhclust <-hclust(distclust,method="complete")
  clusternumber<-cutree(xhclust,k=48)
  clusternumber<-paste(i,0,clusternumber,sep="") #change cluster number into i_number
  i=i+1
}




#Method #2: use Dynamic tree cut
for(i in 1:26){
  xframe<-as.data.frame(xclusters[xclusters[,2]==i,,drop=FALSE])
  distclust <-daisy(xframe,metric=c("euclidean","manhattan","gower"),stand=TRUE)
  xhclust <-hclust(distclust,method="complete")
  dendro<-plot(xhclust)
  clusternumber<-cutreeDynamic(dendro,cutHeight = NULL,minClusterSize = 2,method="hybrid",distM = "daisy",4,maxCoreScatter = 95)
  clusternumber<-paste(i,0,clusternumber,sep="") #change cluster number into i_number
  i=i+1
 #change cluster number into i_number
  i=i+1
}


#Clean Cluster A:
clusterA<-clusterA[!(clusterA$GlobalEmployees__c==0)|!(clusterA$GlobalRevenue__c==0),]
clusterA<-clusterA[!(clusterA$GlobalRevenue__c==0),]
clusterA<-clusterA[,-5]

#Test just cluster A:
#Cutree

distclust <-daisy(clusterA,metric=c("euclidean","gower"),stand=TRUE)
xhclust <-hclust(distclust,method="complete")
clusternumber<-cutree(xhclust,k=2000)
clusterA$clusternumber<-clusternumber
length(which(clusterA$clusternumber==8))
which(clusterA$clusternumber==8)
#Dynamic tree cut

clusterA<-xframe[xframe[,2]==1,]
distclust <-daisy(clusterA,metric="gower",stand=TRUE,type=list())
distclust<-as.matrix(distclust)
typeof(distclust)
xhclust <-hclust(distclust,method="complete")
dendro<-plot(xhclust)
clusternumber_dynamic<-cutreeDynamic(xhclust,cutHeight = 0.250001,minClusterSize = 2,method="hybrid",distM=distclust,4,maxCoreScatter = 95)
clusterA$clusternumber_dynamic <-clusternumber_dynamic
unique(clusterA$clusternumber_dynamic)
which(clusterA$clusternumber_dynamic==0)

#Then use knn to classify outliers
#delete all with 0s
row.names(clusterA)<-clusterA$Accounts
clusterA<-clusterA[,-4]
testknn<-clusterA[which(clusterA$clusternumber==0),]
trainknn<-clusterA[-which(clusterA$clustefrnumber==0),]

#Can be deleted if do it first

trainknn[,1]=(trainknn[,1]-mean(trainknn[,1]))/sd(trainknn[,1])
trainknn[,2]=(trainknn[,2]-mean(trainknn[,2]))/sd(trainknn[,2])

testknn[,1]=(testknn[,1]-mean(testknn[,1]))/sd(testknn[,1])
testknn[,2]=(testknn[,2]-mean(testknn[,2]))/sd(testknn[,2])


ytrain<-trainknn$clusternumber
ynew<-testknn$clusternumber

nearest_1 <- knn(train=trainknn, test=testknn, cl=ytrain, k=1)
nearest_2 <- knn(train=trainknn, test=testknn, cl=ytrain, k=3)
nearest_3 <- knn(train=trainknn, test=testknn, cl=ytrain, k=4)
nearest_1
nearest_2
nearest_3
data.frame(nearest_1,nearest_2,nearest_3)[1:10,]

dim(testknn)
dim(xtrain)
data.frame(ynew,nearest_17,nearest_18,nearest_19)[1:10,]
sum(ynew==nearest_17)

#test error rate
pcorr_1=100*sum(ynew==nearest_1)/length(ynew)

pcorr_1


#Use k prototype (failed)
protocluster<-kproto(clusterA,50,iter.max=5,nstart=1)

as.numeric(clusterA$clusternumber)
max(clusterA$clusternumber)


#Test k means without countries
A_means<-clusterA[,-c(4,5)]
row.names(A_means)<-A_means$Accounts
as.numeric(A_means$GlobalEmployees__c)
typeof(A_means$GlobalRevenue__c)
as.integer64(trainknn$GlobalEmployees__c)
as.integer64(trainknn$GlobalEmployees__c)

scaledA<-scale(A_means)

.Machine$integer.max+1
library(bit64)
as.integer64(.Machine$integer.max)+1

x <-kmeans(A_means,1000,iter.max = 15,nstart=1)
A_means$Aclusters<-x$cluster


#number of rows:
length(which(clusterA$clusternumber==0))



xclusters$similar<-NA

#put all accounts with the same clusternumber into one cell for each account
#under column similar
for (x in 1:nrow(xclusters)) {
  data[x,"similar"]<-as.list(xclusters[grep(data[x,"clusternumber"],xclusters$clusternumber),1])
  x=x+1
}
row.names(clusterA[clusterA[,4]==1,])

#list segments these similar accounts have (for cross sell )
xclusters$othersegments <-NA
for (x in 1:nrow(xclusters))


  
#Show the layered graphs of the selected account's longitudinal data and it's similar accounts' data






 



