#remove all objects
rm(list = ls())

#Setting the working directory

getwd()
setwd("E:/ABHI/COURSES/EDWISOR/PROJECTS/EMPLOYEE ABSENTEISM")

#Importing the libraries
library(xlsx)
library(ggplot2)
library(gridExtra)
library(DMwR)
library(corrgram)
library(forecast)
library(outliers)
library(plyr)

#importing employee absenteeism data set
df = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1 ,header = T)

summary(df)
names(df)
str(df)
#By observing the month column of the dataset we can infer that the dataset given is from July 2007 to July 2010
#Thus, adding a "year" variable in the dataset

df[1:113, 'year'] = 2007
df[114:358, 'year'] = 2008
df[359:570, 'year'] = 2009
df[571:740, 'year'] = 2010

#Setting the year variable position adjacent to month i.e postion number 3
df = df[,c(1,2,3,22,4:21)]

#On observing the tail of our dataset we find that the month of absence is zero, which is not possible
#Thus, we'll remove the last 3 rows

df = df[1:737,]

#CATEGORICAL VARIABLES
cat_variables = c('ID',"Reason.for.absence", "Month.of.absence","Day.of.the.week", "Seasons", "Disciplinary.failure",
                     "Education", "Son", "Social.drinker", "Social.smoker", "Pet","year")

#Converting the data type of categorical variables into factor type
for(i in cat_variables)
{
        df[,i] = as.factor(df[,i])
}
str(df) #Checking to confirm of the variable conversion went successfully

#####   MISSING VALUE ANALYSIS

#Creating another dataframe for missing values
missing_data=data.frame(apply(df,2,function(x){sum(is.na(x))}))

#Renaming the column of missing_data
names(missing_data)[1] =  "missing values"

#Saving the dataset into our system
write.csv(missing_data,"Missing Values.csv")

#colnames(df)
#List of all continuous variables
cont_variables=c('Distance.from.Residence.to.Work', 'Service.time', 'Age', 'Hit.target', 'Work.load.Average.day.',
      'Transportation.expense', 'Weight', 'Height', 'Body.mass.index', 'Absenteeism.time.in.hours')


#MISSING VALUES IMPUTATION


#Checking which method is better for imputation
#deleting observation 2 of variable Height
#df[2,20]
#Actual value [2,20] = 178
#Value form mean method = 172.1448
#Value from median method = 170
#Value from Knn method = 178

#df[2,20] = NA
#Mean method of imputation
#df$Height[is.na(df$Height)] = mean(df$Height,na.rm=T) #na.rm=T will remove all NA values and calculate the mean
#df[2,20]

#df[2,20] = NA
#Median method of imputation
#df$Height[is.na(df$Height)] = median(df$Height,na.rm=T) 
#df[2,20]

#df[2,20] = NA
#Knn method of imputation
df = knnImputation(df,k=3)
#df[2,20]

#Confirming if all the missing values are imputed 
sum(is.na(df))

#Converting all the variables into integer as factor variables are affected after Knn inputation
for(i in cat_variables)
{
        df[,i] = as.integer(df[,i])
}

for(i in cont_variables)
{
        df[,i] = as.integer(df[,i])
}



####### OUTLIER ANALYSIS


for (i in 1:length(cont_variables))
{
       assign(paste0("gn",i), ggplot(aes_string(y = (cont_variables[i]), x = "Absenteeism.time.in.hours"), data = df)+ 
                     stat_boxplot(geom = "errorbar", width = 0.5) +
                    geom_boxplot(outlier.colour="blue", fill = "yellow" ,outlier.shape=12,
                                outlier.size=2, notch=FALSE) +
                  theme(legend.position="bottom")+
                 labs(y=cont_variables[i],x="TOTAL Absenteeism time in hours")+
                ggtitle(paste("Box plot for",cont_variables[i])))
}

## Grid Plot
gridExtra::grid.arrange(gn1,gn2,gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)


#Replacing Outlier Values by NA
for(i in cont_variables)
{
        outlier_val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
        df[,i][df[,i] %in% outlier_val] = NA
}

#Imputation of Outlier values
df = knnImputation(df, k = 3)

#Checking if all the outlier values are imputed
sum(is.na(df))


#Converting all the variables into integer as factor variables are affected after Knn inputation
for(i in cat_variables)
{
        df[,i] = as.integer(df[,i])
}

for(i in cont_variables)
{
        df[,i] = as.integer(df[,i])
}

#Converting the data type of categorical variables into factor type
for(i in cat_variables)
{
        df[,i] = as.factor(df[,i])
}

##################FEATURE SELECTION

#CORRELATION PLOT OF CONTINUOUS VARIABLES
corrgram(df[,cont_variables], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

#From the corrlation plot we can easily observe that Body.mass.index & weight are highly correlated
#Thus, dropping the variable Body mass index from our dataset

df = subset(df,select = -c(Body.mass.index))

#Performing one way ANOVA test for finding the corrleation between the target & categorical features

cat_variables

anova1 = aov(df$Absenteeism.time.in.hours ~ df$ID)
anova2 = aov(df$Absenteeism.time.in.hours ~ df$Reason.for.absence)
anova3 = aov(df$Absenteeism.time.in.hours ~ df$Month.of.absence)
anova4 = aov(df$Absenteeism.time.in.hours ~ df$Day.of.the.week)
anova5 = aov(df$Absenteeism.time.in.hours ~ df$Seasons)
anova6 = aov(df$Absenteeism.time.in.hours ~ df$Disciplinary.failure)
anova7 = aov(df$Absenteeism.time.in.hours ~ df$Education)
anova8 = aov(df$Absenteeism.time.in.hours ~ df$Son)
anova9 = aov(df$Absenteeism.time.in.hours ~ df$Social.drinker)
anova10 = aov(df$Absenteeism.time.in.hours ~ df$Social.smoker)
anova11 = aov(df$Absenteeism.time.in.hours ~ df$Pet)
anova12 = aov(df$Absenteeism.time.in.hours ~ df$year)

#From the above ANOVA test we conclude that 
#all the categorical variables are important and none of them can be dropped from our model


#DATA VISUALIZATION
#Univariate analysis of all continuous variables by plotting DENSITY PLOTS

gd1 = ggplot(df, aes(x = Transportation.expense)) + geom_density()
gd2 = ggplot(df, aes(x = Distance.from.Residence.to.Work)) + geom_density()
gd3 = ggplot(df, aes(x = Age)) + geom_density()
gd4 = ggplot(df, aes(x = Service.time)) + geom_density()
gd5 = ggplot(df, aes(x = Work.load.Average.day.)) + geom_density()
gd6 = ggplot(df, aes(x = Hit.target)) + geom_density()
gd7 = ggplot(df, aes(x = Weight)) + geom_density()
gd8 = ggplot(df, aes(x = Height)) + geom_density()
gd9 = ggplot(df, aes(x = Weight)) + geom_density()
gd10 = ggplot(df, aes(x = Absenteeism.time.in.hours)) + geom_density()

grid.arrange(gd1, gd2, gd3, gd4, gd5, gd6, gd7, gd8, gd9, gd10, ncol=2)

#All the continuous variables are normally distributed

#COUNT PLOTS

#Count for Reason of absence
gc1 = ggplot(df, aes(x = Reason.for.absence)) + geom_bar(stat = 'count')  +
        geom_label(stat='count',aes(label=..count..), size=5) 
gc1

#Reasons for absence according to seasons
gc2 = ggplot(df, aes(x = Seasons)) + geom_bar(stat = 'count', position = 'dodge')+
        geom_label(stat = 'count', aes(label = ..count..)) 
gc2

#Reasons for absence according to Day of the week
gc3 = ggplot(df, aes(x = Day.of.the.week)) + geom_bar(stat = 'count', position = 'dodge')+
        geom_label(stat = 'count', aes(label = ..count..))
gc3


#Reasons for absence according to Month of absence
gc4 = ggplot(df, aes(x = Month.of.absence)) + geom_bar(stat = 'count', position = 'dodge')+
        geom_label(stat = 'count', aes(label = ..count..))
gc4


###DATA ANALYSIS BY GROUPING

#Data analysis by grouping the absenteeism time in hours for disciplinary failure
disciplinary_group =  ddply(df, c( "Disciplinary.failure"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
disciplinary_group
#It is observed that after a disciplinary action is taken employee absenteeism decreases significantly


#Data analysis by grouping the absenteeism time in hours for ID
ID_group =  ddply(df, c( "ID"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
ID_group
#ID 3 has the highest absenteeism hours


#Data analysis by grouping the absenteeism time in hours for Reason for absence 
reason_group=  ddply(df, c("Reason.for.absence"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
reason_group
#Most common reasons are those which dont require medical certifications from the doctor


#Data analysis by grouping the absenteeism time in hours for Son 
son_group=  ddply(df, c("Son"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
son_group
#Employees without any children have highest absentee hours


#Data analysis by grouping the absenteeism time in hours for Social drinker
socialdrinker_group=  ddply(df, c("Social.drinker"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
socialdrinker_group
#Employees who drink socially have higher absentee hours



#Data analysis by grouping the absenteeism time in hours for Social smoker
socialsmoker_group=  ddply(df, c("Social.smoker"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
socialsmoker_group
#Interestingly Employees who smoke socially have lesser absentee hours
#However, the no. of employees who smoke socially is also less



#Data analysis by grouping the absenteeism time in hours for Seasons
seasons_group =  ddply(df, c( "Seasons"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
seasons_group
#Season 3 i.e. winter has the highest absentee hours


#Data analysis by grouping the absenteeism time in hours for Day of the week
day_group =  ddply(df, c( "Day.of.the.week"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
day_group
#Monday has the highest absentee hours


#Data analysis by grouping the absenteeism time in hours for Month of absence
month_group =  ddply(df, c( "Month.of.absence"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
month_group
#Month 3 & 7 i.e. March and July have the highest absentee hours


#########DATA MODELLING

##DECISION TREE REGRESSION
#Dividing the data into train & test
train_index = sample(1:nrow(df),0.8*nrow(df))
train_dt=df[train_index,]
test_dt=df[-train_index,]

# rpart for decision trees
#Load libraries
library(rpart)
library(MASS)
library(DMwR)
dtmodel = rpart(Absenteeism.time.in.hours~., data=train_dt, method="anova")


#Predict value for test data
pred_dt = predict(dtmodel,test_dt[,-21])

#Checking the error
regr.eval(test_dt[,21],pred_dt,stats=c('mae','rmse'))



##LINEAR REGRESSION
#Dividing the data into train & test
train_lr=df[train_index,]
test_lr=df[-train_index,]


# rpart for regression

lrmodel = rpart(Absenteeism.time.in.hours~., data=train_lr)


#Predict value for test data
pred_lr = predict(lrmodel,test_dt[,-21])

#Checking the error
regr.eval(test_dt[,21],pred_lr,stats=c('mae','rmse'))


##RANDOM FOREST
#Dividing the data into train & test
train_rf=df[train_index,]
test_rf=df[-train_index,]


# Random Forest model
library(randomForest)

rfmodel = randomForest(Absenteeism.time.in.hours~., data=train_rf,importance=TRUE,ntree=100)

#Predict value for test data
pred_rf = predict(rfmodel,test_rf[,-21])

#checking the error
regr.eval(test_rf[,21],pred_rf,stats=c('mae','rmse'))



####PREDICTIONS FOR THE YEAR 2011 USING LINEAR REGRESSION

df_month =  ddply(df, c("year", "Month.of.absence"), function(x) colSums(x[c("Absenteeism.time.in.hours")]))
df_month=df_month[2:38, ]
df_month=df_month[,2:3 ]
df_month$Month=1
df_month=df_month[,2:3 ]
for (i in 1:37){
        df_month$Month[i]=i
}

#Rearranging the positions of the dataset
df_month = df_month[,c(2,1)]
lrmodel_2011 = rpart(Absenteeism.time.in.hours~., data=df_month )

#Creating the test dataset
new=data.frame(Month=c(38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54),
                      predict_year=c(2010,2010,2010,2010,2010,2011,2011,2011,2011,2011,2011
                                     ,2011,2011,2011,2011,2011,2011))

#Predicting the absentee hours from August 2010 to December 2011
pred_lr2011 = predict(lrmodel_2011, newdata = new, se.fit = TRUE)
pred_lr2011
