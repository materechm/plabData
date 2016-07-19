
m1 <- 0.64882704 #mean for first set of data
m2 <- 0.672551333 #mean for second set of data
sd1 <- 0.091924365 #standard deviation for first set of data
sd2 <- 0.076447391 #standard deviation for second set of data
num1 <- 1278 #number of elements in first set of data
num2 <- 45 #number of elements in second set of data
se <- sqrt(sd1*sd1/num1+sd2*sd2/num2) #calculates standard error
t <- (m1-m2)/se #calculates t
pt(-abs(t),df=pmin(num1,num2)-1) #returns p value 
