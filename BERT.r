library(MASS)

### BREADTH

Breadth <- function(rtns)
{
  corl <- cor(rtns);                    corl <- 1-abs(corl) + diag(ncol(corl))
  
  corl[upper.tri(corl)] <- 1;           br <- matrix(0, nrow = nrow(corl), ncol = 1)
  
  for(i in 1:nrow(corl))
  {
    br[i,1] <- prod(corl[i,])
  }
  return(sum(br))
}

### remove colinearity above x
rmlinear <- function(mtx,x,wts,result)
{
  mtx <- mtx - diag(ncol(mtx));         rnk <- rank(abs(wts),ties.method = "first")

  for (i in rnk[rnk])
  {
    for (j in rnk[rnk])
    {
      if(mtx[j,i]>x)
      {
        #mtx[,i] <- 0;                   mtx[i,] <- 0
        
        #wts[j] <- wts[j] + wts[i];      wts[i] <- 0              
      }
    }
  }


  if(result == "co")
  {
    return(mtx)
  }else{
    return(wts)
  } 
}

### Optimal Information Ratio

ir <- function(rtns,alphas,wts)
{
  co <- cor(rtns);                      co <- rmlinear(co,0.1,wts,"co")
  
  rtns <- rtns[,c(!colSums(co)==0)];    alphas <- alphas[c(!colSums(co)==0)]
  
  return(sqrt(t(alphas) %*% ginv(52*cov(rtns),tol = 0) %*% alphas))
}

ir2 <- function(rtns,alphas,opts,wts)
{
  co <- cor(rtns);                      co <- rmlinear(co,0.1,wts,"co")
  
  rtns <- rtns[,c(!colSums(co)==0)];    alphas <- alphas[c(!colSums(co)==0)]

  opts <- opts[c(!colSums(co)==0)];      alphas <- alphas * opts


   return(sqrt(t(alphas) %*% ginv(52*cov(rtns),tol = 0) %*% alphas))

}

expostIR <- function(rtns,alphas,wts)
{
  return(alphas%*%t(wts)/sd(rtns %*% t(wts)))
}

### Optimization

optimization <- function(rtns,alphas,wts)
{
  co <- cor(rtns);                      co <-  rmlinear(co,0.1,wts,"co")

  pholdr1 <- 1:ncol(rtns);              pholdr2 <- pholdr1[c(!colSums(co)==0)]

  cova <- 52*cov(rtns);                    cova <- cova[c(!colSums(co)==0),
                                        c(!colSums(co)==0)];                  
  
  alphas <- alphas[c(!colSums(co)==0)]

  inv <- ginv(cova,tol = 0);        opts2 <- inv %*% alphas

  opts1 <- numeric(length(pholdr1));    opts1[pholdr2] <- opts2

  tot <- sum(opts1);                    opts1 <- opts1 - tot/length(opts1)

  opts1 <- opts1/max(abs(opts1)) * 0.05

  return(t(opts1))
}


PC1 <- function(rtns)
{
  pca <- prcomp(rtns,center = TRUE,scale. =TRUE)
  return(pca$x[,1:3])
}

### The Transferr Coeffecient

TC <- function(rtns,wts,opts)
{
  co <- cor(rtns);                      wts <- rmlinear(co,0.1,wts,"wts")
  
  return(cor(t(wts),t(opts)))
}

coMatrix <- function(rtn,type)
{
  if(type == "covariance"){return(cov(rtn))} else {return(cor(rtn))}
}
ExRtns <- function(wts,exp)
{
  exp[exp>0 & wts>0] <- wts[exp>0 & wts>0]

  exp[exp<0 & wts<0] <- wts[exp<0 & wts<0] 

  exp[exp>=0 & wts<=0] <- 0 

  exp[exp<=0 & wts>=0] <- 0.05
  
  return(exp)
}
#rtn <- range.to.data.frame(EXCEL$Application$get_Sheets("residual returns")$get_Range("b177:jb438")$get_Value(),FALSE)
#ave <- range.to.data.frame(EXCEL$Application$get_Sheets("averages")$get_Range("b438:jb438")$get_Value(),FALSE)
#EXCEL$Application$get_Sheets("Risk")$get_Range(paste0("b3:c",3+dim(All_List)[1]))$put_Value(All_List)

weighter <- function(date,nms,wgt,wgts,returns)
{
  if(sum(wgts[,1]==date)>0)
  {
    ans <- t(wgts[wgts[,1]==date,][-1])
    ans[is.na(ans)] <- 0
    ans <- ans/sum(ans)
    return(ans);
  }
  else
  {
    index <- match(nms,substr(returns[1,],1,3))
    returns <- returns[,c(1,index)]
    
    index <- match(c("Row Labels",date-1,date),returns[,1])
    returns <- returns[index,]
    returns[sapply(returns, is.null)] <- NA
    returns[is.na(returns)] <- 1
    returns <- returns[-1,-1]
    
    open <- as.numeric(unlist(returns[1,]))
    close <- as.numeric(unlist(returns[2,]))
    
    close[open==1] <- 1
    open[close==1] <- 1

    ans <- wgt*close/open
    ans <- ans/sum(ans)
   
    return(ans)

  }
}

InDependantBeta <- function(y,ynms,x,xnms,nd)
{
  x <- as.data.frame(x)
  y <- as.data.frame(y)
  nd <- as.data.frame(nd)
  colnames(x) <- xnms
  colnames(y) <- ynms
  colnames(nd) <- xnms
  
  model <- lm(as.matrix(y)~.,data=x)
  
  pred <- predict(model,newdata = nd)

  return(t(as.numeric(pred)))
}

cleandata <- function(data1,r_nms,c_nms)
{
  data1 <- data.frame(data1)
  dims <- dim(data1)

  
  rownames(data1) <- r_nms
  colnames(data1) <- c_nms

  data1[data1< -0.5] <- NA
  data1[data1> 0.5] <- NA

  
  data1 <- sapply(data1,function(x) as.numeric(x))
  data1 <- as.data.frame(data1)
  data1 <- data1[,sapply(data1,function(x) !sum(!is.na(x)))==0]

  for (i in colnames(data1))
  {
    model <- lm(paste0("`",i,"` ~ .JDALST"),data=data1)
    
    
    nd <- sapply(data1,function(x) {x[is.na(x)] <-mean(x,na.rm=TRUE);return(x)})
    nd <- as.data.frame(nd)
    nd <- data1[is.na(data1[i]),]
    
    #data1[i][is.na(data1[i])] <- predict(model,newdata = nd[is.na(data1[i]),])
  }

  data1[is.na(data1)] <- 0
  data1[data1< -0.5] <- 0
  data1[data1> 0.5] <- 0

  return(data1)
}
table2vector <- function(a)
{
  a[is.na(a)] <- 0
  a <- sapply(a,function(x) rbind(x))
}

ColRep <- function(a)
{
  b <- a[-c(1:nrow(a)),]
  a[is.na(a)] <- 0
  for(i in 1:288)
  {
    b <- rbind(b,a)
  }
  return(b)
}

plotss <- function(x,rs,cs)
{
  x <- as.data.frame(x)
  colnames(x) <- cs
  rownames(x) <- rs
  plot(x[1])
}
getforecast <- function(df,d,t)
{
  
  df[sapply(df,is.null)] <- ""
  df <- data.frame(matrix(unlist(df),ncol = ncol(df)),stringsAsFactors = FALSE)
  colnames(df) <- df[1,]
  df <- df[-1,]
  df <- df[df$Ticker == t,match(c('Ticker','Date','Exp'), colnames(df))]

  df <- df[df$Date<d,]
  
  return(df$Exp[nrow(df)])
}
rankcorr <- function(x,y)
{
  x <- unlist(x)
  y <- unlist(y)

  x <- as.numeric(x)
  y <- as.numeric(y)

  xx <- x[!is.na(x) & !is.na(y)]
  yy <- y[!is.na(x) & !is.na(y)]
  #return(xx)
  return(cor.test(as.numeric(xx[]),as.numeric(yy[]),method="spearman")$estimate[[1]])
}
lg <- function(x)
{
  return(log(x))
}
countifr <- function(x,y)
{
  x <- unlist(x)
  y <- unlist(y)

  x <- as.numeric(x)
  y <- as.numeric(y)

  xx <- x[!is.na(x) & !is.na(y)]
  yy <- y[!is.na(x) & !is.na(y)]
  zz <- xx*yy
  #return(xx)
  if(length(zz>0)){return(sum(zz>0)/length(zz))}else{""}
}
countifrsec <- function(x,y,sector,sec)
{
  x <- unlist(x)
  y <- unlist(y)

  x <- as.numeric(x)
  y <- as.numeric(y)

  xx <- x[!is.na(x) & !is.na(y)]
  yy <- y[!is.na(x) & !is.na(y)]

  zz <- xx*yy

  sectorss <- sector[!is.na(x) & !is.na(y)]

  zz <- zz[sectorss==sec]

  #return(xx)
  if(length(zz>0)){return(sum(zz>0)/length(zz))}else{""}
}

reverse <- function(x)
{
  return (rev(x))
}

diagonal <- function(x)
{
  return (diag(t(x)))
}

wrang_regr <- function(ret,fac)
{
  ret <- t(ret)
  fac <- matrix(fac,nrow = length(ret))
  model <- ret ~fac
  return(summary(model))
}