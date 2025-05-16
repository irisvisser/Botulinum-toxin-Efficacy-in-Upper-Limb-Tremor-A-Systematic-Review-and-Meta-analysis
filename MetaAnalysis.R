# Load packages
library(tidyverse)
library(meta)
library(metasens)
library(metafor)
library(dmetar)
library(openxlsx)
library(gridExtra)

# Clear variables
rm(list=ls()) 

# Set working directory
setwd("C:/Users/irisv/Google Drive/PhD/Systematic review/R/MetaAnalysis") 

# TREMOR SEVERITY --------------------------------------------------------------
# Load data
data_TS <- read.xlsx("MetaAnalysisTremorSeverity_PostIntervention.xlsx") 

# Convert data classes
data_TS$n.e <- as.numeric(data_TS$n.e)
data_TS$mean.e <- as.numeric(data_TS$mean.e)
data_TS$sd.e <- as.numeric(data_TS$sd.e)
data_TS$n.c <- as.numeric(data_TS$n.c)
data_TS$mean.c <- as.numeric(data_TS$mean.c)
data_TS$sd.c <- as.numeric(data_TS$sd.c)
data_TS$diagnosis <- as.factor(data_TS$diagnosis)
data_TS$outcome <- as.factor(data_TS$outcome)

m.cont_TS <- metacont(n.e = n.e, mean.e = mean.e, sd.e = sd.e, n.c = n.c, mean.c = mean.c, sd.c = sd.c,
                      studlab = author, 
                      data = data_TS,
                      sm = "SMD", # Standardized mean differences
                      method.smd = "Hedges", # Hedges correction
                      fixed = FALSE, # Do not use fixed-effects model
                      random = TRUE, # Use random-effects model
                      method.tau = "REML", # Restricted Maximum-Likelihood Estimator
                      method.random.ci = "HK") # Knapp-Hartung Adjustment
summary(m.cont_TS)

# GRIP STRENGTH ----------------------------------------------------------------
# Load data
data_grip <- read.xlsx("MetaAnalysisGripStrength_PostIntervention.xlsx")

# Convert data classes
data_grip$n.e <- as.numeric(data_grip$n.e)
data_grip$mean.e <- as.numeric(data_grip$mean.e)
data_grip$sd.e <- as.numeric(data_grip$sd.e)
data_grip$n.c <- as.numeric(data_grip$n.c)
data_grip$mean.c <- as.numeric(data_grip$mean.c)
data_grip$sd.c <- as.numeric(data_grip$sd.c)
data_grip$diagnosis <- as.factor(data_grip$diagnosis)

m.cont_grip <- metacont(n.e = n.e, mean.e = mean.e,sd.e = sd.e, n.c = n.c, mean.c = mean.c, sd.c = sd.c,
                        studlab = author,
                        data = data_grip,
                        sm = "SMD", # Standardized mean differences
                        method.smd = "Hedges", # Hedges correction
                        fixed = FALSE, # Do not use fixed-effects model
                        random = TRUE, # Use random-effects model
                        method.tau = "REML", # Restricted Maximum-Likelihood Estimator
                        method.random.ci = "HK") # Knapp-Hartung Adjustment
summary(m.cont_grip)

# PATIENT GLOBAL IMPRESSION OF CHANGE ---------------------------------------------------------------------------
# Load data
data_PGIC <- read.xlsx("MetaAnalysisPGIC.xlsx")

# Convert data classes
data_PGIC$n.e <- as.numeric(data_PGIC$n.e)
data_PGIC$mean.e <- as.numeric(data_PGIC$mean.e)
data_PGIC$sd.e <- as.numeric(data_PGIC$sd.e)
data_PGIC$n.c <- as.numeric(data_PGIC$n.c)
data_PGIC$mean.c <- as.numeric(data_PGIC$mean.c)
data_PGIC$sd.c <- as.numeric(data_PGIC$sd.c)
data_PGIC$diagnosis <- as.factor(data_PGIC$diagnosis)
data_PGIC$outcome <- as.factor(data_PGIC$outcome)

m.cont_PGIC <- metacont(n.e = n.e, mean.e = mean.e, sd.e = sd.e, n.c = n.c, mean.c = mean.c, sd.c = sd.c,
                        studlab = author,
                        data = data_PGIC,
                        sm = "SMD", # Standardized mean differences
                        method.smd = "Hedges", # Hedges correction
                        fixed = FALSE, # Do not use fixed-effects model
                        random = TRUE, # Use random-effects model
                        method.tau = "REML", # Restricted Maximum-Likelihood Estimator
                        method.random.ci = "HK") # Knapp-Hartung Adjustment
summary(m.cont_PGIC)

# Plotting
svg(file = "forestplot_TS_post.svg",width = 12, height = 4)
FP_TS <- meta::forest(m.cont_TS, 
             studlab = TRUE,
             label.right = "Favours placebo", label.left = "Favours BoNT",
             leftcols = c("studlab", "diagnosis","outcome","n.e","mean.e","sd.e","n.c","mean.c","sd.c"),
             leftlabs = c("Study", "Diagnosis","Outcome measure","N","Mean","SD","N","Mean","SD"),  
             digits.sd = 2)  
dev.off() # Save output

svg(file = "forestplot_GripStrength_post.svg", width = 12, height = 4)
FP_GS <- meta::forest(m.cont_grip, 
             studlab = TRUE,
             label.right = "Grip strength increase",
             label.left = "Grip strength loss",
             leftcols = c("studlab", "diagnosis","n.e","mean.e","sd.e","n.c","mean.c","sd.c"),
             leftlabs = c("Study", "Diagnosis","N","Mean","SD","N","Mean","SD"),  
             digits.sd = 2)  
dev.off() # Save output

svg(file = "forestplot_PGIC.svg", width = 20, height = 4)
FP_PGIC <- meta::forest(m.cont_PGIC, 
             studlab = TRUE,
             label.right = "Favours BoNT",
             label.left = "Favours placebo",
             leftcols = c("studlab", "diagnosis","outcome","n.e","mean.e","sd.e","n.c","mean.c","sd.c"),
             leftlabs = c("Study", "Diagnosis","Outcome measure","N","Mean","SD","N","Mean","SD"),  
             digits.sd = 2)  
dev.off() # Save output

# SECONDARY ANALySIS
# Clear variables
rm(list=ls()) 

# Set working directory
setwd("C:/Users/irisv/Google Drive/PhD/Systematic review/R/MetaAnalysis") 

# TREMOR SEVERITY --------------------------------------------------------------
# Load data
data_TS <- read.csv("MetaAnalysisTremorSeverity_change.csv") 

# Convert data classes
data_TS$n.e <- as.numeric(data_TS$n.e)
data_TS$mean.e <- as.numeric(data_TS$mean.e)
data_TS$sd.e <- as.numeric(data_TS$sd.e)
data_TS$n.c <- as.numeric(data_TS$n.c)
data_TS$mean.c <- as.numeric(data_TS$mean.c)
data_TS$sd.c <- as.numeric(data_TS$sd.c)
data_TS$diagnosis <- as.factor(data_TS$diagnosis)
data_TS$outcome <- as.factor(data_TS$outcome)

m.cont_TS_change <- metacont(n.e = n.e, mean.e = mean.e, sd.e = sd.e, n.c = n.c, mean.c = mean.c, sd.c = sd.c,
                      studlab = author, 
                      data = data_TS,
                      sm = "SMD", # Standardized mean differences
                      method.smd = "Hedges", # Hedges correction
                      fixed = FALSE, # Do not use fixed-effects model
                      random = TRUE, # Use random-effects model
                      method.tau = "REML", # Restricted Maximum-Likelihood Estimator
                      method.random.ci = "HK") # Knapp-Hartung Adjustment
summary(m.cont_TS_change)

png(file = "forestplot_TS_change.png", width = 4000, height = 1000, res = 300)
meta::forest(m.cont_TS_change, 
             studlab = TRUE,
             label.right = "Favours placebo", label.left = "Favours BoNT",
             leftcols = c("studlab", "diagnosis","outcome","n.e","mean.e","sd.e","n.c","mean.c","sd.c"),
             leftlabs = c("Study", "Diagnosis","Outcome measure","N","Mean","SD","N","Mean","SD"),  
             digits.sd = 2)  
dev.off() # Save output

# GRIP STRENGTH ----------------------------------------------------------------
# Load data
data_grip <- read.csv("MetaAnalysisGripStrength_change.csv")

# Convert data classes
data_grip$n.e <- as.numeric(data_grip$n.e)
data_grip$mean.e <- as.numeric(data_grip$mean.e)
data_grip$sd.e <- as.numeric(data_grip$sd.e)
data_grip$n.c <- as.numeric(data_grip$n.c)
data_grip$mean.c <- as.numeric(data_grip$mean.c)
data_grip$sd.c <- as.numeric(data_grip$sd.c)
data_grip$diagnosis <- as.factor(data_grip$diagnosis)

m.cont_grip_change <- metacont(n.e = n.e, mean.e = mean.e,sd.e = sd.e, n.c = n.c, mean.c = mean.c, sd.c = sd.c,
                        studlab = author,
                        data = data_grip,
                        sm = "SMD", # Standardized mean differences
                        method.smd = "Hedges", # Hedges correction
                        fixed = FALSE, # Do not use fixed-effects model
                        random = TRUE, # Use random-effects model
                        method.tau = "REML", # Restricted Maximum-Likelihood Estimator
                        method.random.ci = "HK") # Knapp-Hartung Adjustment
summary(m.cont_grip_change)

png(file = "forestplot_GripStrength_change.png", width = 4000, height = 1000, res = 300)
meta::forest(m.cont_grip_change, 
             studlab = TRUE,
             label.right = "Grip strength increase",
             label.left = "Grip strength loss",
             leftcols = c("studlab", "diagnosis","n.e","mean.e","sd.e","n.c","mean.c","sd.c"),
             leftlabs = c("Study", "Diagnosis","N","Mean","SD","N","Mean","SD"),  
             digits.sd = 2)  
dev.off() # Save output

