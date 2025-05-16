# Clear variables
rm(list=ls())

# Load packes
library(tidyverse)
library(meta)
library(metasens)
library(metafor)
library(dmetar)
library(robvis)
library(ggplot2)

# Set working directory
setwd("C:/Users/irisv/Google Drive/PhD/Systematic review/R/Risk of bias")

# RCTs
Data_RoB <- read.csv("RoB_R_Data.csv", header = TRUE, sep = ";")
Bar_plot_RoB = rob_summary(data = Data_RoB, tool = 'ROB2', weighted = TRUE, overall = TRUE, colour = "colourblind") 
ggsave(plot = Bar_plot_RoB, filename = "RoB_barplot.png", width = 8, height = 2.41, dpi = 1000)
Traffic_light_plot_RoB <- rob_traffic_light(data = Data_RoB, tool = "ROB2",colour = "colourblind")
ggsave(plot = Traffic_light_plot_RoB, filename = "RoB_traffic_light_plot.png", width = 8, height = 8, dpi = 1000)

# Open-label trials and retrospective cohort studies
Data_ROBINS <- read.csv("RoBINS_I_R_Data.csv", header = TRUE, sep = ";")
Bar_plot_ROBINS = rob_summary(data = Data_ROBINS, tool = 'ROBINS-I', weighted = TRUE, overall = TRUE, colour = "colourblind") 
ggsave(plot = Bar_plot_ROBINS, filename = "RoBINS-I_barplot.png", width = 8, height = 2.41, dpi = 1000)
Traffic_light_plot_ROBINS <- rob_traffic_light(data = Data_ROBINS, tool = "ROBINS-I",colour = "colourblind")
ggsave(plot = Traffic_light_plot_ROBINS, filename = "ROBINS-I_traffic_light_plot.png", width = 6, height = 13, dpi = 1000)