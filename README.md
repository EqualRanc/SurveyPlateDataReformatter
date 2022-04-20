# SurveyPlateDataReformatter
Takes .xml files outputted from the Labcyte Echo liquid handler's survey plate function, and summarizes the survey data into an Excel spreadsheet.

The Labcyte Echo liquid handler has a survey function that uses its acoustic transducer to determine approximate volumes (in microliters) of wells in a microplate. This code is intended to process the result files (.xml files) and summarize them in an Excel spreadsheet. It also applies a gradient to create plate heatmaps for easy visualization. Red signifies low volumes and green signifies high volumes. It is currently written for 384W and 1536W plates only.
