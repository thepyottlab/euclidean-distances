# Pyott Lab Protocols @ thepyottlab.com
# R script to determine Euclidean distances between neighboring immunopuncta
# Last updated 07 February 2021


# clear workspace (NOTE: this command removes all your unsaved work in the current R session)
rm(list = ls())

#### This script uses the package "readxl" to import .xlsx files. The following 3 lines install and load this package: ####
packToInstall = 'readxl'
if (!require(packToInstall, character.only = TRUE)){install.packages(packToInstall)}
library(readxl)

#### Set the working directory ####
  # R will load the excel files you want to analyze from this directory, and also export any output there. It is recommended you make a new folder somewhere on your desktop specifically for this analysis.  
    path2analysisFolder = "C:/Users/luscu/Desktop/Reijntjes et al., 2020 STAR Protocols/data" # Set this to a path on your own computer
    setwd(path2analysisFolder)
  
  # To check whether your path is correct use: getwd()
    getwd()
  
#### specify immunolabels ####
  # NB: make sure the titles of the excel files only contain the identifier once per immunolabel.
    fileNameImmunolabel1 = 'CTBP2'
    fileNameImmunolabel2 = 'Synapsin'
  
#### load all .xlsx files in folder (make sure the only .xlsx files in the folder are those containing immunopuncta coordinates) ####
    #files <- dir(pattern= paste(fileNameImmunolabel1,',*', '.xlsx', sep = '')) #e.g. dir(pattern= '*immunolabel1.xlsx')
    files <- Sys.glob(paste0(fileNameImmunolabel1,'*.xlsx'))
  # To check if the right file(s) is/are selected, call object:
    files

# When using Imaris, the information for this analysis is in the 'Position' tab in the exported excel files  
# When using a program that generates a different excel file containing x,y,z, coordinates, change 'Position' in (sheet = 'Position') to specify the name of the required tab in excel.
# When exporting x,y,z, coordinates from Imaris, the first row contains no x,y,z, coordinate data. Therefore the first row is skipped and the Header is imported ('col_names = TRUE, skip = 1,')
# When using a different program than Imaris that changes the excel output, change these two settings accordingly.


#### Calculate the Euclidean distance (e.g. the distance between a given immunopunctum1 and it's closest immunopunctum2) and write the results to a new excel file 'results' for each pair of files ####
  for (file in files){
    immunolabel1 <- readxl::read_xlsx (file, sheet = 'Position', col_names = TRUE, skip = 1)
    names(immunolabel1) <- c('x', 'y', 'z', 'unit', 'category', 'collection', 'time', 'id')
  
    # to remove NA if any use:
    # immunolabel1 < - immunolabel1[rowSums(is.na(immunolabel1)) == 0, ] #activate by uncommenting this line
      if (! file.exists(gsub(fileNameImmunolabel1, fileNameImmunolabel2, file)))
      {
    # # If R produces "cannot find immunolabel 2 file' then probably there is something wrong with the excel file names and the error occurs here # #
      cat('cannot find immunolabel2 file\n')				   
    # # R needs to be able to find an excel file with exactly the same title save for a change in the immunolabel. ##	
      } else {
      immunolabel2 <- readxl::read_xlsx (gsub(fileNameImmunolabel1, fileNameImmunolabel2, file),
                                         sheet = 'Position', col_names = TRUE, skip = 1)
      names(immunolabel2) <- c('x', 'y', 'z', 'unit', 'category', 'collection', 'time', 'id')
    
    #calculating eucledian distances:
      minDist=pointId <- rep(0,  dim(immunolabel1)[1])
      for (iPoint in 1 : nrow(immunolabel1)){
          EuclidDistance = sqrt(
          (immunolabel1$x[iPoint] - immunolabel2$x)^2 + 
          (immunolabel1$y[iPoint] - immunolabel2$y)^2 + 
          (immunolabel1$z[iPoint] - immunolabel2$z)^2)
        minDist[iPoint] = min(EuclidDistance)
        pointId[iPoint] = immunolabel2$id[EuclidDistance == min(EuclidDistance)]
        }
    
    #writing results to the object "results" and setting column names including subject ID
      results <- c(immunolabel1$id, minDist, pointId)
      dim(results) <- c(dim(immunolabel1)[1], 3)
      last_ID_character = gregexpr(pattern ='_', file)[[1]][1]-1 #Finds first "_" character in file name to separate subject ID
      colnames(results) <- c(paste(fileNameImmunolabel1,"ID"), "Distance", paste(fileNameImmunolabel2,"ID"))
    }
  
    #Save file in .csv format. Each result is saved as a separate file.
      #results_filename=gsub(fileNameImmunolabel1, "results", gsub(".xlsx", ".csv", file))
      subjectID=gsub(paste0(fileNameImmunolabel1,"_"), "", gsub("*.xlsx", "", file))
      results_filename=paste0(subjectID,"_results.csv")
      write.csv(results, file=results_filename, row.names = FALSE, quote = FALSE)
    
}

# if you get the return "cannot find immunolabel2 file", the most likely problem is that the two files for your immunolabel1 and immunolabel2 do not correspond. The file names need to be exactly the save for the _immunolabel1 and _immunolabel2 specification at the end of the file name. 
