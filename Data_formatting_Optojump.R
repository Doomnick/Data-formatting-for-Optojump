# Import
library(readxl)
require(dplyr)
require(DataCombine)
require(reshape)
library(tools)

path <- "D:/***/***/nabehovky/tyč muži/OJ/raw.xlsx"
zdata <- read_excel(path)

# Crop data
empty_columns <- sapply(zdata, function(x) all(is.na(x) | x == "")) # Vector with empty columns
bezpraz <- zdata[, !empty_columns] # Remove empty columns
bezpraz['Step'] <- c(NA, head(bezpraz['Step'], dim(bezpraz)[1] - 1)[[1]])
bezpraz <- bezpraz[!grepl("^External trigger$", fixed = F, bezpraz$'#'),] # Shift 'Step' column down by one row
bezpraz <- dplyr::rename(bezpraz, "trigger" = "#") # Rename column "#" to "trigger"
bezpraz$trigger[bezpraz$trigger == "External trigger STOP"] <- "ET"
bezpraz <- bezpraz[!grepl("^External trigger$", fixed = F, bezpraz$trigger),] # Remove rows containing "external trigger"
bezpraz$trigger[bezpraz$trigger == "External trigger STOP"] <- NA
cols.want <- c("Last & first names", "Time", "trigger", "Flight time", "Contact time", "Step", "Speed", "Notes", "Pace[step/s]", "Step time", "L/R") # Desired columns
osek <- bezpraz[, names(bezpraz) %in% cols.want, drop = F] # Remove columns that are not desired

osek <- dplyr::rename(osek, "jmeno" = "Last & first names") # Rename the first column
osek <- dplyr::rename(osek, "Flight time (s)" = 'Flight time')
osek <- dplyr::rename(osek, "Contact time (s)" = 'Contact time')
osek <- dplyr::rename(osek, "Step length (cm)" = 'Step') 
osek <- dplyr::rename(osek, "Speed (m/s)" = 'Speed')  
osek <- dplyr::rename(osek, "Frequency" = 'Pace[step/s]')
if ("Step time" %in% colnames(osek)) {
  osek <- dplyr::rename(osek, "Step duration (s)" = 'Step time') 
}

# Copy the attempt time to a new column
osek$Doba <- osek$Time
osek <- relocate(osek, Doba, .after = Time)
osek <- relocate(osek, "Flight time (s)", .after = "Contact time (s)")
osek <- relocate(osek, "Step length (cm)", .after = "Flight time (s)")

############################# Convert from long to wide format ####################################

unik.cas <- unique(osek[c("Time")]) # Unique time values
unik.cas$poradi.cas <- 1:nrow(unik.cas) # Add values for find-replace

osek <- as.data.frame(osek) # Convert to data.frame for FindReplace function
trigger1 <- replace(osek$trigger, is.na(osek$trigger), "solo")
osek$trigger <- trigger1

trigger1 <- replace(osek$trigger, osek$trigger == "ET", NA) # Replace "no step" with numbers 1:150 <- unique data
osek$trigger <- trigger1
trigger3 <- as.data.frame(osek$trigger)
vals1 <- seq(50, 99, 1)
trigger3[is.na(trigger3)] <- sample(vals1, sum(is.na(trigger3)), replace = TRUE)
trigger3.vector <- trigger3$`osek$trigger`
osek$trigger <- trigger3.vector

trigger <- replace(osek$trigger, osek$trigger == "No step", NA) # Replace "no step" with numbers 1:150 <- unique data
osek$trigger <- trigger
trigger2 <- as.data.frame(osek$trigger)
vals <- seq(100, 150, 1)
trigger2[is.na(trigger2)] <- sample(vals, sum(is.na(trigger2)), replace = TRUE)
trigger2.vector <- trigger2$`osek$trigger`
osek$trigger <- trigger2.vector

# Replace time with code 1-x
osek2 <- FindReplace(data = osek, Var = "Time", replaceData = unik.cas, from = "Time", to = "poradi.cas", exact = T, vector = F)
osek2$Time <- as.factor(osek2$Time)

# Convert long to wide format based on "Time" and "trigger"
sum(duplicated(osek2[, c("Time", "trigger")]))
column_names <- colnames(osek2)
remove_columns <- c("Time", "trigger")
column_names <- column_names[!column_names %in% remove_columns]
datawide <- reshape(osek2, v.names = column_names, timevar = "trigger", idvar = "Time", direction = "wide", sep = " ")
datawide <- rename(datawide, replace = c("jmeno 1" = "Name"))

# Remove columns with names containing "jmeno"
datawide <- datawide[, -grep("jmeno", colnames(datawide))]
datawide <- rename(datawide, replace = c("Time" = "ID"))

datawide2 <- datawide %>% # Numbering attempts
  dplyr::group_by(datawide$Name) %>%
  dplyr::mutate(Attempt = dplyr::row_number()) %>%
  dplyr::ungroup()

# Remove renamed column
datawide2 <- relocate(datawide2, Attempt, .after = Name) # Move Attempt column after Name
datawide2 <- subset(datawide2, select = -c(ID))
colnames(datawide2)[which(names(datawide2) == "Notes 1")] <- "Note" # Rename Notes 1 to Note
if ("Note" %in% colnames(datawide2)) {
  datawide2 <- datawide2[, -grep("Notes", colnames(datawide2))] # Remove other Notes
  datawide2 <- relocate(datawide2, Note, .after = Attempt)
}
empty_columns2 <- sapply(datawide2, function(x) all(is.na(x) | x == "")) # Vector with empty columns
datawide2 <- datawide2[, !empty_columns2] # Remove empty columns

datawide2 <- relocate(datawide2, "Doba 1", .before = Name)
datawide2 <- rename(datawide2, replace = c("Doba 1" = 'Time'))  
datawide2 <- datawide2[, -grep("Doba", colnames(datawide2))]

datawide2 <- dplyr::rename(datawide2, "x" = "datawide$Name")  # Rename column datawide$Name
datawide2 <- subset(datawide2, select = -c(x))

trigger2.vector <- as.numeric(trigger2.vector)
num_no_step <- length(trigger2.vector[trigger2.vector > 100])
num_no_step <- as.numeric(num_no_step)

dk <- colnames(dplyr::select(.data = datawide2, contains("Step length")))
dk <- dk[-1]
fk <- colnames(dplyr::select(.data = datawide2, contains("Frequency")))
of <- colnames(dplyr::select(.data = datawide2, contains("Contact")))
lf <- colnames(dplyr::select(.data = datawide2, contains("Flight")))
ts <- colnames(dplyr::select(.data = datawide2, contains("Duration")))
datawide2$'Average step length' <- rowMeans(datawide2[dk], na.rm = T)
datawide2$'Average frequency' <- rowMeans(datawide2[fk], na.rm = T)
datawide2$'Average contact' <- rowMeans(datawide2[of], na.rm = T)
datawide2$'Average flight' <- rowMeans(datawide2[lf], na.rm = T)
if ("Step time" %in% colnames(osek)) {
  datawide2$'Average duration' <- rowMeans(datawide2[ts], na.rm = T)
}
datawide2$'Contact/Flight' <- datawide2$'Average contact' / datawide2$'Average flight'
datawide2[,'Average step length'] <- round(datawide2[,'Average step length'], 1)
datawide2[,'Average contact'] <- round(datawide2[,'Average contact'], 3)
datawide2[,'Average flight'] <- round(datawide2[,'Average flight'], 3)
datawide2[,'Average frequency'] <- round(datawide2[,'Average frequency'], 2)
datawide2[,'Contact/Flight'] <- round(datawide2[,'Contact/Flight'], 2)
if ("Average duration" %in% colnames(datawide2)) {
  datawide2[,'Average duration'] <- round(datawide2[,'Average duration'], 3)
}

if (num_no_step > 0) { datawide2[nrow(datawide2) + 1,] <- NA }
if (num_no_step > 0) { datawide2 <- rbind(datawide2[nrow(datawide2),], datawide2[1:nrow(datawide2) - 1,]) }
if (num_no_step > 0) { datawide2[1, 1] <- "Step number greater than 100, check (No step)" }

# Export
path <- file_path_sans_ext(path)
writexl::write_xlsx(datawide2, paste(path, "_edited.xlsx", sep = ""), col_names = T, format_headers = T)