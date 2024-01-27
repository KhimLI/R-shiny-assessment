library(shiny)
library(readxl)
library(openxlsx)
library(dplyr)
library(tidyverse)
library(ggplot2)

ui <- fluidPage(
  
  titlePanel("Claim Data Upload"),
  
  sidebarLayout(
    sidebarPanel(
      sliderInput("tail_factor",label = "Please adjust Tail Factor",
                  min = 1,max = 2,value = 1.1,step = 0.1),
      
      fileInput(inputId = "file",label = "Upload your file",accept = ".xlsx")
    ),
    mainPanel(
      tableOutput("content"),
      plotOutput("graph"),
    )
  )
)
server <- function(input,output,session){
  
  excel <- reactive({
    req(input$file)
    data.frame(read_excel(input$file$datapath))
  })
  
  output$content <- renderTable({
    req(input$file)
    data.frame(read_excel("testing.xlsx","sheet3"))
  })
  
  #Graph Creation
  
  data2 <- read_excel("testing.xlsx","sheet3",skip = 1)
  data2 <- as.data.frame(lapply(data2,as.numeric))
  
  data2_long <- data2 %>%
    gather(variable, value, -Loss_Year)
  
  output$graph <- renderPlot({
    req(input$file)
    ggplot(data2_long, aes(x = variable, y = value, color = factor(Loss_Year), label = value, group = Loss_Year)) +
      geom_point() +
      geom_line() +
      geom_text(vjust = -0.5, hjust = 0.5) +  
      labs(title = "Cumulative Paid Claims ($)",
           x = "Variable",
           y = "Cumulative Paid Claim",
           color = "Loss_Year")
  })
  
  observe({
    
    wb <- createWorkbook()
    
    addWorksheet(wb, "sheet1")
    
    
    
#Create Loss Year Column
    loss_year <- unique(data.frame(Loss_Year = excel()[, 1]))
    writeData(wb, "sheet1", loss_year, xy = c("A", "3"))
    
#Create Development Year row header
    DY <- unique(data.frame(Development_Year = excel()[,2]))
    DY_tail <- tail(DY$Development_Year, 1)+1
    writeData(wb,"sheet1",as.data.frame(t(c(DY$Development_Year,DY_tail))),
              xy = c("B", "2"),colNames = TRUE)
    saveWorkbook(wb, "testing.xlsx", overwrite = TRUE)

    testing <- reactive({
      data.frame(read_excel("testing.xlsx"))
    })
    
    #column Header Values for excel.xlsx
    e_column1 <- excel()[,1]
    e_column2 <- excel()[,2]
    
    r <- 2
    c <- 1
    while(c+1 <= ncol(testing())-1){
      while(r <= nrow(testing())){
        triangle_input <- sum(excel()[e_column1 == testing()[r,1] & e_column2 <= testing()[1,c+1],][,3])
        writeData(wb, "sheet1", triangle_input, xy = c(c+1, r+2))
        r <- r+1
      }
      c <- c+1
      r <- 2
    }
    saveWorkbook(wb, "testing.xlsx", overwrite = TRUE)
    
    #Filter the repeated value at the end
      
    
    filtered_triangle <- t(data.frame(read_excel("testing.xlsx")))
    filtered_triangle <- apply(filtered_triangle, 2, function(x) ifelse(duplicated(x), NA, x[!duplicated(x)]))
    filtered_triangle <- apply(filtered_triangle, 2, function(x) ifelse(duplicated(x), NA, x[!duplicated(x)]))
    filtered_triangle <- t(filtered_triangle)
    addWorksheet(wb,"sheet2")
    writeData(wb,"sheet2",x = filtered_triangle)
    saveWorkbook(wb, "testing.xlsx", overwrite = TRUE)
    
    addWorksheet(wb,"sheet3")
    writeData(wb,"sheet3",x = filtered_triangle)
    saveWorkbook(wb, "testing.xlsx", overwrite = TRUE)
    
    origin <- read_excel("testing.xlsx","sheet2",skip = 1)
    origin[] <- lapply(origin,as.numeric)
    
    updated_origin <- read_excel("testing.xlsx","sheet3",skip = 1)
    updated_origin[] <- lapply(updated_origin,as.numeric)
    
    r <- 1
    c <- 2
    while(c <= ncol(origin)-1){
      while(r <= nrow(origin)){
        
        column_sum <- sum(origin[,c],na.rm = TRUE)
        non_na_indices <- which(!is.na(origin[, c]))
        
        column_divide <- sum(origin[non_na_indices,c-1],na.rm = TRUE)
        
        #PS: [[1]] is use to extract the first value from the tibble
        multiply <- as.numeric((read_excel("testing.xlsx","sheet3",skip = 1)[r,c-1])[[1]])
        
        
        if(is.na(origin[r,c])){
          writeData(wb,"sheet3",x = round(column_sum*multiply/column_divide,2),xy = c(c,r+2))
          
          saveWorkbook(wb, "testing.xlsx", overwrite = TRUE)
        }
        r <- r+1
      }
      c <- c+1
      r <- 1
    }
    
    r <- 1
    c <- ncol(origin)
    while(c <= ncol(origin)){
      while(r <= nrow(origin)){
        if(is.na(origin[r,c])){
          estimate <- input$tail_factor*as.numeric((read_excel("testing.xlsx","sheet3",skip = 1)[r,c-1])[[1]])
          writeData(wb,"sheet3",x = estimate,xy = c(c,r+2))
          
          saveWorkbook(wb, "testing.xlsx", overwrite = TRUE)
        }
        r <- r+1
      }
      c <- c+1
      r <- 1
    }
  })
}



shinyApp(ui,server)
