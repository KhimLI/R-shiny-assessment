library(shiny)
library(readxl)

ui <- fluidPage(
  
  titlePanel("Claim Data Upload"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput(
        inputId = "file",
        label = "Upload your file",
        accept = ".xlsx"
        )
      ),
    mainPanel(
      tableOutput("content")
      )
    )
)
server <- function(input,output,session){
  output$content <- renderTable({
    req(input$file)
    read_excel(input$file$datapath)
  })
}

shinyApp(ui,server)