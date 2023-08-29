Attribute VB_Name = "Módulo1"

Sub slicers(slName As String)
Dim resp
    Dim slPlataformas As slicerItem, slFic As slicerItem
    Dim slBox As SlicerCache
    Set slBox = ActiveWorkbook.SlicerCaches(slName)
    For Each slPlataformas In slBox.SlicerItems
        slBox.ClearManualFilter
        For Each slFic In slBox.SlicerItems
            slFic.Selected = (slFic.Name = slPlataformas.Name)
        Next slFic
        MsgBox slPlataformas.Name
    Next slPlataformas
End Sub

Sub test()
    Call slicers("SegmentaçãodeDados_PRODUTO")
End Sub


Sub UpdatePowerPointWithIconAndLink()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim excelWorkbook As Workbook
    Dim excelWorksheet As Worksheet
    Dim excelChart As ChartObject
    Dim excelTable As ListObject
    Dim slicerItem As slicerItem
    Dim chartLeft As Integer
    Dim chartTop As Integer
    Dim chartWidth As Integer
    Dim chartHeight As Integer
    Dim tableLeft As Integer
    Dim tableTop As Integer
    Dim iconLeft As Integer
    Dim iconTop As Integer
    Dim iconWidth As Byte
    Dim iconHeight As Byte
    Dim targetSlideID As Integer
    
    ' Define posições
    chartLeft = 100
    chartTop = 100
    chartWidth = 300
    chartHeight = 300
    tableLeft = 100
    tableTop = 100
    iconLeft = 200
    iconTop = 400
    iconWidth = 30
    iconHeight = 30
    
    ' Abre ppt
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations.Open("C:\Users\antho\OneDrive\Documentos\TesteVBA\vbaPlataformas.pptx")
    
    ' Filtraremos o slicer ** possivelmente remover
    Set slicerItem = ThisWorkbook.SlicerCaches("SegmentaçãodeDados_PRODUTO").SlicerItems("A1")
    slicerItem.Selected = True
    
    ' Copia gráfico do excel
    Set excelWorkbook = ThisWorkbook
    Set excelWorksheet = excelWorkbook.Sheets("Planilha2")
    Set excelChart = excelWorksheet.ChartObjects("grafico1")
    excelChart.Copy
    
    ' Cola gráfico no ppt
    Set pptSlide = pptPresentation.Slides(1)
    pptSlide.Shapes.PasteSpecial(DataType:=2).Select
    Set pptShape = pptSlide.Shapes(pptSlide.Shapes.Count)
    pptShape.Left = chartLeft
    pptShape.Top = chartTop
    pptShape.Width = chartWidth
    pptShape.Height = chartHeight
    
'     Copy table from Excel
'    Set excelTable = excelWorksheet.ListObjects("TableName") ' Change to your table name
'    excelTable.Range.Copy
'
'     Add a new slide and paste table to PowerPoint
'    pptSlide.Shapes.PasteSpecial(DataType:=2).Select
'    Set pptShape = pptSlide.Shapes(pptSlide.Shapes.Count)
'    pptShape.Left = tableLeft
'    pptShape.Top = tableTop
    
    pptPresentation.Slides(2).Name = "regiao2" ' Altera o nome do slide
        targetSlideID = CStr(pptPresentation.Slides("regiao2").SlideID)
        targetSlideIndex = CStr(pptPresentation.Slides("regiao2").SlideIndex)
    Set pptSlide = pptPresentation.Slides(1)
    ' Adiciona o ícone ao slide 1 para o slide da plataforma
    Set pptShape = pptSlide.Shapes.AddPicture("https://cdn.hubblecontent.osi.office.net/icons/publish/icons_magnifyingglass/magnifyingglass.svg", msoFalse, msoTrue, iconLeft, iconTop, iconWidth, iconHeight)
    With pptShape.ActionSettings(1) ' 1 corresponde ao ppMouseClick
        .Action = 5 ' 5 corresponde ao Hyperlink
        .Hyperlink.Address = ""
        .Hyperlink.SubAddress = targetSlideID & "," & targetSlideIndex & ",Title "
    End With

    Set pptShape = Nothing
    Set pptSlide = Nothing
    Set pptPresentation = Nothing
    Set pptApp = Nothing
    Set excelTable = Nothing
    Set excelChart = Nothing
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
End Sub

