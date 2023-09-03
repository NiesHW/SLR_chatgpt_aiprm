Sub CreatePowerPointPresentation()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim pptTextbox As Object
    
    ' Create a new PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    
    ' Create a new PowerPoint presentation
    Set pptPresentation = pptApp.Presentations.Add
    
    ' Add title slide
    Set pptSlide = pptPresentation.Slides.Add(1, ppLayoutTitle)
    pptSlide.Shapes.Title.TextFrame.TextRange.Text = "Reinforcement Learning in Bioinformatics"
    
    ' Loop through each article and create a slide
    For i = 1 To 5
        ' Add a content slide
        Set pptSlide = pptPresentation.Slides.Add(i + 1, ppLayoutText)
        
        ' Add article information to the slide
        Set pptTextbox = pptSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                                    Left:=50, Top:=100, Width:=600, Height:=50)
        pptTextbox.TextFrame.TextRange.Text = "Title: [Insert Article Title]"
        pptTextbox.TextFrame.TextRange.Font.Size = 24
        
        Set pptTextbox = pptSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                                    Left:=50, Top:=200, Width:=600, Height:=50)
        pptTextbox.TextFrame.TextRange.Text = "Authors: [Insert Authors]"
        pptTextbox.TextFrame.TextRange.Font.Size = 20
        
        Set pptTextbox = pptSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                                    Left:=50, Top:=300, Width:=600, Height:=50)
        pptTextbox.TextFrame.TextRange.Text = "Journal/Conference: [Insert Journal/Conference]"
        pptTextbox.TextFrame.TextRange.Font.Size = 20
        
        Set pptTextbox = pptSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                                    Left:=50, Top:=400, Width:=600, Height:=50)
        pptTextbox.TextFrame.TextRange.Text = "Year: [Insert Year]"
        pptTextbox.TextFrame.TextRange.Font.Size = 20
    Next i
    
    ' Show the PowerPoint application
    pptApp.Visible = True
    
    ' Clean up
    Set pptTextbox = Nothing
    Set pptSlide = Nothing
    Set pptPresentation = Nothing
    Set pptApp = Nothing
End Sub
