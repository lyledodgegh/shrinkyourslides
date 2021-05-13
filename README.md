# Shrink your PowerPoint Slides

This is a copy / paste of a link I found a while back on shrinking your PowerPoint slides. When I find the original author/link I'll update here.

All it does is goes through the Slide Masters and deletes unused Slide Masters. Typically this saves a ton of space. If you're really trying to optimize space, you can also unzip the pptx (i.e. treat it like a zip file), and go into ppt/media folder to see what large image files are in your slides and fix those.

Steps:

- Open your PowerPoint deck
- Click on the toolbar View, then Macros
- Type any Macro name in, you'll overwrite it
- Click Create
- Copy the text below, overwriting the content
- Click Run
- Close the PowerPoint deck or Save it
- When asked if you want it to be Macro free - say YES to remove the macro

## The Code




    Sub CleanupDesigns()
        Dim I As Integer
        Dim J As Integer
        Dim oPres As Presentation
        Set oPres = ActivePresentation
        On Error Resume Next
        With oPres
            For I = 1 To .Designs.Count
                For J = .Designs(I).SlideMaster.CustomLayouts.Count To 1 Step -1
                    .Designs(I).SlideMaster.CustomLayouts(J).Delete
                Next
            Next I
        End With
    End Sub



