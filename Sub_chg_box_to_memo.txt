Sub 図形変更_四角をメモに()
  Dim objSlide As Slide
  Dim objShape As Shape

For Each objSlide In ActivePresentation.Slides
    With objSlide
        For Each objShape In .Shapes
            With objShape
                If .TextFrame.HasText = False And (.Name Like "*Rectangle*" Or .Name Like "*メモ*") Then
                    .AutoShapeType = msoShapeFoldedCorner
                    .Line.Visible = msoFalse
                        With .Shadow
                            .OffsetX = -5
                            .OffsetY = 5
                            .ForeColor.RGB = vbBlack
                            .Transparency = 0.12
                            .Obscured = True
                            .Visible = True
                        End With
                End If
            End With
        Next objShape
    End With
Next objSlide
End Sub
