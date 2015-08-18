Sub マスク文字を前面に()

  Dim objSlide As Slide
  Dim objShape As Shape
  
  For Each objSlide In ActivePresentation.Slides
    With objSlide
      For Each objShape In .Shapes
        With objShape
        If .TextFrame.HasText = False And (.Name Like "*Rectangle*" Or .Name Like "*メモ*") Then _
          .ZOrder msoBringToFront
        End With
      Next objShape
    End With
  Next objSlide

End Sub