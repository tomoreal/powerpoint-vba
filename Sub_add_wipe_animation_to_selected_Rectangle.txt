Sub マスクにアニメを選択追加()
'教えて！goo >パワーポイントのアニメーション設定について
'http://oshiete.goo.ne.jp/qa/5240351.html
Dim objSlide As Slide
Dim objShape As Shape
Set objSlide = ActiveWindow.Selection.SlideRange(1)
        For Each objShape In ActiveWindow.Selection.ShapeRange
            With objSlide.TimeLine.InteractiveSequences.Add.AddEffect( _
                        Shape:=objShape, _
                        effectId:=msoAnimEffectWipe, _
                        Trigger:=msoAnimTriggerOnShapeClick)
                .Timing.TriggerShape = objShape
                .Exit = msoTrue
                .EffectParameters.Direction = msoAnimDirectionLeft
            End With
        Next
End Sub
