Sub アニメーション全削除()
'http://qa.itmedia.co.jp/qa3329849.html
    Dim sld As Slide
    Dim shp As Shape
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            shp.AnimationSettings.TextLevelEffect = ppAnimateLevelNone
        Next shp
    Next sld
End Sub
