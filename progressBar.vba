Sub BarreDeProgression()
'Génère une barre de progression

'Valeurs à adapter selon besoin
Const Longueur As Single = 0.4    'Longueur totale de la barre (% de  la longueur de la diapo (0.25 =25%))
Const Hauteur As Single = 0.02    'Hauteur totale de la barre (% de  la hauteur de la diapo)
Const PositionX As Single = 0.1   'Position en X de la barre (% de  la longueur de la diapo en partant de la gauche)
Const PositionY As Single = 0.93  'Position en Y de la barre (% de  la hauteur de la diapo en partant de la gauche)


'Récupération des infos
Set Pres = ActivePresentation
H = Pres.PageSetup.SlideHeight
W = Pres.PageSetup.SlideWidth * Longueur
nb = Pres.Slides.Count
counter = 1

'Pour chaque Slide
For Each SLD In Pres.Slides

    If counter = 1 Or counter = nb Then
        GoTo nextLoop
    End If

    'Supprime l'ancienne barre de progression
    nbShape = SLD.Shapes.Count
    del = 0
    For a = 1 To nbShape
        If Left(SLD.Shapes.Item(a - del).Name, 2) = "PB" Then
            SLD.Shapes.Item(a - del).Delete
            del = del + 1
        End If
    Next
    
    'pose la nouvelle barre de progression
    For i = 1 To nb - 1
        Set OBJ = SLD.Shapes.AddShape(msoShapeRectangle, (W * i / nb) + W / nb * (PositionX / 2) - 10, H * (1 - PositionY), (W / nb) * (1 - PositionX), H * Hauteur)
        OBJ.Name = "PB" & i
        OBJ.Line.Visible = msoFalse
        If i + 1 > counter Then
            OBJ.Fill.ForeColor.RGB = RGB(156, 156, 156)
        Else
            OBJ.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End If
    Next
    
nextLoop:
    counter = counter + 1
Next
    
End Sub
