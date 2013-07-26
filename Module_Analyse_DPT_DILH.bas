Attribute VB_Name = "Module_Analyse_DPT_DILH"
'-------------------------------------------
'Copyright 2013
'Auteur : Simon Verley
'Version : 1.1.6

Sub analyseDefautDPTetDILH()
    Debug.Print
    Debug.Print "================================"
    Debug.Print "Lancement analyseDefautDPTetDILH"
    Debug.Print "================================"
    Dim i As Long
    'Dim jour As Date
    Dim jour_precedent As Date
    'Dim duree_rapi As Double
    'Dim t_deb_rapi As Double
    'Dim t_fin_rapi As Double
    Dim horaires_im() As String
    Dim duree_im As Double
    Dim t_deb_im As Double
    Dim t_fin_im As Double
    Dim NbPP As Integer


    Heure_deb = 5
    Minute_deb = 0
    Heure_fin = 2
    Minute_fin = 20

    C_PT_Confirme = getWordCol("PT_Confirme", 1, True) '10
    If C_PT_Confirme = 0 Then
        MsgBox "Ce fichier n'est pas compatible. (" & ActiveWorkbook.FullName & ")"
        Exit Sub
    End If
    C_Train = C_PT_Confirme
    C_rapi = getWordCol("Red�m_API", 1, True) '20  'RAPI
    C_IM = getWordCol("Info_Maint", 1, True) '21  'IM
    C_Annee = 1
    C_Mois = getWordCol("Mois", 1) '2
    C_Jour = getWordCol("Jour", 1, True) '3
    C_Heure = getWordCol("heure", 1, True) '4
    C_Minute = getWordCol("min", 1, True) '5
    C_Seconde = C_Minute + 1

    i = 2
    last_train = 1 'Valeur par default
    compteur_train = 0
    compteur_train_im = 0
    compteur_train_dilh1 = 0
    compteur_train_dilh2 = 0
    
    last_rapi = 0 'Valeur par default
    compteur_rapi = 0
    'duree_rapi = 0
    horaires_rapi = ""
    
    last_im = 1 'Valeur par default
    compteur_im = 0
    duree_im = 0
    analyse_im = False
    'horaires_im = ""
    ' Acquisition du jour pr�c�dent l'analyse
    jour_precedent = DateSerial(Cells(i, C_Annee), Cells(i, C_Mois), Cells(i, C_Jour))

    If Cells(i, C_Heure) < 12 Then jour_precedent = jour_precedent - 1
    ' D�tection du nombre de PP
    C_Acq = getWordCol("E_Acq", 1, True)
    If C_Acq <> 0 Then
        NbPP = C_Acq - C_PT_Confirme - 1
    Else
        NbPP = getWordCol("E_Def_DPT", 1, True) - C_PT_Confirme - 1
    End If
    ' D�tection colonne Def DTP
    C_Def_DPT = getWordCol("E_Def_DPT", 1, True)
    ' D�tection colonne premier capteur laser
    C_SLG = getWordCol("DonneesRecCor", 1, True)
    'Debug.Print C_SLG
    '
    'Boucle sur toutes les cellules de la colonne A
    'et on sort si on passe Heure_fin:Minute_fin de jour_precedent_2
    '
    For i = 2 To CLng(Range("A65536").End(xlUp).Row)
        jour = DateSerial(Cells(i, C_Annee), Cells(i, C_Mois), Cells(i, C_Jour))
        If jour = jour_precedent Then GoTo Line1
        If jour = jour_precedent + 1 And Cells(i, C_Heure) * 60 + Cells(i, C_Minute) < Heure_deb * 60 + Minute_deb Then GoTo Line1
        If jour = jour_precedent + 2 And Cells(i, C_Heure) * 60 + Cells(i, C_Minute) > Heure_fin * 60 + Minute_fin Then Exit For
        '
        'Compteur trains
        '
        nouveau_train = 0
        valeur = Cells(i, C_Train)
        'Test si changement de valeur
        If valeur <> last_train Then
           'Test si descendant
            If valeur = 0 Then nouveau_train = 1
        End If
        'Memorise l'etat precedent
        last_train = valeur
        compteur_train = compteur_train + nouveau_train
        '
        'Compteur redem API
        '
        valeur = Cells(i, C_rapi)
        'Test si changement de valeur
        If valeur <> last_rapi Then
           'Test si montant ou descendant
           If valeur = 1 Then
              compteur_rapi = compteur_rapi + 1
              't_deb_rapi = jour * 24 * 3600 + Cells(i, C_Heure) * 3600 + Cells(i, C_Minute) * 60 + Cells(i, C_Seconde)
              'horaires_rapi = horaires_rapi & Cells(i, C_Heure) & ":" & Cells(i, C_Minute) & ":" & Cells(i, C_Seconde) & " ; "
              'MsgBox TimeValue(Cells(i, C_Heure) & ":" & Cells(i, C_Minute) & ":" & CInt(Cells(i, C_Seconde)))
              horaires_rapi = horaires_rapi & Format(TimeSerial(Cells(i, C_Heure), Cells(i, C_Minute), CInt(Cells(i, C_Seconde))), "hh:mm") & " ; "
           'ElseIf valeur = 0 Then
              't_fin_rapi = jour * 24 * 3600 + Cells(i, C_Heure) * 3600 + Cells(i, C_Minute) * 60 + Cells(i, C_Seconde)

              'duree_rapi = duree_rapi + t_fin_rapi - t_deb_rapi
              'horaires_rapi = horaires_rapi & " (" & CLng(t_fin_rapi - t_deb_rapi) & "s) ; "

           End If
        End If
        'Memorise l'etat precedent
        last_rapi = valeur
        '
        'Compteur IM
        '
        valeur = Cells(i, C_IM)
        ' D�tection redem API
        If Cells(i + 1, C_rapi) = 1 Then valeur = 1
        'Test si changement de valeur
        If valeur <> last_im Then
            'Test si montant ou descendant
            If valeur = 0 And last_rapi = 0 Then
                analyse_im = True
                compteur_im = compteur_im + 1
                t_deb_im = jour * 24 * 3600 + Cells(i, C_Heure) * 3600 + Cells(i, C_Minute) * 60 + Cells(i, C_Seconde)
                ReDim Preserve horaires_im(1 To compteur_im)
                horaires_im(compteur_im) = Format(TimeSerial(Cells(i, C_Heure), Cells(i, C_Minute), CInt(Cells(i, C_Seconde))), "hh:mm") '& " ; "
                Dim PP_Def_DPT() As String
                ReDim PP_Def_DPT(0 To 2)
                Dim PP_Def_DILH() As String
                ReDim PP_Def_DILH(1 To NbPP * 9)
                'NbPP_Def = 0
            ElseIf valeur = 1 And analyse_im Then
                analyse_im = False
                t_fin_im = jour * 24 * 3600 + Cells(i, C_Heure) * 3600 + Cells(i, C_Minute) * 60 + Cells(i, C_Seconde)
                duree_im = duree_im + t_fin_im - t_deb_im
                horaires_im(compteur_im) = horaires_im(compteur_im) & " (" & TimeString(CLng(t_fin_im - t_deb_im)) & ") ; " & " D�fauts : " & Chr(10)   '" (" & CInt(t_fin_im - t_deb_im) & "s) ; "
                'Debug.Print ">>>>>>>>>>>>>>>>>>>>>>"
                defaut = False
                premier = True
                For Each d In PP_Def_DPT
                    If d <> "" Then
                        'NbPP_Def = NbPP_Def + 1
                        If premier Then
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "    DPT : " & Chr(10)
                            premier = False
                        End If
                        defaut = True
                        horaires_im(compteur_im) = horaires_im(compteur_im) & "      - " & d & " ; " & Chr(10)
                    End If
                Next d
                'If Not premier Then horaires_im(compteur_im) = horaires_im(compteur_im) & Chr(10)
                premier = True
                For Each d In PP_Def_DILH
                    If d <> "" Then
                        'NbPP_Def = NbPP_Def + 1
                        If premier Then
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "    DILH : " & Chr(10)
                            premier = False
                        End If
                        defaut = True
                        horaires_im(compteur_im) = horaires_im(compteur_im) & "      - " & d & " ; " & Chr(10)
                    End If
                Next d
                If Not defaut Then horaires_im(compteur_im) = horaires_im(compteur_im) & "      Aucun d�faut d�tect� (Red�marrage API ?)" & Chr(10)
                'Debug.Print "<<<<<<<<<<<<<<<<<<<<<<"
                horaires_im(compteur_im) = horaires_im(compteur_im) & Chr(10)
            Else
                analyse_im = False
            End If
        End If
        ' Analyse des defauts
        'If i = 12867 Then Debug.Print " i = 12867 : " & jour & " " & Format(Cells(i, C_Heure), "00") & ":" & Cells(i, C_Minute) & " " & valeur
        If analyse_im And valeur = 0 Then
            For c = 0 To 2
'                If i = 14737 Then Debug.Print " PP_Def_DPT(c) >" & PP_Def_DPT(c) & "<"
                If Cells(i, C_Def_DPT + c) = 0 And PP_Def_DPT(c) = "" Then
                    PP_Def_DPT(c) = Cells(1, C_Def_DPT + c) & " (" & Cells(i, C_Heure) & ":" & Format(Cells(i, C_Minute), "00") & ")"
                End If
            Next c
            nb_capteur = 0
            For pp = 1 To NbPP
                nb_capteur_pp = 0
                For c = 1 To 3
                    def_capteur = False
                    For d = 1 To 3
                        dcpp = (pp - 1) * 3 + (c - 1) * 3 + d
                        C_dcpp = C_SLG + (pp - 1) * 12 + (c - 1) * 4 + d
                        'If i = 12867 Then Debug.Print " pp=" & pp & " c=" & c & " d=" & d & " dcpp=" & dcpp & " C_dcpp=" & C_dcpp & " " & Cells(i, C_dcpp) & " " & Cells(1, C_dcpp)
                        If Cells(i, C_dcpp) = 1 Then
                            'Debug.Print i & " " & cpp
                            def_capteur = True
                            If PP_Def_DILH(dcpp) = "" Then PP_Def_DILH(dcpp) = Cells(1, C_PT_Confirme + pp) & " " & Cells(1, C_dcpp) & " (" & Cells(i, C_Heure) & ":" & Format(Cells(i, C_Minute), "00") & ")"
                        End If
                    Next d
                    If def_capteur Then nb_capteur_pp = nb_capteur_pp + 1
                Next c
                nb_capteur = WorksheetFunction.Max(nb_capteur, nb_capteur_pp)
            Next pp
            If nb_capteur = 1 Then
                compteur_train_dilh1 = compteur_train_dilh1 + nouveau_train
            ElseIf nb_capteur > 1 Then
                compteur_train_dilh2 = compteur_train_dilh2 + nouveau_train
            End If
            If nb_capteur > 0 Then compteur_train_im = compteur_train_im + nouveau_train
        End If
        'Memorise l'etat precedent
        last_im = valeur

Line1:
        'Incr�mente la variable d'une unit� afin de tester la cellule suivante
    Next i
    
    Rapport = _
           " Rapport d'analyse du fichier : " & ActiveWorkbook.FullName & Chr(10) & Chr(10) & _
           "Trains : " & Chr(10) & _
           "  " & compteur_train & " passages de train." & Chr(10) & _
           "  dont " & compteur_train_im & " pendant Info Maintenance." & Chr(10)
    If compteur_train_dilh1 > 0 Then Rapport = Rapport & _
           "     DILH d�faut simple : " & compteur_train_dilh1 & Chr(10)
    If compteur_train_dilh2 > 0 Then Rapport = Rapport & _
           "     DILH d�faut double : " & compteur_train_dilh2 & Chr(10)
    Rapport = Rapport & Chr(10) & _
           "Red�marrage API : " & Chr(10) & _
           "  " & compteur_rapi & " Red�marrage d'API." & Chr(10)
    If compteur_rapi > 0 Then Rapport = Rapport & "  Horaires : " & horaires_rapi & Chr(10)
    Rapport = Rapport & Chr(10) & _
           "Info Maintenance : " & Chr(10) & _
           "  " & compteur_im & " Info Maintenance pour une dur�e de " & TimeString(CLng(duree_im)) & "." & Chr(10)
    If compteur_im > 0 Then
        Rapport = Rapport & "  Horaires : " & Chr(10)
        For Each Line In horaires_im
            Rapport = Rapport & "  @ " & Line
        Next Line
    End If
    
    'MsgBox "Trains : " & Chr(10) & _
           "       " & compteur_train & " passages de train." & Chr(10) & _
           "       dont " & compteur_train_im & " pendant Info Maintenance." & Chr(10) & Chr(10) & _
           "Redemarrage API : " & Chr(10) & _
           "       " & compteur_rapi & " Redemarrage d'API." & Chr(10) & _
           "       Horaires : " & horaires_rapi & Chr(10) & Chr(10) & _
           "Info Maintenance : " & Chr(10) & _
           "       " & compteur_im & " Info Maintenance pour une duree de " & TimeString(CLng(duree_im)) & "." & Chr(10) & _
           "       Horaires : " & Chr(10) & h _
           , vbOKOnly, ActiveWorkbook.Name

    Fichier_sortie = ActiveWorkbook.FullName & "_Analyse.txt"
    'Fichier_sortie = "Analyse.txt"

    ' Declare a FileSystemObject.
    'Dim fso As New FileSystemObject
    ' Declare a TextStream.
    'Dim stream As TextStream
    ' Create a TextStream.
    'Set stream = fso.CreateTextFile(Fichier_sortie, True)
    'stream.WriteLine " Rapport d'analyse du fichier : " & ActiveWorkbook.FullName
    'stream.WriteLine "Trains : "
    'stream.WriteLine "  " & compteur_train & " passages de train."
    'stream.WriteLine "  dont " & compteur_train_im & " pendant Info Maintenance."
    'stream.WriteLine ""
    'stream.WriteLine "Red�marrage API : "
    'stream.WriteLine "  " & compteur_rapi & " Red�marrage d'API."
    'stream.WriteLine "  Horaires : " & horaires_rapi
    'stream.WriteLine ""
    'stream.WriteLine "Info Maintenance : "
    'stream.WriteLine "  " & compteur_im & " Info Maintenance pour une dur�e de " & TimeString(CLng(duree_im)) & "."
    'stream.WriteLine "  Horaires : "
    'stream.WriteLine h
    ' Close the file.
    'stream.Close

    file2Write = FreeFile() ' assign next free file number to this variable
    Open Fichier_sortie For Output As file2Write ' output is for writing to a file
    Print #file2Write, Rapport
    Close #file2Write

    MsgBox "Le rapport d'analyse a �t� g�n�r� dans le fichier " & Fichier_sortie '& " dans Mes Documents"

    'commande Pour XP, pour W7 cherchez le chemin
    Shell "C:\WINDOWS\explorer.exe " & Fichier_sortie

End Sub

'-----------------------------------------------------

Public Function getWordCol(ByVal sExpression As String, ByVal iLineNumber As Integer, Optional ByVal bPartial As Boolean = False, Optional ByVal bSelectResult As Boolean = False, Optional vsSheetName As Variant) As Integer
'   sExpression       mot(s) ou partie de mot � chercher
'   iLineNumber      numero de la ligne dans laquelle chercher
'   bPartial              choix sur le mot comlet ou partie du  mot
'   bSelectResult     s�lectionner la cellule de  la premi�re occurence trouv�e
'   vsSheetName     nom de la feuille dans laquelle cherche, celle active par d�faut
'   RETURN              num�ro de la colonne de la premi�re occurence trouv�e
    Dim iColStop    As Integer
    Dim i           As Integer

    'selection  feuille
    If Not IsMissing(vsSheetName) Then Sheets(vsSheetName).Select

    'derni�re cellule
    iColStop = Range("iv1").End(xlToLeft).Column

    If bPartial Then
        For i = 1 To iColStop
            If Cells(iLineNumber, i) Like "*" & sExpression & "*" Then
                getWordCol = i
                If bSelectResult Then Cells(iLineNumber, i).Select
                Exit For
            End If
        Next i
    Else
        For i = 1 To iColStop
            If Cells(iLineNumber, i) = sExpression Then
                getWordCol = i
                If bSelectResult Then Cells(iLineNumber, i).Select
                Exit For
            End If
        Next i
    End If
End Function

Public Function TimeString(Secondes As Long) As String

    Dim nb_heure As Long
    Dim nb_minute As Integer
    Dim nb_seconde As Integer

    nb_heure = CLng(Secondes / 3600 - 0.5)
    nb_minute = CInt((Secondes - nb_heure * 3600) / 60 - 0.5)
    nb_seconde = CInt(Secondes - nb_heure * 3600 - nb_minute * 60)

    TimeString = nb_heure & "h" & nb_minute & "m" & nb_seconde & "s"
    'MsgBox TimeString
    'TimeString = Format(TimeValue(TimeString), Fmt)

End Function



