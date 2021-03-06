Attribute VB_Name = "Module_Analyse_DPT_DILH"
Public Const Version = "1.3.6"
'=====================
'Copyright 2013
'Auteur  : Simon Verley
'Version : 1.3.6
' BUG        : - Correction de la recherche du nombre de col relai et du nombre de col alim
'Version : 1.3.5
' BUG        : - Correction de la recherche de la derniere colonne (getWordLastCol)
'Version : 1.3.4
' BUG        : - Correction de la recherche de date
'Version : 1.3.3
' BUG        : - Correction du compteur de defaut dilh
' Parametrage: - Modification du delai seuil de defaut DPT (passage a 1.5)
'Version : 1.3.2
' Nouveaut�s : - Report des passages de trains dans completeSuivi
'Version : 1.3.1
' BUG        : - Correction de la detection du quai dans completeSuivi
'              - Correction de la date de report
' Nouveaut�s : - Report des defauts Alim et Relai dans completeSuivi
'              - Modification du nom du fichier de report dans completeSuivi (ajout d'un suffix de la version de cette macro)
'Version : 1.3.0
' Nouveaut�s : - Detection des defauts Alim et Relai (ToDO: les faire parvenir jusqu'a completeSuivi)
'Version : 1.2.8
' BUG : Verifie l'existance des colonnes COPP_SAET et PPFV
' Embellissement :
'              - Mise en forme de la colonne IFPP
'              - Uniformisation des mises en forme des colonnes Diag et Defaut
'              - Message d'arrivee en derniere ligne modifie
'Version : 1.2.7
' Nouveaut�s : - comptage des passages de train sans PPFV ou COPP_SAET
' Embellissement :
'              - Mise en forme de la colonne COPP_SAET
'              - Horaire du dernier train et non derniere ligne < Heure de fin
'              - Message si la macro arrive en derniere ligne (Heure de fin > derniere ligne)
'              - Message si train encore a quai apres heure de fin (stationnement...)
'Version : 1.2.6
' BUG : le filtrage des default DTP < 1.2s est encore am�lior�
' Nouveaut�s : - comptage des d�faults dilh et dpt s�par� et report dans Suivi d�fault DIL.xls
'              - remplissage du Suivi d�fault DIL.xls plus complet
'Version : 1.2.5
' Nouveaut�s : - traiteFichier devient completeSuivi
'              - macro de mise en forme automatique et conditionnelle
'              - mise en forme appliqu�e dans analyseFichier
'Version : 1.2.4
' BUG : le filtrage des default DTP < 1.2s est plus complet
'Version pr�c�dente : 1.2.3
' BUG : le comptage des trains en cours d'IM n'est plus bas� sur les d�tections lasers
' Nouveaut�s : Filtrages des defauts DPT < 1.2s
' Embellissement : Message si 1 IM sans default d�tect�
'=====================

Sub analyseFichier()
'   --------------
    Dim jour As Date
    Dim Station As String
    Dim Quai As String
    Dim compteur_train As Integer
    Dim compteur_train_im As Integer
    Dim compteur_train_dilh1 As Integer
    Dim compteur_train_dilh2 As Integer
    Dim horaires_train As String
    Dim horaires_default_train As String
    Dim compteur_rapi As Integer
    Dim horaires_rapi As String
    Dim compteur_im As Integer
    Dim duree_im As Long
    Dim horaires_im() As String
    Dim compteur_dilh As Integer
    Dim compteur_dpt As Integer
    Dim compteur_alim As Integer
    Dim compteur_rel As Integer
    
    If analyseDefautDPTetDILH(jour, Quai, compteur_train, compteur_train_im, compteur_train_dilh1, compteur_train_dilh2, horaires_train, horaires_default_train, _
        compteur_rapi, horaires_rapi, compteur_im, duree_im, horaires_im, compteur_dilh, compteur_dpt, compteur_alim, compteur_rel) _
        <> 0 Then Exit Sub
    
    genereRapport compteur_train, compteur_train_im, compteur_train_dilh1, compteur_train_dilh2, horaires_train, horaires_default_train, _
        compteur_rapi, horaires_rapi, compteur_im, duree_im, horaires_im, compteur_dilh, compteur_dpt, compteur_alim, compteur_rel
    
    metForme
    
    MsgBox "Le rapport d'analyse du fichier " & ActiveWorkbook.FullName & " a �t� g�n�r� (suffixe '_Analyse.txt adjoint')."

End Sub

Sub completeSuivi()
'    ------------
    Dim jour As Date
    Dim Station As String
    Dim Quai As String
    Dim compteur_train As Integer
    Dim compteur_train_im As Integer
    Dim compteur_train_dilh1 As Integer
    Dim compteur_train_dilh2 As Integer
    Dim horaires_train As String
    Dim horaires_default_train As String
    Dim compteur_rapi As Integer
    Dim horaires_rapi As String
    Dim compteur_im As Integer
    Dim duree_im As Long
    Dim horaires_im() As String
    Dim compteur_dilh As Integer
    Dim compteur_dpt As Integer
    Dim compteur_alim As Integer
    Dim compteur_rel As Integer
    
    If analyseDefautDPTetDILH(jour, Quai, compteur_train, compteur_train_im, compteur_train_dilh1, compteur_train_dilh2, horaires_train, horaires_default_train, _
        compteur_rapi, horaires_rapi, compteur_im, duree_im, horaires_im, compteur_dilh, compteur_dpt, compteur_alim, compteur_rel) _
        <> 0 Then Exit Sub
    
    genereRapport compteur_train, compteur_train_im, compteur_train_dilh1, compteur_train_dilh2, horaires_train, horaires_default_train, _
        compteur_rapi, horaires_rapi, compteur_im, duree_im, horaires_im, compteur_dilh, compteur_dpt, compteur_alim, compteur_rel, False
    
    
    ' D�tection de la station � partir du nom de fichier
    If InStr(ActiveWorkbook.Name, "BAST") > 0 Or InStr(ActiveWorkbook.Name, "BASTILLE") > 0 Or InStr(ActiveWorkbook.Name, "Bastille") > 0 Then
        Station = "BASTILLE"
    ElseIf InStr(ActiveWorkbook.Name, "NATN") > 0 Or InStr(ActiveWorkbook.Name, "NATION") > 0 Or InStr(ActiveWorkbook.Name, "Nation") > 0 Then
        Station = "NATION"
    ElseIf InStr(ActiveWorkbook.Name, "CHGE") > 0 Or InStr(ActiveWorkbook.Name, "ETOILE") > 0 Or InStr(ActiveWorkbook.Name, "Etoile") > 0 Then
        Station = "ETOILE"
    Else
        MsgBox "Le nom de la station n'a pas pu �tre extrait du nom du fichier de donn�e."
        'Exit Sub
    End If
        
    
    Fichier_suivi = "Suivi d�faut DIL_" & Version & ".xls"
    
    On Error GoTo ErrHandler:
        Windows(Fichier_suivi).Activate
ErrHandler:
    'MsgBox Err.Number
    If Err.Number = 9 Then
        test = MsgBox(Prompt:="Le fichier de suivi est-il d�j� ouvert ?", _
            Buttons:=vbYesNoCancel, Title:="Fichier de suivi")
        If test = vbYes Then
            Fichier_suivi = InputBox(Prompt:="Veuiller entrer le nom du fichier de suivi :", _
          Title:="Fichier de suivi", Default:="Suivi d�faut DIL.xls")
        ElseIf test = vbNo Then
            Fichier_suivi = Application.GetOpenFilename _
             (FileFilter:="Fichier de suivi (*.xls),*.xls,Tout type (*.*),*.*", _
             Title:="Ouvrir le fichier de suivi", MultiSelect:=False)
            If Fichier_suivi = False Then Exit Sub
            Workbooks.Open Fichier_suivi
        Else
            Exit Sub
        End If
        Resume Next
    End If

    On Error GoTo StationErrHandler:
        Worksheets(Station).Activate
StationErrHandler:
    'MsgBox Err.Number
    If Err.Number = 9 Then
        Station = InputBox(Prompt:="La feuille '" & Station & "' n'a pas pu �tre trouv�e dans le fichier " & Fichier_suivi & Chr(10) _
          & "Veuiller entrer le nom de la feuille correspondante � la station analys�e :", _
          Title:="Nom de la station", Default:=Station)
        Resume Next
    End If
    
    C_Date = getWordCol(jour, 1)
    If C_Date < 2 Then
        MsgBox "La date " & jour & " n'a pas �t� trouv�e dans le fichier " & Fichier_suivi
        Exit Sub
    End If
    nb_lignes_quai = 44
    L_Quai = getWordLine(Quai, 1, True)
    If L_Quai = 0 Then
        MsgBox "Le quai '" & Quai & "' n'a pas �t� trouv� dans le fichier " & Fichier_suivi
        Exit Sub
    End If
    Pos_Q1 = getWordLine("Q", 1, True, True)
    Pos_AlarmIM = getWordLine("Alarme IM", 1, False, True) - Pos_Q1
    Pos_Duree = getWordLine("Dur�e de la panne", 1, False, True) - Pos_Q1
    Pos_DefautSL = getWordLine("D�faut SL", 1, False, True) - Pos_Q1
    Pos_DefautDPT = getWordLine("D�faut DPT", 1, False, True) - Pos_Q1
    Pos_DefautUD1N = getWordLine("D�faut autre UD1N", 1, False, True) - Pos_Q1
    'Pos_IM = getWordLine("Intervention Maintenance", 1, False, True) - Pos_Q1
    Pos_RedemAPI = getWordLine("R�dem API", 1, False, True) - Pos_Q1
    'Pos_ChgmtUD1N = getWordLine("Changement UD1N", 1, False, True) - Pos_Q1
    Pos_NbTrains = getWordLine("Nombre echange voyageurs total", 1, False, True) - Pos_Q1
    Pos_NbTrainsDouble = getWordLine("Nb �ch voy avec exploitant", 1, False, True) - Pos_Q1
    Pos_NbTrainsIM = getWordLine("Nb �ch voy IM active (Fonctionnement d�grad� 2/2)", 1, False, True) - Pos_Q1
    
    Cells(L_Quai + Pos_AlarmIM, C_Date) = compteur_im
    Cells(L_Quai + Pos_Duree, C_Date) = Format(SecondsToDate(duree_im&), "hh:mm")
    Cells(L_Quai + Pos_DefautSL, C_Date) = compteur_dilh
    Cells(L_Quai + Pos_DefautDPT, C_Date) = compteur_dpt
    Cells(L_Quai + Pos_DefautUD1N, C_Date) = compteur_alim + compteur_rel
    'Cells(L_Quai + 8, C_Date) = compteur_im
    Cells(L_Quai + Pos_RedemAPI, C_Date) = compteur_rapi
    
    Cells(L_Quai + Pos_NbTrains, C_Date) = compteur_train
    Cells(L_Quai + Pos_NbTrainsIM, C_Date) = compteur_train_im
    Cells(L_Quai + Pos_NbTrainsDouble, C_Date) = compteur_train_dilh2
    
    'C_Date = getWordCol(
    'Range("A1") = Rapport
End Sub

Sub metForme()
'
' Macro1 Macro
'
'
    ' le nom des tableaux suivant correspond a "Couleur si 0"+"Couleur si 1"
    VertGris = Array("PPFV*", "E_PT_MP05*", "PT_Confirme*")
    GrisVert = Array("E_PT_MP89*", "E_Acq_*", "E_Acq_*", "Red�m_API", "COPP_SAET")
    GrisRouge = Array("SLG_PP*", "SLC_PP*", "SLD_PP*", "E_DILF_SL*", "UTG*", "UTC*", "UTD*", "*.Defaut_Dyna", "*.Defaut_SurfRef", "*.Incoherent", "Dyn*_DF*", "SF*_DF*", "DFQ*_SL*", "Defaut_*")
    RougeGris = Array("UTH*", "Att_acq")
    RougeVert = Array("Info_Maint", "E_Def_DPT*", "Diag_*", "*.DonneesRecCor", "IFPP_*")
    
    C_PT_Confirme = getWordCol("PT_Confirme", 1, True) '10
    If C_PT_Confirme = 0 Then
        MsgBox "Ce fichier n'est pas compatible. (" & ActiveWorkbook.FullName & ")"
        Exit Sub
    End If
    C_PPFV = getWordCol("PPFV", 1, True)
    ' D�tection du nombre de PP
    If getWordCol("Att_acq", 1, False) <> 0 Then
        NbPP = getWordCol("Att_acq", 1, False) - C_PT_Confirme - 1
    ElseIf getWordCol("E_Acq", 1, True) <> 0 Then
        NbPP = getWordCol("E_Acq", 1, True) - C_PT_Confirme - 1
    Else
        NbPP = getWordCol("E_Def_DPT", 1, True) - C_PT_Confirme - 1
    End If
    
    Application.ScreenUpdating = False
    ' Application du filtre automatique
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.UsedRange.AutoFilter
    End If
    
    With Rows("1:1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 90
    End With
    Rows("1:1").Interior.ColorIndex = 15
    ' Figeage des volets
    Range("G2").Activate
    ActiveWindow.FreezePanes = True
    ' Colonnes dates et heure
    Columns("A:F").ColumnWidth = 5
    Columns("D:F").Interior.ColorIndex = 15
    
    Dim c As Integer
    ' Boucle sur colonne
    ColSansMEF = ""
    For c = C_PPFV To Range("iv1").End(xlToLeft).Column
        If IsInArray(Cells(1, c), RougeVert) Then
            appliqueMEFC c, 3, 4
        ElseIf IsInArray(Cells(1, c), GrisVert) Then
            appliqueMEFC c, 15, 43
        ElseIf IsInArray(Cells(1, c), VertGris) Then
            appliqueMEFC c, 43, 15
        ElseIf IsInArray(Cells(1, c), RougeGris) Then
            appliqueMEFC c, 53, 15
        ElseIf IsInArray(Cells(1, c), GrisRouge) Then
            appliqueMEFC c, 15, 53
        ElseIf c > C_PT_Confirme And c <= C_PT_Confirme + NbPP Then
        Else
            ColSansMEF = ColSansMEF & " " & Cells(1, c)
        End If
    Next c
    For c = C_PT_Confirme + 1 To C_PT_Confirme + NbPP
        appliqueMEFC c, 53, 15
    Next c
    If Not ColSansMEF = "" Then MsgBox "Les colonnes " & ColSansMEF & " n'ont pas �t� mises en forme."
    
    Application.ScreenUpdating = True
    
End Sub

'====================================================================================================================

Private Sub appliqueMEFC(col As Integer, couleur_0 As Integer, couleur_1 As Integer) 'couleur() As String)
    Columns(col).Select
    Selection.ColumnWidth = 3#
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
    Selection.FormatConditions(1).Interior.ColorIndex = couleur_0
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="1"
    Selection.FormatConditions(2).Interior.ColorIndex = couleur_1
End Sub

Function analyseDefautDPTetDILH(ByRef jour As Date, ByRef Quai As String, _
        ByRef compteur_train As Integer, ByRef compteur_train_im As Integer, ByRef compteur_train_dilh1 As Integer, ByRef compteur_train_dilh2 As Integer, ByRef horaires_train As String, ByRef horaires_default_train As String, _
        ByRef compteur_rapi As Integer, ByRef horaires_rapi As String, _
        ByRef compteur_im As Integer, ByRef duree_im As Long, ByRef horaires_im() As String, ByRef compteur_dilh As Integer, ByRef compteur_dpt As Integer, ByRef compteur_alim As Integer, ByRef compteur_rel As Integer _
        ) As Integer
        
    Debug.Print
    Debug.Print "================================"
    Debug.Print "Lancement analyseDefautDPTetDILH"
    Debug.Print "================================"
    
    Dim i As Long
    Dim jour_precedent As Date
    Dim t_deb_im As Double
    Dim t_fin_im As Double
    Dim NbPP As Integer


    Heure_deb = 5
    Minute_deb = 0
    Heure_fin = 2
    Minute_fin = 20
    Delai_defaut_DPT = 1.5

    C_PT_Confirme = getWordCol("PT_Confirme", 1, True) '10
    If C_PT_Confirme = 0 Then
        analyseDefautDPTetDILH = 1
        MsgBox "Ce fichier n'est pas compatible. (" & ActiveWorkbook.FullName & ")"
        Exit Function
    End If
    
    Quai = Split(Cells(1, C_PT_Confirme), "_")(UBound(Split(Cells(1, C_PT_Confirme), "_")))
    
    C_COPP = getWordCol("COPP_SAET", 1, True) '10
    C_PPFV = getWordCol("PPFV", 1, True) '10
    C_rapi = getWordCol("Red�m_API", 1, True) '20  'RAPI
    C_IM = getWordCol("Info_Maint", 1, True) '21  'IM
    C_Annee = 1
    C_Mois = getWordCol("Mois", 1) '2
    C_Jour = getWordCol("Jour", 1, True) '3
    C_Heure = getWordCol("heure", 1, True) '4
    C_Minute = getWordCol("min", 1, True) '5
    C_Seconde = C_Minute + 1
    
    C_Train = C_PT_Confirme

    i = 2
    last_train = 1 'Valeur par default
    compteur_train = 0
    compteur_train_im = 0
    compteur_train_dilh1 = 0
    compteur_train_dilh2 = 0
    horaires_train = ""
    horaires_default_train = ""
    
    last_rapi = 0 'Valeur par default
    compteur_rapi = 0
    'duree_rapi = 0
    horaires_rapi = ""
    
    last_im = 1 'Valeur par default
    compteur_im = 0
    duree_im = 0
    analyse_im = False
    prof_recherche = 100
    compteur_dilh = 0
    compteur_dpt = 0
    compteur_alim = 0
    compteur_rel = 0
    
    ' Acquisition du jour pr�c�dent l'analyse
    jour_precedent = DateSerial(Cells(i, C_Annee), Cells(i, C_Mois), Cells(i, C_Jour))

    If Cells(i, C_Heure) < 12 Then jour_precedent = jour_precedent - 1
    ' D�tection colonne Def DTP
    C_Def_DPT = getWordCol("E_Def_DPT", 1, True)
    ' D�tection colonne premier capteur laser
    C_SLG = getWordCol("DonneesRecCor", 1, True)
    C_SLG_f = getWordLastCol("DonneesRecCor", 1, True)
    ' D�tection du nombre de PP
    NbPP = (C_SLG_f + 3 - C_SLG + 1) / 12
    'If getWordCol("Att_acq", 1, False) <> 0 Then
    '    NbPP = getWordCol("Att_acq", 1, False) - C_PT_Confirme - 1
    'ElseIf getWordCol("E_Acq", 1, True) <> 0 Then
    '    NbPP = getWordCol("E_Acq", 1, True) - C_PT_Confirme - 1
    'Else
    '    NbPP = getWordCol("E_Def_DPT", 1, True) - C_PT_Confirme - 1
    'End If
    ' D�tection colonne premier defaut alim
    C_ALIM = getWordCol("Diag_Alim", 1, True)
    C_ALIM_f = getWordLastCol("Diag_Bat", 1, True)
    ' D�tection du nombre de capteur alim
    NbALIM = C_ALIM_f - C_ALIM + 1
    ' D�tection colonne premier defaut relais
    'C_REL = getWordCol("Defaut_Acq", 1, True)
    C_REL = C_ALIM_f + 1
    ' D�tection du nombre de capteur relais
    NbREL = getWordLastCol("Defaut_Rel", 1, True) - C_REL + 1
    
    '
    'Boucle sur toutes les cellules de la colonne A
    'et on sort si on passe Heure_fin:Minute_fin de jour_precedent+2
    '
    flag_ppfv = False
    flag_copp = False
    Lasti = False
    Firsti = 0
    NbLignes = CLng(Range("A65536").End(xlUp).Row)
    Debug.Print "  Jour      : " & jour_precedent + 1
    Debug.Print "  Nb Lignes : " & NbLignes
    Debug.Print "  Nb PP     : " & NbPP & " " & C_SLG & " " & C_SLG_f
    Debug.Print "  Nb Alim   : " & NbALIM & " " & C_ALIM & " " & getWordLastCol("Diag_Bat", 1, True)
    Debug.Print "  Nb Rel    : " & NbREL & " " & C_REL & " " & getWordLastCol("Defaut_Rel", 1, True)
    For i = 2 To NbLignes
        jour = DateSerial(Cells(i, C_Annee), Cells(i, C_Mois), Cells(i, C_Jour))
        heure = jour * 24 * 3600 + Cells(i, C_Heure) * 3600 + Cells(i, C_Minute) * 60 + Cells(i, C_Seconde)
        If jour = jour_precedent Then
            GoTo Nexti
        ElseIf jour = jour_precedent + 1 And Cells(i, C_Heure) * 60 + Cells(i, C_Minute) < Heure_deb * 60 + Minute_deb Then
            GoTo Nexti
        ElseIf Firsti = 0 Then
            Firsti = i
        End If
        If (jour = jour_precedent + 2 And Cells(i + 1, C_Heure) * 60 + Cells(i + 1, C_Minute) > Heure_fin * 60 + Minute_fin) Or i = NbLignes Then
            'If i = NbLignes Then MsgBox "La macro a atteint la derniere ligne du fichier. Deux raisons peuvent � l'origine de ce message : " & Chr(10)
            If i = 65536 Then
                MsgBox "Le fichier est trop long pour une analyse compl�te." & Chr(10) _
              & "    > Veuillez effacer les lignes de 2 � " & Firsti & " en dehors d'Excel (via Bloc-notes par exemple) et relancer l'analyse ensuite."
            'ElseIf i = NbLignes Then
                'MsgBox "Le fichier de log se termine avant l'heure de fin d'analyse :" & Chr(10) _
              & "    > V�rifiez que la derni�re ligne du log correspond � l'horaire " & last_horaire_train & Chr(10)
            End If
            If Not Lasti Then
                Lasti = True
                horaires_train = horaires_train & " � " & last_horaire_train & " le " & jour
                If Cells(i, C_Train) = 0 Then horaires_train = horaires_train & " (train encore � quai � " & Heure_fin & "h" & Minute_fin & ")"
                GoTo FinAnalyse
            End If
            Exit For
        End If
        '
        'Compteur trains
        '
        nouveau_train = 0
        valeur = Cells(i, C_Train)
        'Test si changement de valeur
        If valeur <> last_train Then
           'Test si descendant
            If valeur = 0 Then
                nouveau_train = 1
                last_horaire_train = Format(TimeSerial(Cells(i, C_Heure), Cells(i, C_Minute), CInt(Cells(i, C_Seconde))), "hh:mm:ss")
                If compteur_train = 0 Then horaires_train = " de " & last_horaire_train & " le " & jour
            ElseIf valeur = 1 Then
                If Not flag_ppfv Or Not flag_copp Then
                    horaires_default_train = horaires_default_train & last_horaire_train & " ("
                    If Not flag_ppfv Then horaires_default_train = horaires_default_train & "PPFV"
                    If Not flag_copp Then horaires_default_train = horaires_default_train & " COPP"
                    horaires_default_train = horaires_default_train & ") ; "
                End If
                flag_ppfv = False
                flag_copp = False
            End If
        End If
        'Memorise l'etat precedent
        last_train = valeur
        compteur_train = compteur_train + nouveau_train
        'Verification si possible de coherence PPFV / COPP_SAET avec PT_confirm
        If valeur = 0 Then
            If C_PPFV = 0 Then
                flag_ppfv = True
            ElseIf Cells(i, C_PPFV) = 0 Then
                flag_ppfv = True
            End If
            If C_COPP = 0 Then
                flag_copp = True
            ElseIf Cells(i, C_COPP) = 1 Then
                flag_copp = True
            End If
        End If
        '
        'Compteur redem API
        '
        valeur = Cells(i, C_rapi)
        'Test si changement de valeur
        If valeur <> last_rapi Then
           'Test si montant ou descendant
           If valeur = 1 Then
              compteur_rapi = compteur_rapi + 1
              horaires_rapi = horaires_rapi & Format(TimeSerial(Cells(i, C_Heure), Cells(i, C_Minute), CInt(Cells(i, C_Seconde))), "hh:mm") & " ; "
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
                t_deb_im = heure
                ReDim Preserve horaires_im(1 To compteur_im)
                horaires_im(compteur_im) = Format(TimeSerial(Cells(i, C_Heure), Cells(i, C_Minute), CInt(Cells(i, C_Seconde))), "hh:mm") '& " ; "
                Dim PP_Def_DPT() As String
                ReDim PP_Def_DPT(0 To 2)
                Dim PP_Def_DILH() As String
                ReDim PP_Def_DILH(1 To NbPP * 9)
                Dim PP_Def_ALIM() As String
                ReDim PP_Def_ALIM(0 To NbALIM)
                Dim PP_Def_REL() As String
                ReDim PP_Def_REL(0 To NbREL)
                'NbPP_Def = 0
            ElseIf valeur = 1 Then
FinAnalyse:
                If analyse_im Then
                    analyse_im = False
                    t_fin_im = heure
                    duree_im = duree_im + t_fin_im - t_deb_im
                    horaires_im(compteur_im) = horaires_im(compteur_im) & " (" & TimeString(CLng(t_fin_im - t_deb_im)) & ") ; " & " D�fauts : " & Chr(10)   '" (" & CInt(t_fin_im - t_deb_im) & "s) ; "
                    defaut = False
                    premier = True
                    For Each d In PP_Def_DPT
                        If d <> "" Then
                            If premier Then
                                horaires_im(compteur_im) = horaires_im(compteur_im) & "    DPT : " & Chr(10)
                                premier = False
                            End If
                            defaut = True
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "      - " & d & " ; " & Chr(10)
                        End If
                    Next d
                    premier = True
                    For Each d In PP_Def_DILH
                        If d <> "" Then
                              If premier Then
                                horaires_im(compteur_im) = horaires_im(compteur_im) & "    DILH : " & Chr(10)
                                premier = False
                            End If
                            defaut = True
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "      - " & d & " ; " & Chr(10)
                        End If
                    Next d
                    premier = True
                    For Each d In PP_Def_ALIM
                        If d <> "" Then
                              If premier Then
                                horaires_im(compteur_im) = horaires_im(compteur_im) & "    ALIM : " & Chr(10)
                                premier = False
                            End If
                            defaut = True
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "      - " & d & " ; " & Chr(10)
                        End If
                    Next d
                    premier = True
                    For Each d In PP_Def_REL
                        If d <> "" Then
                              If premier Then
                                horaires_im(compteur_im) = horaires_im(compteur_im) & "    RELAIS : " & Chr(10)
                                premier = False
                            End If
                            defaut = True
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "      - " & d & " ; " & Chr(10)
                        End If
                    Next d
                    If Not defaut Then
                        If compteur_im > 1 Then
                            compteur_im = compteur_im - 1
                            ReDim Preserve horaires_im(1 To compteur_im)
                            duree_im = duree_im - t_fin_im + t_deb_im
                        Else
                            horaires_im(compteur_im) = horaires_im(compteur_im) & "      Aucun d�faut d�tect� automatiquement" & Chr(10)
                        End If
                    Else
                        horaires_im(compteur_im) = horaires_im(compteur_im) & Chr(10)
                    End If
                Else
                    analyse_im = False
                End If
            End If
        End If
        ' Analyse des defauts
        ' ----------------
        If analyse_im And valeur = 0 Then
            ' Default DPT
            For c = 0 To 2
                If Cells(i, C_Def_DPT + c) = 0 And PP_Def_DPT(c) = "" Then
                    For j = 1 To prof_recherche
                        jour_suivant = DateSerial(Cells(i + j, C_Annee), Cells(i + j, C_Mois), Cells(i + j, C_Jour))
                        heure_suivante = jour_suivant * 24 * 3600 + Cells(i + j, C_Heure) * 3600 + Cells(i + j, C_Minute) * 60 + Cells(i + j, C_Seconde)
                        If (heure_suivante - heure > Delai_defaut_DPT Or heure_suivante = 0 Or j = prof_recherche) And Cells(i + j, C_Def_DPT + c) = 0 Then
                            compteur_dpt = compteur_dpt + 1
                            PP_Def_DPT(c) = Cells(1, C_Def_DPT + c) & " (" & Cells(i, C_Heure) & ":" & Format(Cells(i, C_Minute), "00") & " pdt " & CInt(heure_suivante - heure) & "s)"
                            Exit For
                        ElseIf Cells(i + j, C_Def_DPT + c) = 1 Then
                            Exit For
                        End If
                    Next j
                End If
            Next c
            ' Default capteur laser
            nb_capteur = 0
            For pp = 1 To NbPP
                nb_capteur_pp = 0
                For c = 1 To 3
                    def_capteur = False
                    For d = 1 To 3
                        dcpp = (pp - 1) * 3 + (c - 1) * 3 + d
                        C_dcpp = C_SLG + (pp - 1) * 12 + (c - 1) * 4 + d
                        If Cells(i, C_dcpp) = 1 Then
                            If PP_Def_DILH(dcpp) = "" Then
                                PP_Def_DILH(dcpp) = Cells(1, C_PT_Confirme + pp) & " " & Cells(1, C_dcpp) & " (" & Cells(i, C_Heure) & ":" & Format(Cells(i, C_Minute), "00") & ")"
                                If Not def_capteur Then compteur_dilh = compteur_dilh + 1
                            End If
                            def_capteur = True
                        End If
                    Next d
                    If def_capteur Then
                        nb_capteur_pp = nb_capteur_pp + 1
                    End If
                Next c
                nb_capteur = WorksheetFunction.Max(nb_capteur, nb_capteur_pp)
            Next pp
            If nb_capteur = 1 Then
                compteur_train_dilh1 = compteur_train_dilh1 + nouveau_train
            ElseIf nb_capteur > 1 Then
                compteur_train_dilh2 = compteur_train_dilh2 + nouveau_train
            End If
            For c = 0 To NbALIM - 1
                If (Cells(i, C_ALIM + c) = 0 And PP_Def_ALIM(c) = "") Then
                    PP_Def_ALIM(c) = Cells(1, C_ALIM + c) & " (" & Cells(i, C_Heure) & ":" & Format(Cells(i, C_Minute), "00") & ")"
                    compteur_alim = compteur_alim + 1
                End If
            Next c
            ' Defauts Relais
            For c = 0 To NbREL - 1
                If (Cells(i, C_REL + c) = 1 And PP_Def_REL(c) = "") Then
                    PP_Def_REL(c) = Cells(1, C_REL + c) & " (" & Cells(i, C_Heure) & ":" & Format(Cells(i, C_Minute), "00") & ")"
                    compteur_rel = compteur_rel + 1
                End If
            Next c
            ' Compteur trains avec IM
            compteur_train_im = compteur_train_im + nouveau_train
        End If
        'Memorise l'etat precedent
        last_im = valeur
Nexti:
        'Incr�mente la variable d'une unit� afin de tester la cellule suivante
    Next i
    
    'Corrige la date d'analyse en cas de fin > 00h00
    Debug.Print "  Dernier train � " & last_horaire_train & " le " & jour
    jour = jour_precedent + 1
    Debug.Print "================================"
    Debug.Print " Fin de analyseDefautDPTetDILH"
    Debug.Print "================================"
    
End Function

Sub genereRapport(compteur_train As Integer, compteur_train_im As Integer, compteur_train_dilh1 As Integer, compteur_train_dilh2 As Integer, horaires_train As String, horaires_default_train As String, _
        compteur_rapi As Integer, horaires_rapi As String, _
        compteur_im As Integer, duree_im As Long, horaires_im() As String, compteur_dilh As Integer, compteur_dpt As Integer, compteur_alim As Integer, compteur_rel As Integer, _
        Optional afficheRapport As Boolean = True)
        
    Rapport = _
           " Rapport d'analyse du fichier : " & ActiveWorkbook.FullName & Chr(10) & Chr(10) & _
           "Trains : " & Chr(10) & _
           "  " & compteur_train & " passages de train " & horaires_train & Chr(10) & _
           "  dont " & compteur_train_im & " pendant Info Maintenance." & Chr(10)
    If horaires_default_train <> "" Then Rapport = Rapport & _
           "  ATTENTION : Passages de trains sans PPFV ou COPP : " & horaires_default_train & Chr(10)
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
        If compteur_dilh > 0 Then Rapport = Rapport & "    " & compteur_dilh & " d�fault laser." & Chr(10)
        If compteur_dpt > 0 Then Rapport = Rapport & "    " & compteur_dpt & " d�fault DPT." & Chr(10)
        If compteur_alim > 0 Then Rapport = Rapport & "    " & compteur_alim & " d�fault d'alimentation." & Chr(10)
        If compteur_rel > 0 Then Rapport = Rapport & "    " & compteur_rel & " d�fault de relai." & Chr(10)
        Rapport = Rapport & "  Horaires : " & Chr(10)
        For Each Line In horaires_im
            Rapport = Rapport & "  @ " & Line
        Next Line
    End If
    
    If InStr(Application.OperatingSystem, "Macintosh") > 0 Then
        Fichier_sortie = "Analyse.txt"
    Else
        Fichier_sortie = ActiveWorkbook.FullName & "_Analyse.txt"
    'End If
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
    End If
    'MsgBox "Le rapport d'analyse a �t� g�n�r� dans le fichier " & Fichier_sortie '& " dans Mes Documents"
    If afficheRapport Then
    If InStr(Application.OperatingSystem, "Macintosh") > 0 Then
       ' Dim strScript As String
       ' strScript = strScript & "tell application ""TextEdit"" " & vbCr
       ' strScript = strScript & "open """ & Fichier_sortie & """ " & vbCr
       ' strScript = strScript & "activate" & vbCr
       ' strScript = strScript & "end tell"
       '
       ' MacScript strScript
       ' AppActivate "TextEdit", True
        MsgBox Rapport
    Else
    'commande de lancement de IE
        Shell "C:\WINDOWS\explorer.exe " & Fichier_sortie
    End If
    End If
End Sub

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
    iColStop = 256
    If Cells(iLineNumber, iColStop) = "" Then iColStop = Range("iv1").End(xlToLeft).Column

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

Public Function getWordLastCol(ByVal sExpression As String, ByVal iLineNumber As Integer, Optional ByVal bPartial As Boolean = False, Optional ByVal bSelectResult As Boolean = False, Optional vsSheetName As Variant) As Integer
'   sExpression       mot(s) ou partie de mot � chercher
'   iLineNumber      numero de la ligne dans laquelle chercher
'   bPartial              choix sur le mot comlet ou partie du  mot
'   bSelectResult     s�lectionner la cellule de  la premi�re occurence trouv�e
'   vsSheetName     nom de la feuille dans laquelle cherche, celle active par d�faut
'   RETURN              num�ro de la colonne de la derni�re occurence trouv�e
    Dim iColStop    As Integer
    Dim i           As Integer

    'selection  feuille
    If Not IsMissing(vsSheetName) Then Sheets(vsSheetName).Select

    'derni�re cellule
    iColStart = 256
    If Cells(iLineNumber, iColStart) = "" Then iColStart = Range("iv1").End(xlToLeft).Column
    
    If bPartial Then
        For i = iColStart To 1 Step -1
            If Cells(iLineNumber, i) Like "*" & sExpression & "*" Then
                getWordLastCol = i
                If bSelectResult Then Cells(iLineNumber, i).Select
                Exit For
            End If
        Next i
    Else
        For i = iColStart To 1 Step -1
            If Cells(iLineNumber, i) = sExpression Then
                getWordLastCol = i
                If bSelectResult Then Cells(iLineNumber, i).Select
                Exit For
            End If
        Next i
    End If
End Function

Public Function getWordLine(ByVal sExpression As String, ByVal iColNumber As Integer, Optional ByVal bPartial As Boolean = False, Optional ByVal bCheck As Boolean = False, Optional ByVal bSelectResult As Boolean = False, Optional vsSheetName As Variant) As Integer
'   sExpression     mot(s) ou partie de mot � chercher
'   iColNumber      numero de la colonne dans laquelle chercher
'   bPartial        choix sur le mot comlet ou partie du  mot
'   bCheck          verifie que le resultat est > 0
'   bSelectResult   s�lectionner la cellule de  la premi�re occurence trouv�e
'   vsSheetName     nom de la feuille dans laquelle cherche, celle active par d�faut
'   RETURN          num�ro de la colonne de la premi�re occurence trouv�e
    Dim iLineStop   As Long
    Dim i           As Integer

    'selection  feuille
    If Not IsMissing(vsSheetName) Then Sheets(vsSheetName).Select

    'derni�re cellule
    iLineStop = 65536
    If Cells(iLineStop, iColNumber) = "" Then iLineStop = Range("A65536").End(xlUp).Row
    
    If bPartial Then
        For i = 1 To iLineStop
            If Cells(i, iColNumber) Like "*" & sExpression & "*" Then
                getWordLine = i
                If bSelectResult Then Cells(i, iColNumber).Select
                Exit For
            End If
        Next i
    Else
        For i = 1 To iLineStop
            If Cells(i, iColNumber) = sExpression Then
                getWordLine = i
                If bSelectResult Then Cells(i, iColNumber).Select
                Exit For
            End If
        Next i
    End If
    If getWordLine = 0 Then MsgBox ("La ligne comprenant le texte " & sExpression & " n'a pas pu etre trouv�e dans la colonne num�ro " & iColNumber)
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

Public Function SecondsToDate(Secondes As Long, Optional Fmt As String = "hh:mm:ss") As Date

    Dim nb_heure As Long
    Dim nb_minute As Integer
    Dim nb_seconde As Integer

    nb_heure = CLng(Secondes / 3600 - 0.5)
    nb_minute = CInt((Secondes - nb_heure * 3600) / 60 - 0.5)
    nb_seconde = CInt(Secondes - nb_heure * 3600 - nb_minute * 60)

    SecondsToDate = Format(TimeValue(nb_heure & ":" & nb_minute & ":" & nb_seconde), Fmt)

End Function

Public Function IsInArray(FindValue As String, arrSearch As _
   Variant, Optional ByVal bPartial As Boolean = False) As Boolean

    On Error GoTo LocalError
    If Not IsArray(arrSearch) Then Exit Function
    
        'IsInArray = InStr(1, vbNullChar & Join(arrSearch, _
     vbNullChar) & vbNullChar, vbNullChar & FindValue & _
     vbNullChar) > 0
    'Else
    For Each patron In arrSearch
        If bPartial Then patron = "*" & patron & "*"
        If FindValue Like patron Then
            IsInArray = True
            Exit Function
        End If
    Next patron
    'End If
Exit Function
LocalError:
    'Justin (just in case)
End Function


