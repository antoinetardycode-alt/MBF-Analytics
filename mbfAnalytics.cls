' **************************************************
' MBF Analytics
' Banque de méthodes de Données financières en ligne
' Simple utilisation
' Robuste
' Personnalisable API Key
' ***************************************************


' References :  Outils > References
' 1- Microsoft HTML Object Library
' 2- Microsoft WinHTTP Services, version 5.1

' Classe
' 1- Webtool
' 2- Dictionnary


' Constantes des la classe
Private Const keylen As Integer = 50

' Propriétés de la classe
Private apikey As String    ' code perso pour se connecter à Rapid API
Private h As Object         ' h objet pour les requêtes HTTP
Public status As String     ' état de la connection avec Rapid API :
                            ' ok : tout va bien
                            ' API Key invalide : problème de clé API
                            ' Compte épuisé : plus de requêtes possible
                            ' Paramètres de requête incohérents : Invalid request
                            ' Endpoint Inconnue : Endpoint * does not exist
                            ' Erreur non reconnue
                            

' ******************
' Méthodes Internes
' ******************
' Constructeur de la classe
Private Sub Class_Initialize()
    Set h = New WebTool
    status = "ok"
End Sub

Public Sub initKey(key As String)
    
    apikey = Trim(key)
    
    If Len(apikey) <> keylen Then
        status = "API Key invalide"
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        End
    Else
        status = "ok"
    End If
        
    Set h = New WebTool

End Sub

' ------------------------------------------------------------
' Fonction de recherche ou création d'une feuille Excel nommée
' Figée
' Paramètre
' - sheetname : nom de la feuille Excel recherchée
' Retour : objet feuille Excel
Private Function getOrCreateSheet(ByVal sheetname As String) As Worksheet
    Dim w As Worksheet
    
    Set w = Nothing
    
    For Each w In Worksheets
        If w.Name = sheetname Then Exit For
    Next w
    
    If w Is Nothing Then
        Set w = Worksheets.Add
        w.Name = sheetname
    End If

    Set getOrCreateSheet = w
End Function


' ***********************************************
' ***********************************************
' Collection de Méthodes financières de la classe
' ***********************************************
' ***********************************************

' ****************************************************************
' Batterie de test des erreurs possibles du résultat des requêtes
' ****************************************************************
' Paramètre
' - resultat : le résultat de la requête RapidAPI
' Retour : MsgBox puis End si une erreur est détectée
Private Sub testErreur(Optional ByVal resultat As String = "")
    
    If resultat Like "*" & "not subscribed * API" & "*" Then status = "Souscription à l'API nécessaire"

    If resultat Like "*" & "Endpoint * does not exist" & "*" Then status = "Endpoint inconnue"
    
    If resultat Like "*" & "Invalid request" & "*" Then status = "Paramètres de requête incohérents"

    If resultat Like "*" & "Wrong * parameter" & "*" Then status = "Paramètres de requête incohérents"
        
    If resultat Like "*" & "data"":null" & "*" Then status = "Pas de résultats. Vérifier les paramètres"
    
    If resultat Like "*" & "have exceeded * quota" & "*" Then status = "Quota de requêtes dépassé"
    
    ' ajouter ses propres détctions d'erreur ici et signaler l'auteur et l'objectif
    ' By XXXXX Objet YYYYY
    ' If resultat Like "*" & "??? * ???" & "*" Then status = "?????"
        

    If status <> "ok" Then
        MsgBox "Erreur : " & status & vbCrLf & vbCrLf & resultat, vbCritical, "MBF Analytics"
        End
    End If

End Sub


' *************************************
' Composantes financières d'une société
' *************************************
' Auteur : Yves Bertini
' Lien page API : https://rapidapi.com/vibapidev/api/bloomberg-real-time/playground/apiendpoint_ac53fec3-fbdc-4a73-ba91-1217728e55c9
' Params avec valeur d'exemple :
'  sheetname = "bloomberg_financials" nom de la feuille crée
'  symbol ou id = "ORA:FP"  Orange sur bourse de Paris France Paris, ticker Bloomberg
'  currencyname = "EUR"     Euros
' Retour :
'  résultats formatés sur feuille excel sheetname
'
' Exemple invocation Bloomberg Financial
'    Call m.bloomberg_financials(sheetname:="Orange Soc", symbol:="ORA:FP", currencyname:="EUR")
'
Public Sub bloomberg_financials(Optional ByVal sheetname = "bloomberg_financials", Optional ByVal symbol = "ORA:FP", Optional ByVal currencyname = "EUR")
    
    ' Etat de la connexion Rapide API
    Call testErreur
    
    ' Parties de la Requête à l'API
    hote = "bloomberg-real-time.p.rapidapi.com"
    epqs = "/stocks/financials?query=SYMBOL&currency=CURRENCYNAME" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    symbol = WorksheetFunction.EncodeURL(symbol)
    currencyname = WorksheetFunction.EncodeURL(currencyname)
    
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "SYMBOL", symbol)
    epqs = Replace(epqs, "CURRENCYNAME", currencyname)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = o("data")(3)("securityID")
    w.[B1] = currencyname
    
    
    ' Puis écrire toutes les informationss
    
    ' Income Statement, Balance Sheet, CashFlow
    n1 = 0
    For k = 1 To o("data").Count        ' tous les for avec extrémité Count
        
        Set co = w.[A2].Rows(2 + n1)  ' cellule origine du tableau
        co.Value = o("data")(k)("name"): n = n + 1
        n2 = 0
        For t = 1 To o("data")(k)("timeBasedSheets").Count
            Set ca = co.Rows(2 + n2)   ' cellule origine des data
            
            ca.Value = o("data")(k)("timeBasedSheets")(t)("name")
            p = 1
            For i = 1 To o("data")(k)("timeBasedSheets")(t)("columnHeadings").Count
                ca.Columns(2).Columns(i) = o("data")(k)("timeBasedSheets")(t)("columnHeadings")(i)
            Next i
            
            Set cd = ca.Rows(2)
            For n = 1 To o("data")(k)("timeBasedSheets")(t)("chartData").Count
                cd.Rows(n) = o("data")(k)("timeBasedSheets")(t)("chartData")(n)("name")
                For i = 1 To o("data")(k)("timeBasedSheets")(t)("chartData")(n)("values").Count
                    cd.Columns(2).Rows(n).Columns(i) = o("data")(k)("timeBasedSheets")(t)("chartData")(n)("values")(i)
                    cd.Columns(2).Rows(n).Columns(i).Style = "Currency"
                Next i
                p = p + 1
            Next n
            n2 = n2 + p + 1
        Next t
        n1 = n1 + n2 + 1
    Next k
    
    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
End Sub



' ***************************************
' Historique récent des Cotation de titre
' ***************************************
' Auteur : Yves Bertini
' Lien vers la page API : https://rapidapi.com/StoxApi/api/real-time-quotes1/playground/endpoint_c974b05a-9606-45c7-a541-5295f24857c2
' Params avec valeur d'exemple :
'  sheetname = "real_time_quotes1" nom de la feuille crée
'  interval = "4hour" ou 1min 5min 15min 30min 1hour 4hour
'  symbol = "ORA.PA"  Orange sur bourse de Paris France Paris, ticker Yahoo Finance
'  from  = "2025-01-01"     date de départ ISO en texte, FR UTC+1
'  until  = "2025-11-01"     date de fin ISO en texte, FR UTC+1
' Retour :
'  résultats formatés sur feuille excel sheetname
'
' Exemple invocation  Time Quotes1
' Call m.real_time_quotes1(sheetname:="Orange Sto", interval:="4hour", symbol:="ORA.PA", fromdt:="2025-06-01", untildt:="2025-11-01")
'
Public Sub real_time_quotes1(Optional ByVal sheetname = "real_time_quotes1", Optional ByVal interval = "4hour", Optional ByVal symbol = "ORA.PA", Optional ByVal fromdt = "2025-06-01", Optional ByVal untildt = "2025-11-01")
        
    ' Etat de la connexion Rapide API
    Call testErreur
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "real-time-quotes1.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/api/v2/historical/stock?interval=INTERVAL&symbol=SYMBOL&from=FROMDT&until=UNTILDT"
    
    
    ' Intégration des PARAMS dans la Query String
    ' Replace sensible MAJ MIN
    epqs = Replace(epqs, "INTERVAL", interval)
    epqs = Replace(epqs, "SYMBOL", symbol)
    epqs = Replace(epqs, "FROMDT", fromdt)
    epqs = Replace(epqs, "UNTILDT", untildt)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = symbol
    w.[B1] = interval
    w.[C1] = fromdt
    w.[d1] = untildt

    n = o.Count     ' nombre de lignes
    p = o(1).Count  ' nombre de colonnes

    w.[A3].Resize(1, p) = o(1).Keys()
    For k = 1 To n
        w.[A4].Rows(k).Resize(1, p) = o(k).Items()
    Next k
    
    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
End Sub



' *************************************
' Cotation journalière historique d'une période donnée Yahoo Finance
' *************************************
' Auteur : Théo, Mathieu, Noé
' Lien vers la page API : https://rapidapi.com/calvinlo/api/yh-finance-complete
' Params avec valeur d'exemple :
'   sheetname = "yhf_historical" nom de la feuille crée
'   ticker    = "AAPL"           Ticker boursier (ex: AAPL, TSLA, ORA.PA)
'   sdate     = "2024-01-10"     Date de début (format YYYY-MM-DD) à partir du 1 janvier 1970
'   edate     = "2024-02-16"     Date de fin (format YYYY-MM-DD) jusqu'à j-7
' Retour
'   Feuille Excel avec historique de cotation nettoyé et formaté
Public Sub yhf_historical(Optional ByVal sheetname = "yhf_historical", Optional ByVal ticker = "AAPL", Optional ByVal sdate = "2024-01-10", Optional ByVal edate = "2024-02-16")
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "yh-finance-complete.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/yhfhistorical?ticker=TICKER&sdate=SDATE&edate=EDATE"
    
    ' Encodage URL des PARAMS (Optionnel, ticker peut contenir des caractères spéciaux)
    ' ticker = WorksheetFunction.EncodeURL(ticker)
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "TICKER", ticker)
    epqs = Replace(epqs, "SDATE", sdate)
    epqs = Replace(epqs, "EDATE", edate)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    ' Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur API : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Vérification qu'il y a des données
    If o.Count = 0 Then
        MsgBox "Aucune donnée.", vbExclamation
        Exit Sub
    End If

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les informations (Titre et Période)
    With w.Range("A1")
        .Value = "Historique : " & ticker
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    w.Range("A2").Value = "Période : " & sdate & " au " & edate
    
    
    ' Variables de dimension
    Dim n As Long, p As Long
    n = o.Count
    p = o(1).Count
    
    ' Récupération des clés (en-têtes)
    Dim keys As Variant
    keys = o(1).keys()
    
    ' Identification de la colonne Date pour le nettoyage
    Dim colDate As Integer
    colDate = -1
    
    Dim i As Integer
    For i = 0 To p - 1
        Dim kName As String
        kName = LCase(keys(i))
        ' On repère la colonne qui contient "date"
        If kName = "date" Or kName = "datetime" Then colDate = i
    Next i
    
    ' Écriture des En-têtes ligne 4
    w.Cells(4, 1).Resize(1, p).Value = keys
    
    ' Écriture des Données avec nettoyage de la date
    Dim startRow As Long
    startRow = 5
    
    Dim r As Long
    Dim rowItems As Variant
    Dim cleanDate As String
    
    For r = 1 To n
        ' Récupère la ligne de données sous forme de tableau (Array)
        rowItems = o(r).Items()
        
        If colDate >= 0 Then
            cleanDate = CStr(rowItems(colDate))
            
           
            cleanDate = Replace(cleanDate, "T", " ")   ' Remplace le T par un espace
            cleanDate = Replace(cleanDate, "000Z", "") ' Supprime le 000Z
          
            
            rowItems(colDate) = cleanDate
        End If
        
        ' Écriture de la ligne modifiée dans Excel
        w.Cells(startRow + r - 1, 1).Resize(1, p).Value = rowItems
    Next r


    
    Dim tblRange As Range
    Set tblRange = w.Range(w.Cells(4, 1), w.Cells(startRow + n - 1, p))
    
    ' Style En-têtes
    With w.Range(w.Cells(4, 1), w.Cells(4, p))
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .RowHeight = 25
    End With
    
    ' Bordures du tableau
    tblRange.Borders.LineStyle = xlContinuous
    tblRange.Borders.Color = RGB(200, 200, 200)
    
    ' Formatage des Nombres et des Dates
    Dim c As Integer
    For c = 1 To p
        ' Index colonne Excel = Index Array + 1
        
        ' 1. Formatage DATE
        If (c - 1) = colDate Then
            w.Columns(c).NumberFormat = "dd/mm/yyyy hh:mm"
            w.Columns(c).HorizontalAlignment = xlRight
        
        ' 2. Formatage VOLUME (Nombre entier)
        ElseIf InStr(1, LCase(keys(c - 1)), "volume") > 0 Then
             w.Columns(c).NumberFormat = "#,##0"
             
        ' 3. Formatage PRIX (Décimales)
        Else
             w.Columns(c).NumberFormat = "#,##0.00"
        End If
    Next c
    
    ' Ajustement automatique des largeurs
    w.UsedRange.Columns.AutoFit

    ' Figer les volets
    w.Activate
    ActiveWindow.FreezePanes = False
    w.Rows("5:5").Select
    ActiveWindow.FreezePanes = True
    w.Range("A1").Select
    
End Sub


' *************************************
' Méthode générique à personnaliser
' *************************************
' Auteur : ** les auteurs **
' Lien vers la page API : https://rapidapi.com/...
' Params avec valeur d'exemple :
'  sheetname = " ** nom de la méthode ** " nom de la feuille crée
'  params1 = "exemple1"
'  params2 = "exemple2"
' Retour :
'  résultats formatés sur feuille excel sheetname
'
Public Sub maMethode(Optional ByVal sheetname = " ** maMethode ** ", Optional ByVal monparam1 = "exemple1", Optional ByVal monparam2 = "exemple2")
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = " ** mettre son nom d'hôte ici ** "
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/ ** son endpoint ? param1=PARAM1&param2=PARAM2" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    symbol = WorksheetFunction.EncodeURL(symbol)
    currencyname = WorksheetFunction.EncodeURL(currencyname)
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "PARAM1", monparam1)
    epqs = Replace(epqs, "PARAM2", monparam2)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = param1
    w.[B1] = param2
    
    
    ' Puis écire toutes les informations sur la feuille w
    ' ...
    
End Sub



