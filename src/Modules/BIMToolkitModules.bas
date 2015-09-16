Attribute VB_Name = "BIMToolkitModules"
'''''''''''''''''''''''''''''
' BIM Toolkit API Sample Code
'
' BIMToolkitModules
' =================
' These are the main methods that are called from the buttons
'

Option Explicit

Public g_objIE As Object

''''''''''''''''''''''''''''''''''''''''''''''''''
' Method - LaunchLODInWebBrowser
' ==============================
'
' This method takes a classification and a banding
' to display Level of Detail (LOD) guidance.
' It creates a URL that can be launched in a new
' webbrowser window or embedded inline within an
' existing webpages.
'
' Test values:
' classification: Ss_25_12_60_60
' banding = 3
' These will show the LOD guidance for banding 3 for the item 'Panel cubicle systems'
Public Sub LaunchLODInWebBrowser()
    Dim Url As String
    Dim classification As String
    Dim banding As String
    On Error GoTo handle_error
    
    ' Get the values from the worksheet
    classification = Trim(Cells.Item(2, "B").Value)
    banding = Trim(Cells.Item(3, "B").Value)
    
    ' Build the URL
    Url = "https://toolkit.thenbs.com/definitions/" & classification & "/?type=lod&detailLevel=" & banding
    
    ' Launch the URL in a web browser
    If g_objIE Is Nothing Then Set g_objIE = CreateObject("Internetexplorer.Application")
    g_objIE.Visible = True
    g_objIE.Navigate Url
    
    Exit Sub
    
handle_error:
    ' An unexpected error has occurred - fail gracefully without a crash
    ClearCells
    Set g_objIE = Nothing
    Cells(8, 1).Value = "An unexpected error as occurred."
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' Method - LaunchLODInWebBrowser
' ==============================
'
' This method takes a classification and a banding
' to display Level of Information (LOI) properties.
' It creates a URL that can be launched in a new
' webbrowser window or embedded inline within an
' existing webpages.
'
' Test values:
' classification: Pr_60_60_13_04
' banding = 4
' These will show the LOI properties for banding 4 for the item 'Air cooled liquid chillers'
Public Sub LaunchLOIInWebBrowser()
    Dim Url As String
    Dim classification As String
    Dim banding As String
    On Error GoTo handle_error
    
    ' Get the values from the worksheet
    classification = Trim(Cells.Item(2, "B").Value)
    banding = Trim(Cells.Item(3, "B").Value)
    
    ' Build the URL
    Url = "https://toolkit.thenbs.com/definitions/" & classification & "/?type=loi&detailLevel=" & banding
    
    ' Launch the URL in a web browser
    If g_objIE Is Nothing Then Set g_objIE = CreateObject("Internetexplorer.Application")
    g_objIE.Visible = True
    g_objIE.Navigate Url
    
    Exit Sub
    
handle_error:
    ' An unexpected error has occurred - fail gracefully without a crash
    ClearCells
    Set g_objIE = Nothing
    Cells(8, 1).Value = "An unexpected error has occurred."
End Sub

''''''''''''''''''''
' Method - ReturnLOI
' ==================
'
' This method takes a classification and a banding
' to return Level of Information (LOI) properties.
' It uses a webclient to return these properties from
' a classification and a banding code.
'
' Test values:
' classification: Pr_60_60_13_04
' banding = 4
' These will return the LOI guidance for banding 4 for the item 'Air cooled liquid chillers'
Public Sub ReturnLOI()
    ' Create and intialise the webclient
    Dim Client As New WebClient
    Dim clientID As String
    Dim clientSecret As String
    
    GetAPIKey clientID, clientSecret
    If clientID = "" Or clientSecret = "" Then
        MsgBox "Please note that the clientID and clientSecret must be added to cells O2 and O3 to enable the API.", vbExclamation
        Exit Sub
    End If
    
    ' Setup Oauth2 authentication handler
    Dim Auth As New BIMToolkitAuthenticator
    Auth.Setup clientID, clientSecret
    Set Client.Authenticator = Auth

    ' Get the values from the worksheet
    Dim classification As String
    Dim banding As String
    classification = Trim(Cells.Item(2, "B").Value)
    banding = Trim(Cells.Item(3, "B").Value)

    ' Create specific request
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Request.Resource = "https://toolkit-api.thenbs.com/definitions/loi/" & classification & "/" & banding
    ' Get response from request
    Set Response = Client.Execute(Request)


    If Response.StatusCode = 200 Then
        ' Successful call
        Dim i As Integer
        
        ClearCells
        Cells(7, 1).Value = "Name"
        Cells(7, 1).Font.Bold = True
        Cells(7, 2).Value = "Description"
        Cells(7, 2).Font.Bold = True
        
        On Error GoTo handle_error
        
        For i = 1 To Response.data("Data").Count
            Dim objItem As Object
            Dim j As Integer
            
            Set objItem = Response.data("Data")(i)
            
            ' Note - Comment in the lines if the 'name' value is required
            'Dim name As String
            Dim camelCaseName As String
            Dim description As String
            
            'name = objItem("Name")
            camelCaseName = objItem("CamelCaseName")
            description = objItem("Definition")
            
            'Cells(7 + i, 1).Value = name
            Cells(7 + i, 1).Value = camelCaseName
            Cells(7 + i, 2).Value = description
        Next i
    Else
        ' Unsuccessful call - inform the user why
        ClearCells
        Cells(8, 1).Value = "Error: " & Response.StatusCode & " " & Response.Content & "."
    End If
    
    Exit Sub
    
handle_error:
    ' An unexpected error has occurred - fail gracefully without a crash
    ClearCells
    If Not Response Is Nothing And Not Response.data Is Nothing And Response.data("data") = Empty Then
        Cells(8, 1).Value = "No data for LOI " + banding
    Else
        Cells(8, 1).Value = "An unexpected error has occurred."
    End If
End Sub

''''''''''''''''''''''''''''''''
' Method - ReturnClassifications
' ==============================
'
' This method takes a classification and returns
' all child classifications.
' It uses a webclient to return these items from
' a parent classification code.
'
' Test values:
' classification: Pr_60_60_13_04
' banding = 4
' These will return the LOI guidance for banding 4 for the item 'Air cooled liquid chillers'
Public Sub ReturnClassifications()
    ' Create and intialise the webclient
    Dim Client As New WebClient
    Dim clientID As String
    Dim clientSecret As String
    
    GetAPIKey clientID, clientSecret
    If clientID = "" Or clientSecret = "" Then
        MsgBox "Please note that the clientID and clientSecret must be added to cells O2 and O3 to enable the API.", vbExclamation
        Exit Sub
    End If
    
    ' Setup Oauth2 authentication handler
    Dim Auth As New BIMToolkitAuthenticator
    Auth.Setup clientID, clientSecret
    Set Client.Authenticator = Auth

    ' Get the values from the worksheet
    Dim classification As String
    classification = Trim(Cells.Item(2, "B").Value)
    
    ' The logic below works if you want sibling classifications as opposed to child classifications
    'If Len(classification) > 2 Then
    '    classification = Trim(Left(classification, InStrRev(classification, "_") - 1))
    'End If

    ' Create specific request
    Dim Request As New WebRequest
    Dim Response As WebResponse
    Request.Resource = "https://toolkit-api.thenbs.com/definitions/uniclass2015/" & classification & "/1"
    ' Get response from request
    Set Response = Client.Execute(Request)

    If Response.StatusCode = 200 Then
        ' Successful call
        Dim i As Integer
        
        ' Clear range of cells
        Range("A7:B300").Clear
        Cells(7, 1).Value = "Name"
        Cells(7, 1).Font.Bold = True
        Cells(7, 2).Value = "Description"
        Cells(7, 2).Font.Bold = True
        
        On Error GoTo handle_error
        
        Cells(8, 1).Value = Response.data("Notation")
        Cells(8, 2).Value = Response.data("Title")
        
        For i = 1 To Response.data("Children").Count
            Dim objItem As Object
            Dim j As Integer
            
            Set objItem = Response.data("Children")(i)
            
            Dim notation As String
            Dim title As String
            
            notation = objItem("Notation")
            title = objItem("Title")
            
            Cells(8 + i, 1).Value = notation
            Cells(8 + i, 2).Value = title
        Next i
    Else
        ' Unsuccessful call - inform the user why
        ClearCells
        Cells(8, 1).Value = "Error: " & Response.StatusCode & " " & Response.Content & "."
    End If
    
    Exit Sub
    
handle_error:
    ' An unexpected error has occurred - fail gracefully without a crash
    ClearCells
    Cells(8, 1).Value = "An unexpected error has occurred."
End Sub

Public Sub ClearCells()
    ' Clear range of cells
    Range("A7:B300").Clear
End Sub

Public Sub GetAPIKey(ByRef clientID As String, ByRef clientSecret As String)
    clientID = Cells.Item(2, "O").Value
    clientSecret = Cells.Item(3, "O").Value
End Sub
