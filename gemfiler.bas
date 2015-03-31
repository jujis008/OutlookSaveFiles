Attribute VB_Name = "Module1"
' hhs@sfi.dk 5/6-2014
' modul til outlook: Gemmer alle vedh�ftede filer af de markede emails i �n mappe.
Sub GemFiler()

    ' Declarations
        'Outlook elementer (til navigation mm)
        Dim myItems, myItem, myAttachments, myAttachment As Object
        Dim myOlApp As New Outlook.Application
        Dim myOlExp As Outlook.Explorer
        Dim myOlSel As Outlook.Selection
        
        ' Elementer til browsedialog
        Dim oApp As Object
        Dim RootPath As String
        Set oApp = CreateObject("Shell.Application")
        
        ' �vrige elementer
        Dim fuldsti As String
        Dim lokation
        
    
    'Dialog der browser
    Set lokation = oApp.BrowseForFolder(0, "Hvor skal filerne gemmes? (Bem�rk: filer overskrives uden advarsel) ", 512)
    If lokation Is Nothing Then Exit Sub
    'Definer fuld sti
    fuldsti = lokation.Self.Path & "\"
    
    ' Tjek om mappen eksisterer (det burde ikke kunne ske)
    If Dir(fuldsti, vbDirectory) = "" Then
             MsgBox "Mappen findes ikke, pr�v igen."
    End If
    
    ' hvis noget g�r i ged, s� afslut bare
    On Error Resume Next
    
    'Nu g�r vi henover de markerede emails
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    
    ' for alle emails
    For Each myItem In myOlSel
    
        ' find den vedh�ftede fil
        Set myAttachments = myItem.Attachments
        
             ' hvis der er en vedh�ftet fil s�..
             If myAttachments.Count > 0 Then
                    
                 'gem all filer
                  For i = 1 To myAttachments.Count
            
                     'i destinationen specificeret under fuldsti (og navnet som den vedh�ftede fil har)
                      myAttachments(i).SaveAsFile fuldsti & myAttachments(i).DisplayName
                
                     ' s�t emailen som l�st
                     myItem.UnRead = False
              
                     ' og g� videre til den n�ste vedh�ftede fil
                    Next i
            ' og s� slut
            End If
            
    'videre til n�ste email
    Next
    
    'free variables
    Set myItems = Nothing
    Set myItem = Nothing
    Set myAttachments = Nothing
    Set myAttachment = Nothing
    Set myOlApp = Nothing
    Set myOlExp = Nothing
    Set myOlSel = Nothing
    
End Sub













