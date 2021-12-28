Sub AddDateToSubject()


Dim olFolder As MAPIFolder
Dim olSelection As Selection
Dim olItem As Object
Dim iCountMeetingItems As Integer

Dim olItemCurrentDate As Date

Set olFolder = Application.ActiveExplorer.CurrentFolder

If olFolder.DefaultItemType = olMailItem Then
    Set olSelection = Application.ActiveExplorer.Selection
    
    For Each olItem In olSelection
    
    With olItem
       If TypeOf olItem Is MailItem Then
         olItemCurrentDate = .ReceivedTime
        .Subject = Year(olItemCurrentDate) & "-" & Format(Month(olItemCurrentDate), "00") & "-" & Format(Day(olItemCurrentDate), "00") & " " & .Subject
        .Save
        Else
        'ElseIf TypeOf olItem Is MeetingItem Then --> if you want match MeetingItems
        iCountMeetingItems = iCountMeetingItems + 1
       End If
   
    End With

    Next
    
    If iCountMeetingItems > 0 Then

    MsgBox "Hinweis: Ihre Auswahl enthält: " & CStr(iCountMeetingItems) & " Obejekt(e). Diese können bei der automatischen Umbennenung des Betreffs nicht berücksichtigt werden.", vbInformation
       
    End If
    
End If
End Sub
