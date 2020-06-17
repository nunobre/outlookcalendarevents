Attribute VB_Name = "loadEventsIntoCalendar"
Option Explicit

Sub Calendar_PEARL_loadEvents()

'---
'
Const defaultCalendarName = "[MY_CALENDAR_NAME]"

Dim listOfJiraVersions As Object
Dim jiraVersion As Object

Dim myJIRA As New JiraRest
Dim Calendar As New Calendar

    With myJIRA
        .UserName = "[USERNAME]"
        .Password = "[PASSWORD]"
        .URL = "https://MY_JIRA_URL"
        If .Login = False Then Exit Sub
    End With
    
    Set listOfJiraVersions = myJIRA.JiraVersions

    myJIRA.Logout
    
    Dim existingCalendarEvents As Outlook.Items
    Set existingCalendarEvents = Calendar.CalendarEvents(defaultCalendarName)
    
    Dim filter As String
    Dim index As Integer
    Dim releaseName As String
    Dim releaseDate As String
    'Dim eventFound As Object
    'Const PropTag  As String = "https://schemas.microsoft.com/mapi/proptag/"
    
    For index = 1 To listOfJiraVersions.Count / 9
        releaseName = listOfJiraVersions("obj(" & index & ").name")
        releaseDate = listOfJiraVersions("obj(" & index & ").releaseDate")
        filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = '" & releaseName & "'"
        
        'Set eventFound = existingCalendarEvents.Find(filter)
        
        Debug.Print releaseName
        Debug.Print releaseDate
        
        If Not existingCalendarEvents.Find(filter) Is Nothing Then
            ' item already exists in calendar
            ' do nothing
            Debug.Print "item already exists in calendar: do nothing"
            
        ElseIf Len(releaseDate) > 0 Then
            ' new item: add to calendar
            Debug.Print "new item: add to calendar"
            Call Calendar.CreateCalendarEvent(defaultCalendarName, releaseName, releaseDate, "Release")
        Else
            'TODO: release date empty. Create todo!?"
            Debug.Print "TODO: release date empty. Create todo!?"
        End If
        
    Next

End Sub
