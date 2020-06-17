VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



Public Function CreateCalendarEvent(CalendarName As String, Subject As String, EventDate As String, Location As String)
 
    Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
    Dim fldCalendar As Outlook.MAPIFolder
    Dim Ns As Outlook.NameSpace
    Dim myItem As Outlook.AppointmentItem

    Set Ns = Application.GetNamespace("MAPI")
    Set fldCalendar = Ns.GetDefaultFolder(olFolderCalendar).Folders(CalendarName)
    
    Set myItem = fldCalendar.Items.Add(olAppointmentItem)
         With myItem
         .MeetingStatus = olEvent
         .Subject = Subject
         .Location = Location
         .AllDayEvent = True
         .Start = EventDate
         .ReminderSet = alse
         .Save
        End With
End Function


Public Function CalendarEvents(CalendarName As String) As Object

    Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
    Dim fldCalendar As Outlook.MAPIFolder
    Dim Ns As Outlook.NameSpace

    Set Ns = Application.GetNamespace("MAPI")
    Set fldCalendar = Ns.GetDefaultFolder(olFolderCalendar).Folders(CalendarName)
    
    Set CalendarEvents = fldCalendar.Items

End Function
