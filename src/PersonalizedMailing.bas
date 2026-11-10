Attribute VB_Name = "PersonalizedMailing"
'===============================================================================
' Module      : PersonalizedMailing
' Description : Sends a personalized HTML email campaign through Outlook,
'               using a list of contacts stored in the active Excel sheet.
'
' Expected sheet layout (row 1 = headers, data starts on row 2):
'   Column A : ID         -> unique contact ID (appended to the button link)
'   Column B : CIVILITY  -> title (e.g. "Mr", "Mrs")
'   Column C : NAME        -> last name
'   Column D : FIRST_NAME     -> first name
'   Column E : EMAIL      -> recipient email address
'
' Requirements:
'   - Microsoft Excel and Microsoft Outlook installed on the same machine
'   - The sender account configured in Outlook
'
' Usage:
'   1. Fill in the constants in the CONFIGURATION section below
'   2. Edit the email body in the BuildEmailBody function
'   3. Open the contacts sheet and run SendPersonalizedEmails
'
' Tip: set SEND_DIRECTLY to False first to review the emails before sending.
'===============================================================================

Option Explicit

'------------------------------- CONFIGURATION ---------------------------------
' Outlook account used to send the emails (must exist in Outlook)
Private Const SENDER_ADDRESS As String = "your.address@example.com"

' Name of the Outlook signature file (found in %APPDATA%\Microsoft\Signatures\)
Private Const SIGNATURE_FILE As String = "MySignature.htm"

' Email subject
Private Const EMAIL_SUBJECT As String = "Your email subject"

' Base URL of the call-to-action button (the contact ID is appended to it)
Private Const BUTTON_BASE_URL As String = "https://www.example.com/form?id="

' Text displayed on the call-to-action button
Private Const BUTTON_TEXT As String = "Click here"

' True = send immediately / False = open each email for review (.Display)
Private Const SEND_DIRECTLY As Boolean = False
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' Main procedure: loops through the contact list and sends one email per row
'-------------------------------------------------------------------------------
Sub SendPersonalizedEmails()

    Dim OutApp As Object            ' Outlook application
    Dim OutMail As Object           ' Email item being created
    Dim SenderAccount As Object     ' Outlook account used to send
    Dim LastRow As Long             ' Last row containing data
    Dim i As Long                   ' Row counter
    Dim contactId As String
    Dim title As String
    Dim lastName As String
    Dim firstName As String
    Dim email As String
    Dim htmlBody As String          ' Full HTML body of the email
    Dim signatureHTML As String     ' HTML of the Outlook signature
    Dim signatureFolder As String   ' Outlook signatures folder
    Dim buttonLink As String        ' Personalized link for the button

    ' Start (or connect to) Outlook
    Set OutApp = CreateObject("Outlook.Application")

    ' Look for the Outlook account matching the sender address.
    ' If no account matches, the loop ends with SenderAccount = Nothing
    ' and Outlook's default account will be used.
    For Each SenderAccount In OutApp.Session.Accounts
        If LCase(SenderAccount.SmtpAddress) = LCase(SENDER_ADDRESS) Then
            Exit For
        End If
    Next SenderAccount

    ' Path to the Outlook signatures folder of the current Windows user
    signatureFolder = Environ("APPDATA") & "\Microsoft\Signatures\"

    ' Load the signature (with image paths fixed), or a fallback message
    If Dir(signatureFolder & SIGNATURE_FILE) <> "" Then
        signatureHTML = GetSignatureHTML(signatureFolder & SIGNATURE_FILE, signatureFolder)
    Else
        signatureHTML = "<p>Signature not found</p>"
    End If

    ' Find the last filled row, based on the EMAIL column (E)
    LastRow = Cells(Rows.Count, "E").End(xlUp).Row

    ' Loop through each contact (row 1 contains the headers)
    For i = 2 To LastRow

        ' Read the contact data from the current row
        contactId = Cells(i, "A").Value
        title = Cells(i, "B").Value
        lastName = Cells(i, "C").Value
        firstName = Cells(i, "D").Value
        email = Cells(i, "E").Value

        ' Build the personalized link for this contact
        buttonLink = BUTTON_BASE_URL & contactId

        ' Skip rows without an email address
        If email <> "" Then

            ' Create a new email item (0 = olMailItem)
            Set OutMail = OutApp.CreateItem(0)

            ' Build the personalized body and append the signature
            htmlBody = BuildEmailBody(title, lastName, firstName, buttonLink)
            htmlBody = htmlBody & signatureHTML

            With OutMail
                ' Use the chosen sender account if it was found
                If Not SenderAccount Is Nothing Then
                    Set .SendUsingAccount = SenderAccount
                End If
                .To = email
                .Subject = EMAIL_SUBJECT
                .HTMLBody = htmlBody

                ' Send directly or open the email for review
                If SEND_DIRECTLY Then
                    .Send
                Else
                    .Display
                End If
            End With

            Set OutMail = Nothing
        End If

    Next i

    ' Release the Outlook object
    Set OutApp = Nothing

    MsgBox "Campaign finished.", vbInformation

End Sub


'-------------------------------------------------------------------------------
' Builds the personalized HTML body of the email.
' Edit the text below to write your own message.
'-------------------------------------------------------------------------------
Private Function BuildEmailBody(ByVal title As String, ByVal lastName As String, _
                                ByVal firstName As String, ByVal buttonLink As String) As String

    Dim body As String

    ' Greeting, personalized with the contact's title and last name
    body = "<p>Hello " & title & " " & lastName & ",</p>"

    ' Main paragraphs of the message
    body = body & "<p>[Introduce yourself / your company here.]</p>"
    body = body & "<p>[Explain your offer and its benefits here.]</p>"

    ' Call-to-action button: an HTML link styled as a button
    ' (inline styles are used because many email clients ignore CSS classes)
    body = body & "<p>" & _
                  "<a href='" & buttonLink & "' style='background-color:#2f60c3; color:white; " & _
                  "padding:12px 24px; text-decoration:none; font-weight:bold; display:inline-block;'>" & _
                  BUTTON_TEXT & _
                  "</a></p>"

    ' Optional bulleted list
    body = body & "<ul>" & _
                  "<li>[Item 1]</li>" & _
                  "<li>[Item 2]</li>" & _
                  "<li>[Item 3]</li>" & _
                  "</ul>"

    ' Closing
    body = body & "<p>[Closing sentence.]</p>"
    body = body & "<p>Best regards,</p>"

    BuildEmailBody = body

End Function


'-------------------------------------------------------------------------------
' Reads an Outlook signature HTML file and converts relative image paths
' into absolute paths, so the images are found when the email is built.
'
' Parameters:
'   htmlFile        : full path of the signature .htm file
'   signatureFolder : Outlook signatures folder (used as base path)
'-------------------------------------------------------------------------------
Function GetSignatureHTML(htmlFile As String, signatureFolder As String) As String

    Dim FSO As Object
    Dim Stream As Object
    Dim content As String
    Dim RegEx As Object
    Dim Matches As Object
    Dim Match As Object
    Dim imagePath As String

    ' Read the whole signature file
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Stream = FSO.OpenTextFile(htmlFile, 1)   ' 1 = ForReading
    content = Stream.ReadAll
    Stream.Close

    ' Regular expression matching every src="...image..." attribute
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Global = True
    RegEx.IgnoreCase = True
    RegEx.Pattern = "src=""([^""]*\.(png|jpg|jpeg|gif|bmp)([^""]*)?)"""

    Set Matches = RegEx.Execute(content)

    For Each Match In Matches
        imagePath = Match.SubMatches(0)

        ' Only process relative paths (an absolute path contains ":")
        If InStr(imagePath, ":") = 0 Then
            ' Prefix with the signatures folder, unless the path goes up a level
            If Left(imagePath, 3) <> "../" Then
                imagePath = signatureFolder & imagePath
            End If
            ' Replace the original src attribute with the corrected one
            content = Replace(content, Match.Value, "src=""" & imagePath & """")
        End If
    Next

    GetSignatureHTML = content

End Function
