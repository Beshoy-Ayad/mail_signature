On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
If IsEmpty(objUser) Then
MsgBox "No connection with LDAP information.", vbCritical, "Error": WScript.Quit (1)
End If

strName = objUser.FullName
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strPhone = objUser.telephoneNumber
StrMobile = objUser.mobile
strFax = objUser.facsimileTelephoneNumber
StrAdd = objUser.streetAddress


'location of logo and social media images. Mind to use links to a network share, to which users will have access.

StrLogo0 = "file://mail2.spinneys-egypt.com/Signature/image001.jpg"
strLogo = "file://mail2.spinneys-egypt.com/Signature/image002.png"
strLogo2 ="file://mail2.spinneys-egypt.com/Signature/winnersignature1.png"
strSoc1 = "file://mail2.spinneys-egypt.com/Signature/image004.png"
strSoc2 = "file://mail2.spinneys-egypt.com/Signature/image005.jpg"
strSoc3 = "file://mail2.spinneys-egypt.com/Signature/image006.png"
strSoc4 = "file://mail2.spinneys-egypt.com/Signature/image007.png"
strSoc5= "file://mail2.spinneys-egypt.com/Signature/image008.jpg"
strSoc6 = "file://mail2.spinneys-egypt.com/Signature/image009.jpg"
strWeb = "https://Spinneys-egypt.com"
strFBlink = "https://www.facebook.com/SpinneysEgypt"
strYTlink = "https://www.youtube.com/channel/UCx86OHcfQXXFuLrCSIayY3w/videos"
strTWlink = "https://twitter.com/SpinneysEgypt"
strLNlink = "https://www.linkedin.com/company/spinneysegypt?trk=biz-companies-cym"
strInsLink = "https://www.instagram.com/spinneysegypt/"

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

' Beginning of signature block

Set objRange = objDoc.Range()

'Create table for signature content

objDoc.Tables.Add objRange,10,9
Set objTable = objDoc.Tables(1)
objdoc.Paragraphs.SpaceAfter = 0

'Merge selected table cells

objTable.Cell(1,1).Merge objTable.Cell(1,5)
objTable.Cell(2,1).Merge objTable.Cell(2,5)
objTable.Cell(3,1).Merge objTable.Cell(3,5)
objTable.Cell(4,1).Merge objTable.Cell(4,6)
objTable.Cell(5,1).Merge objTable.Cell(5,5)
objTable.Cell(6,1).Merge objTable.Cell(6,9)
objTable.Cell(9,1).Merge objTable.Cell(9,9)
objtable.cell(10,1).Merge objtable.cell(10,5)



'Add user's data

objTable.Cell(1,1).Range.Font.Name = "Calibri"
objTable.Cell(1,1).Range.Font.Size = "10"
objTable.Cell(1,1).Range.Font.Bold = True
objTable.Cell(1,1).Range.Text = strName

objTable.cell(2,1).Range.Font.Name = "Calibri"
objTable.cell(2,1).Range.Font.Size = "10"
objTable.cell(2,1).Range.Font.Bold = True
objtable.cell(2,1).Range.Font.color = rgb(168, 168, 168)
objTable.Cell(2,1).Range.Text = strTitle & "| Spinneys Egypt"

objTable.cell(4,1).Range.font.name = "Calibri"
objTable.cell(4,1).Range.font.size = "9"
objTable.cell(4,1).Range.font.Bold = True
objTable.cell(4,1).Range.Text = "Phone: " & strPhone & " |Mobile: " & StrMobile & " |Fax: " & strFax

objTable.cell(5,1).Range.font.name = "Calibri"
objTable.cell(5,1).Range.font.size = "9"
objTable.cell(5,1).Range.font.Bold = True
objtable.cell(5,1).Range.Text = StrAdd

'add Line
objTable.Cell(6,1).Range.InlineShapes.AddPicture(StrLogo0)

'add Company Logos 
objtable.Cell(7,1).width = 150
objTable.Cell(7,1).Range.InlineShapes.AddPicture(strLogo)
'objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(2), strWeb
objtable.Cell(7,2).width = 100
objtable.cell(7,2).height = 50
objTable.Cell(7,2).Range.InlineShapes.AddPicture(strLogo2)

'add Social Media Logos
objtable.Cell(7,4).width = 26
objtable.Cell(7,4).height = 30
'objtable.cell(7,4).borderwidth = 0
'objtable.cell(7,4).leftPadding = 0
objtable.cell(7,4).RightPadding = 0
objtable.cell(7,4).topPadding = 3
objtable.cell(7,4).margins = 0
objTable.Cell(7,4).Range.InlineShapes.AddPicture(strSoc1)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(4), strWeb

objtable.Cell(7,5).width = 20
objtable.Cell(7,5).height = 30
'objtable.cell(7,5).borderwidth = 0
objtable.cell(7,5).leftPadding = 0
objtable.cell(7,5).RightPadding = 0
objTable.Cell(7,5).Range.InlineShapes.AddPicture(strSoc2)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(5), strFBlink

objtable.Cell(7,6).width = 20
objtable.Cell(7,6).height = 30
'objtable.cell(7,5).borderwidth = 0
objtable.cell(7,6).leftPadding = 0
objtable.cell(7,6).RightPadding = 0
objTable.Cell(7,6).Range.InlineShapes.AddPicture(strSoc3)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(6), strTWlink

objtable.Cell(7,7).width = 20
objtable.Cell(7,7).height = 30
'objtable.cell(7,5).borderwidth = 0
objtable.cell(7,7).leftPadding = 0
objtable.cell(7,7).RightPadding = 0
objTable.Cell(7,7).Range.InlineShapes.AddPicture(strSoc4)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(7), strInsLink

objtable.Cell(7,8).width = 20
objtable.Cell(7,8).height = 30
'objtable.cell(7,5).borderwidth = 0
objtable.cell(7,8).leftPadding = 0
objtable.cell(7,8).RightPadding = 0
objTable.Cell(7,8).Range.InlineShapes.AddPicture(strSoc5)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(8), strYTlink

objtable.Cell(7,9).width = 30
objtable.Cell(7,9).height = 30
objtable.cell(7,9).borderwidth = 0
objtable.cell(7,9).leftPadding = 0
objtable.cell(7,9).RightPadding = 0
objTable.Cell(7,9).Range.InlineShapes.AddPicture(strSoc6)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(9), strLNlink

'add the last two lines

objTable.cell(9,1).Range.Font.Name = "Calibri"
objTable.cell(9,1).Range.Font.Size = "9"
objTable.cell(9,1).Range.Font.Bold = True
objtable.cell(9,1).Range.Font.color = rgb(168, 168, 168)
objTable.Cell(9,1).Range.Text = "This message is subject to the terms at the link:"

objTable.cell(10,1).Range.Font.Name = "Calibri"
objTable.cell(10,1).Range.Font.Size = "9"
objTable.cell(10,1).Range.Font.Bold = False
objtable.cell(10,1).Range.Font.color = rgb(33, 82, 243)
objTable.Cell(10,1).Range.Text = "www.spinneys-egypt.com/Disclaimer"


' End of signature block

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Spinneys-Egypt_Signature", objSelection

objSignatureObject.NewMessageSignature = "Spinneys-Egypt_Signature"

objSignatureObject.ReplyMessageSignature = "Spinneys-Egypt_Signature"

objDoc.Saved = True

objWord.Quit