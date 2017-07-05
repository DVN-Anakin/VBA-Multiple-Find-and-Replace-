# VBA-Multiple-Find-and-Replace
VBA Macro for finding multiple text strings in a given Word document and replacing them with the appropriate strings based on an Excel database.

----------------
1. WHAT IS THAT FOR
----------------
So, imagine You have a website that You have created from scratch i.e. You are not using any CMS (WP, Prestashop, Drupal, ...) whatsoever. Now, You want to create multiple language versions of your website. First solution that You can think of is obviously to duplicate those .html files and manually translate/rewrite the strings in the source code. This "solution" is surely logical however very laborious indeed - it would take so much of your time, especially if your website has a lot of source files to go through. Another solution is to just copy&paste the source code of each page into the Google translator. Of course, a setback is that both Google and Bing translators suck.

Because I had the same problem, I tried to come up with solme solution. In my case, I was given an Excel database with the original strings and their translations. The database looks like this:

![alt tag](https://github.com/DVN-Anakin/VBA-Multiple-Find-and-Replace-/blob/master/excel-database.png)

So, I have created a VBA macro to deal with this kind of issue. In the given Word document with your source code, It automatically goes through the code and finds the strings located in the left collumn of the database and replace them with their translations from the next collumn.

----------------
2. BACKLOGS
----------------
Of course, there are still many backlogs to deal with - for example duplicate strings, very similar strings, etc. I hope that smarter people can use this for their own benefit.

----------------
3. VBA MACRO CODE
----------------
I have created this macro using the latest version of MS Excel (2013). So, in case You are using previous versions of MS Excel, You will need to modify the settings in "Tools" → "References". Microsoft Office 15.0 Object Library was used for this macro. 

```VB.net
Sub multiple_find_and_replace()
    Dim Wbk As Workbook: Set Wbk = ThisWorkbook
    Dim Wrd As Object
    Set Wrd = CreateObject("Word.Application")
    Dim Dict As Object
    Dim RefList As Range
    Dim RefElem As Range
    Wrd.Visible = True
    Dim WDoc As Object
    Set WDoc = Wrd.Documents.Open("C:\Users\Anh\test.docx") 'Your Word document - rename and modify the path.
    Set Dict = CreateObject("Scripting.Dictionary")
    Set RefList = Wbk.Sheets("Sheet1").Range("A1:A4") 'Range of your strings in the database - modify this.

    With Dict
        For Each RefElem In RefList
            If Not .Exists(RefElem) And Not IsEmpty(RefElem) Then
                .Add RefElem.Value, RefElem.Offset(0, 1).Value
            End If
        Next RefElem
    End With

    For Each Key In Dict
        With WDoc.Content.Find
            .Execute FindText:=Key, ReplaceWith:=Dict(Key)
        End With
    Next Key
End Sub
```
