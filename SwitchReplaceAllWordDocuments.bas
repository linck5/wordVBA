Attribute VB_Name = "NewMacros"
'Replaces a string with another and vice versa, switching the two.
'The strings are defined in as arguments in the SwitchReplace function, which can be called multiple times
'This will be executed in all .doc .dot and .docx files in a folder including sub flders
'The folder is defined in the strStartPath variable
'This will replace strings even areas outside the body of the document
'like in text boxes, headers, etc. Not doing that would make it run faster

'References used:
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=76
'http://word.mvps.org/faqs/macrosvba/FindReplaceAllWithVBA.htm

Option Explicit
 
Dim scrFso As Object
Dim scrFolder As Object
Dim scrSubFolders As Object
Dim scrFile As Object
Dim scrFiles As Object
 
Sub OpenAllFilesInFolder()
     
     'starting place for trav macro
     'strStartPath is a path to start the traversal on
    Dim strStartPath As String
    strStartPath = "C:\Users\Felipe\Desktop\ReplaceTest2\"
     
     'stop the screen flickering
    Application.ScreenUpdating = False
     
     'open the files in the start folder
    OpenAllFiles strStartPath
     'search the subfolders for more files
    SearchSubFolders strStartPath
     
     'turn updating back on
    Application.ScreenUpdating = True
     
End Sub
 
 
Sub SearchSubFolders(strStartPath As String)
     
     'starts at path strStartPath and traverses its subfolders and files
     'if there are files below it calls OpenAllFiles, which opens them one by one
     'once its checked for files, it calls itself to check for subfolders.
     
    If scrFso Is Nothing Then Set scrFso = CreateObject("scripting.filesystemobject")
    Set scrFolder = scrFso.getfolder(strStartPath)
    Set scrSubFolders = scrFolder.subfolders
    For Each scrFolder In scrSubFolders
        Set scrFiles = scrFolder.Files
        If scrFiles.Count > 0 Then OpenAllFiles scrFolder.Path 'if there are files below, call openFiles to open them
        SearchSubFolders scrFolder.Path 'call ourselves to see if there are subfolders below
    Next
     
End Sub


Sub OpenAllFiles(strPath As String)
     
     ' runs through a folder oPath, opening each file in that folder,
     ' calling a macro called samp, and then closing each file in that folder
     
    Dim strName As String
    Dim wdDoc As Document
     
    If scrFso Is Nothing Then Set scrFso = CreateObject("scripting.filesystemobject")
    Set scrFolder = scrFso.getfolder(strPath)
    For Each scrFile In scrFolder.Files
        strName = scrFile.Name 'the name of this file
        Application.StatusBar = strPath & "\" & strName 'the status bar is just to let us know where we are
         'we'll open the file fName if it is a word document or template
        If Right(strName, 4) = ".doc" Or Right(strName, 4) = ".dot" Or Right(strName, 5) = ".docx" Then
            Set wdDoc = Documents.Open(FileName:=strPath & "\" & strName, _
            ReadOnly:=False, Format:=wdOpenFormatAuto)
             
             'Call the replace method that will switch the first string with the second one
             'Can be called as many times as needed
            SwitchReplace "Eng", "Embaixador", wdDoc
            SwitchReplace "Teste1", "Teste2", wdDoc
            SwitchReplace "l1", "l2", wdDoc
            SwitchReplace "L1", "L2", wdDoc
            SwitchReplace "loja 1", "loja 2", wdDoc
            SwitchReplace "L 1", "L 2", wdDoc
             
             'we close saving changes
            wdDoc.Close wdSaveChanges
        End If
    Next
     
     'return control of status bar to Word
    Application.StatusBar = False
End Sub
 
 'the method that do the Replace
Sub SwitchReplace(stringA As String, stringB As String, wdDoc As Document)
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    'Sets the properties of the Find procedure
    'More info about the properties here: https://msdn.microsoft.com/en-us/library/office/dn320652.aspx
    With Selection.Find
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .NoProofing = False
        .Forward = True
    End With
        
    'Switch the strings
    With Selection.Find
        .Text = stringA
        .Replacement.Text = "_REPLACETEMP_sadlkfj320jsdf_"
        .Execute Replace:=wdReplaceAll
    
        .Text = stringB
        .Replacement.Text = stringA
        .Execute Replace:=wdReplaceAll
    
        .Text = "_REPLACETEMP_sadlkfj320jsdf_"
        .Replacement.Text = stringB
        .Execute Replace:=wdReplaceAll
    End With
    
    'This is the part where it replaces in areas outside the body of the document
    'Comment this out if not needed, because it will be faster
    Dim myStoryRange As Range
    
    For Each myStoryRange In ActiveDocument.StoryRanges
    
        With myStoryRange.Find
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .NoProofing = False
            .Forward = True
        End With
    
       If myStoryRange.StoryType <> wdMainTextStory Then
            With myStoryRange.Find
            
                .Text = stringA
                .Replacement.Text = "_REPLACETEMP_sadlkfj320jsdf_"
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
                
                .Text = stringB
                .Replacement.Text = stringA
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
                
                .Text = "_REPLACETEMP_sadlkfj320jsdf_"
                .Replacement.Text = stringB
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
                
            End With
            Do While Not (myStoryRange.NextStoryRange Is Nothing)
               Set myStoryRange = myStoryRange.NextStoryRange
                With myStoryRange.Find
                
                    .Text = stringA
                    .Replacement.Text = "_REPLACETEMP_sadlkfj320jsdf_"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                    
                    .Text = stringB
                    .Replacement.Text = stringA
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                    
                    .Text = "_REPLACETEMP_sadlkfj320jsdf_"
                    .Replacement.Text = stringB
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Loop
        End If
    Next myStoryRange
    
End Sub



