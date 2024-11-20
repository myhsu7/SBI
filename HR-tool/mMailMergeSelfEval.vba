' vba script
Option Explicit

Const FOLDER_SAVED As String = "C:\Users\myhsu\Downloads\YERR\SelfEval\"
Const SOURCE_FILE_PATH As String = "C:\Users\myhsu\Downloads\SBI Employee Self Evaluation Form 員工績效自我評量表.xlsx"

Sub MailMergeSelfEval()
Dim MainDoc As Document, TargetDoc As Document
Dim dbPath As String
Dim sFolderPath As String
Dim recordNumber As Long, totalRecord As Long

Set MainDoc = ActiveDocument
With MainDoc.MailMerge
    
        '// if you want to specify your data, insert a WHERE clause in the SQL statement
        .OpenDataSource Name:=SOURCE_FILE_PATH, sqlstatement:="SELECT * FROM [Form1$]"
            
        totalRecord = .DataSource.RecordCount

        For recordNumber = 1 To totalRecord
        
            With .DataSource
                .ActiveRecord = recordNumber
                .FirstRecord = recordNumber
                .LastRecord = recordNumber
            End With
            
            .Destination = wdSendToNewDocument
            .Execute False
            
            Set TargetDoc = ActiveDocument

            sFolderPath = FOLDER_SAVED & .DataSource.DataFields("Manager").Value & "\"

            If Dir(sFolderPath, vbDirectory) = "" Then
            'If folder is not existed, create it
                MkDir sFolderPath
            End If
            
            TargetDoc.SaveAs2 sFolderPath & .DataSource.DataFields("Name").Value & ".docx", wdFormatDocumentDefault
            TargetDoc.ExportAsFixedFormat sFolderPath & .DataSource.DataFields("Name").Value & ".pdf", exportformat:=wdExportFormatPDF
            
            TargetDoc.Close False
            
            Set TargetDoc = Nothing
                    
        Next recordNumber

End With

Set MainDoc = Nothing
End Sub
