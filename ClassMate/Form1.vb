Imports System.IO
Imports GrapeCity.ActiveReports

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Viewer1.LoadDocument(Application.StartupPath + "\..\..\classmate_changedataset.rdlx")
        Dim rptPath As New FileInfo("..\..\classmate_changedataset.rdlx")
        Dim definition As New PageReport(rptPath)
        AddHandler definition.Document.LocateDataSource, AddressOf OnLocateDataSource
        Viewer1.LoadDocument(definition.Document)

    End Sub
    Private Sub OnLocateDataSource(ByVal sender As Object, ByVal args As LocateDataSourceEventArgs)
        Dim db = New clsssDatabase.classSQLite
        Dim tbl = New DataTable
        tbl = db.GetTableObject("Select * from classmate")
        args.Data = tbl
    End Sub
End Class
