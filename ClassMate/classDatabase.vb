Imports System.Data.SQLite


Public Class clsssDatabase
    ''' <summary>
    ''' SQLiteクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class classSQLite

        ''' <summary>
        ''' データベース接続文字列を設定します。
        ''' </summary>
        ''' <remarks></remarks>
        Dim _ConnectString As String = ""
        ''' <summary>
        ''' データ取得、更新時に使用するパラメータを設定します。
        ''' </summary>
        ''' <remarks></remarks>
        Dim _SqlParameters As New ArrayList
        ''' <summary>
        ''' 変数-コネクション
        ''' </summary>
        ''' <remarks></remarks>
        Dim _SqlCn As SQLite.SQLiteConnection
        Public Property COMMAND_TIMEOUT As Integer = 120

        Public Sub New()
            Me._ConnectString = "Data Source=" + Application.StartupPath + "\..\..\classmate.sqlite3"
        End Sub

        ''' <summary>
        ''' 指定したSQLを実行し、その結果をDataTableオブジェクトで取得します。
        ''' </summary>
        ''' <param name="Query"></param>
        ''' <returns>DataTable</returns>
        ''' <remarks></remarks>
        Public Function GetTableObject(ByVal Query As String) As DataTable

            Dim tbl = New DataTable()

            ' Connect Connection
            Using con As New SQLiteConnection(Me._ConnectString)
                ' Open Comand
                Using cmd As New SQLite.SQLiteCommand(Query, con)
                    Dim da = New SQLiteDataAdapter(cmd)
                    ' Set Parameter
                    For Each p In Me._SqlParameters
                        da.SelectCommand.Parameters.Add(p)
                    Next
                    ' Set Data in DataTable
                    da.Fill(tbl)
                End Using
            End Using

            Return tbl

        End Function

        ''' <summary>
        ''' パラメータをクリアします。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ParameterClear()
            Me._SqlParameters = New ArrayList
        End Sub


        ''' <summary>
        ''' パラメータを追加します。
        ''' </summary>
        ''' <param name="MyName">パラメータ名を設定してください。（例：@Mynameなど）</param>
        ''' <param name="MyValue">値を設定してください。</param>
        ''' <remarks></remarks>
        Public Sub AddParametar(ByVal MyName As String, ByVal MyValue As Object)

            Dim p As New SQLite.SQLiteParameter(MyName, MyValue)

            ' パラメータを追加します。
            Me._SqlParameters.Add(p)

        End Sub

        ''' <summary>
        ''' パラメータを追加します。
        ''' </summary>
        ''' <param name="MyName">パラメータ名を設定してください。（例：@Mynameなど）</param>
        ''' <param name="MyValue">値を設定してください。</param>
        ''' <remarks></remarks>
        Public Sub AddParametarDate(ByVal MyName As String, ByVal MyValue As Date)
            Me._SqlParameters.Add(New SQLite.SQLiteParameter(MyName, MyValue))
        End Sub

        Public Sub Dispose()
            ' パラメーターをクリアする
            Me._SqlParameters.Clear()

        End Sub

        Public Function SqlExecute(ByVal Query As String) As Integer
            Dim result As Integer

            Using con As New SQLite.SQLiteConnection(Me._ConnectString)
                ' Don't Forget!
                con.Open()
                ' Begin Transaction
                Using trans As System.Data.Common.DbTransaction = con.BeginTransaction()
                    Using cmd As New SQLite.SQLiteCommand(Query, con)
                        For Each p In Me._SqlParameters
                            cmd.Parameters.Add(p)
                        Next
                        cmd.CommandTimeout = COMMAND_TIMEOUT
                        ' Execute Query Update/Insert
                        result = cmd.ExecuteNonQuery
                    End Using
                    ' Commit
                    trans.Commit()
                End Using
            End Using

            Return result

        End Function

        Public Function CheckDBConnectTest_NoThrow() As Boolean

            Using cn As New SQLite.SQLiteConnection(Me._ConnectString)
                'Dim a As Integer = cn.GetActiveConnectionsCount()
                Try
                    ' 接続テスト
                    'cn.Open()
                    'cn.Close()

                Catch Ex As Exception
                    ' エラー情報を取得
                    Return False
                End Try
            End Using

            Return True

        End Function

    End Class
End Class

