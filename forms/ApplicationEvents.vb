Namespace My

    ' 次のイベントは MyApplication に対して利用できます:
    ' 
    ' Startup: アプリケーションが開始されたとき、スタートアップ フォームが作成される前に発生します。
    ' Shutdown: アプリケーション フォームがすべて閉じられた後に発生します。このイベントは、通常の終了以外の方法でアプリケーションが終了されたときには発生しません。
    ' UnhandledException: ハンドルされていない例外がアプリケーションで発生したときに発生するイベントです。
    ' StartupNextInstance: 単一インスタンス アプリケーションが起動され、それが既にアクティブであるときに発生します。 
    ' NetworkAvailabilityChanged: ネットワーク接続が接続されたとき、または切断されたときに発生します。
    Partial Friend Class MyApplication






        'Private Sub MyApplication_StartupNextInstance( _
        '        ByVal sender As Object, _
        '        ByVal e As Microsoft.VisualBasic.ApplicationServices. _
        '        StartupNextInstanceEventArgs) _
        '        Handles Me.StartupNextInstance
        '    '  Console.WriteLine("二重起動されました")

        '    '後で起動されたアプリケーションのコマンドライン引数を表示


        '    For Each cmd As String In e.CommandLine
        '        Form1.Act_Main(cmd)
        '    Next

        '    '先に起動しているアプリケーションをアクティブにしない
        '    e.BringToForeground = False
        'End Sub






    End Class


End Namespace

