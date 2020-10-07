Imports System.IO
Imports System.Windows.Forms

Module GlobalVariable
    Public PathResourses As String
    Public Const PROVIDER_JET As String = "Provider=Microsoft.Jet.OLEDB.4.0;"

    Public gMainFomMdiParent As frmMain
    Public clAir As MathematicalLibrary.Air

    Public Function BuildCnnStr(ByVal provider As String, ByVal dataBase As String) As String
        'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database Locking Mode=1;Data Source="D:\ПрограммыVBNET\RUD\RUD.NET\bin\Ресурсы\Channels.mdb";Jet OLEDB:Engine Type=5;Provider="Microsoft.Jet.OLEDB.4.0";Jet OLEDB:System database=;Jet OLEDB:SFP=False;persist security info=False;Extended Properties=;Mode=Share Deny None;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Global Bulk Transactions=1
        Return String.Format("{0}Data Source={1};", provider, dataBase)
    End Function

    ''' <summary>
    ''' True - файла нет
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    Private Function FileNotExists(ByVal path As String) As Boolean
        'FileExists = CBool(Dir(FileName) = vbNullString) 
        Return Not File.Exists(path)
    End Function

    ''' <summary>
    ''' Проверка существования файла
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    Public Function CheckExistsFile(ByVal path As String) As Boolean
        If FileNotExists(path) Then
            MessageBox.Show($"В каталоге нет файла <{path}> !", "Провека существования файла", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            Return True
        End If
    End Function
End Module
