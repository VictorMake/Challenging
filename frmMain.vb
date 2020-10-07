Imports System.Windows.Forms
Imports BaseFormKT

' имя класса состоит из имени DLL (varName) и имени главной формы визуального наследования (frmMain)
' Dim ClassName As String = varName & ".frmMain"
' имя класса DLL и тег визуально наследуемой формы должны совпадать (в данном случае Me.Tag=Gasgenerator)
' assembly name  и root namespace - FormOne на странице свойств проекта также совпадает
' 4.	В формк frmMain.vb меняю TAG на новое ???
' 5.	В паке G:\DiskD\ПрограммыVBNET\RegistrationНаследование\bin\Ресурсы\МодулиСбораКТ копирую ???.mdb , ???.xml  и создаю папку ??? После компиляции туда же копирую ???.DLL
' Для теста в форме TestBaseInherit меняю настройки app.config для     Private AssemblyName As String = System.Configuration.ConfigurationManager.AppSettings("CalculationAssemblyFilename")
' Для         Dim ClassName As String = System.Configuration.ConfigurationManager.AppSettings("DiagramClassName")
' Для         КлассГрафик.Manager.ПутьКБазеПараметров = "G:\DiskD\ПрограммыVBNET\RegistrationНаследование\bin\Ресурсы\МодулиСбораКТ\Challenging.mdb"
' 6.	Тестирую

Public Class frmMain
    Public WithEvents myClassCalculation As ClassCalculation

    'Private mClassCalculation As IClassCalculation
    Public Overrides Property ClassCalculation() As IClassCalculation
        Get
            Return myClassCalculation
        End Get
        Set(ByVal value As IClassCalculation)
            myClassCalculation = CType(value, ClassCalculation)
        End Set
    End Property

    Private mOwnCatalogue As String '= IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) & "\" & Me.Tag
    Public Overrides Property OwnCatalogue() As String
        Get
            Return mOwnCatalogue
        End Get
        Set(ByVal value As String)
            mOwnCatalogue = value
            If Not System.IO.Directory.Exists(mOwnCatalogue) Then
                MessageBox.Show(Me, "Рабочий каталог " & OwnCatalogue & " для модуля расчета отсутствует. Необходимо его скопировать.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                'System.Windows.Forms.Application.Exit()
                'System.Environment.Exit(0)
            End If
        End Set
    End Property

    ''' <summary>
    ''' Видима DLL или нет, т.е. имеет ли другие окна или она только вычисляет
    ''' </summary>
    ''' <remarks></remarks>
    Private mIsDLLVisible As Boolean = False
    Public Overrides ReadOnly Property IsDLLVisible() As Boolean
        Get
            Return mIsDLLVisible
        End Get
    End Property

    'Private varProjectManager As ProjectManager
    'Public Overrides Property Manager() As ProjectManager
    '    Get
    '        Return varProjectManager
    '    End Get
    '    Set(ByVal value As ProjectManager)
    '        varProjectManager = value
    '    End Set
    'End Property


    Public Overrides Sub GetValueTuningParameters()
        If myClassCalculation IsNot Nothing Then myClassCalculation.ПолучитьЗначенияНастроечныхПараметров()
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        'WindowManagerPanel1.AutoDetectMdiChildWindows = False
        'MyBase.WindowManagerPanel1.CloseAllWindows

        ' выполняется после MyBase_Load
        mIsDLLVisible = True 'False 'это свойство определяет поведение расчетного модуля
        If mIsDLLVisible Then
            'Dim m_ChildFormNumber As Integer
            'For I As Integer = 1 To 1
            '    Dim ChildForm As New Explorer 'System.Windows.Forms.Form
            '    ' Make it a child of this MDI form before showing it.
            '    ChildForm.MdiParent = Me

            '    m_ChildFormNumber += 1
            '    ChildForm.Text = "Window " & m_ChildFormNumber
            '    ChildForm.Show()
            '    'MyBase.WindowManagerPanel1.SetActiveWindow(CType(ChildForm, System.Windows.Forms.Form))

            '    'MyBase.WindowManagerPanel1.AddWindow(ChildForm)
            'Next

            'MyBase.Manager.ПутьКБазеПараметров = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) & "\Ресурсы\МодулиРасчета\" & Me.Tag & ".mdb"
            ' каталог нужен только для видимых форм, где производится сохранение результатов работы
            OwnCatalogue = MyBase.Manager.PathCatalog & "\" & Me.Tag.ToString
            PathResourses = OwnCatalogue
        End If

        MyBase.FrmBaseLoad()

        ' сделать невозможным закрытие формы
        Me.Manager.LoadConfiguration() 'из XML
        gMainFomMdiParent = Me
        ' идет вслед за Me.Manager.СчитатьНастройки()
        ' в varТемпературныеПоля инициализируется ПараметрыПоляНакопленные, которая используется в конструкторе New ClassCalculation(
        myClassCalculation = New ClassCalculation(Me.Manager)
        Me.ClassCalculation = myClassCalculation

        'MyBase.varfrmGraf.tsTextBoxСбор.Visible = False
        ' если какое-то сообщение будет до загрузки сеток то перерисовка их будет вызывать исключения, поэтому
        ' myClassCalculation.ПолучитьЗначенияНастроечныхПараметров() идет после Me.Manager.FillCombo()
        myClassCalculation.ПолучитьЗначенияНастроечныхПараметров()

        'If Not varIsDLLVisible Then Me.Hide()

        '' установить фокус вначале на mdi child
        ''Me.MdiChildren(0).BringToFront()
        '' эквивалентный метод: 
        ''WindowManagerPanel1.SetActiveWindow(MdiChildren.Count - 1)
        ''рекомендуемый (хотя и не необходимый) использовать WindowManager methods

        'Application.DoEvents()
        'Thread.Sleep(10000)

        'Application.DoEvents()
        'Me.WindowManagerPanel1.SetActiveWindow(CType(varТемпературныеПоля, System.Windows.Forms.Form))
        'Application.DoEvents()
        TimerUpdate.Enabled = True

        'WindowManagerPanel1.ShowCloseButton = True
        'WindowManagerPanel1.ShowLayoutButtons = True
    End Sub

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles Me.FormClosed
        gMainFomMdiParent = Nothing
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        'If Not CType(Me.MdiParent, BaseFormKT.frmBaseKT).ЗакрытьОкно Then
        If Not gMainFomMdiParent.IsWindowClosed Then
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub myClassCalculation_DataError(ByVal sender As IClassCalculation, ByVal e As IClassCalculation.DataErrorEventArgs) Handles myClassCalculation.DataError
        'если произошла ошибка в классе или ошибка была специально сгенерирована то вывести сообщение
        'sender.Manager.ПутьККаталогу - получить дополнительную информацию
        Dim TITLE As String = "Подсчет для модуля " & Text
        MessageBox.Show($"Ошибка в подсчете ClassCalculation.dll:{ Environment.NewLine}{e.Message}{Environment.NewLine}{e.Description}",
                        TITLE, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub TimerUpdate_Tick(ByVal sender As System.Object, ByVal e As EventArgs) Handles TimerUpdate.Tick
        TimerUpdate.Enabled = False
        MyBase.WindowManagerPanel1.SetActiveWindow(0)
        MyBase.WindowManagerPanel1.SetActiveWindow(CType(MyBase.varFormMeasurementKT, Form))
    End Sub
End Class