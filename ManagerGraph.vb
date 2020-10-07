Imports System.IO
Imports System.Windows.Forms
Imports MathematicalLibrary
Imports MathematicalLibrary.Spline3

' экземпляры классов организуются в коллекцию в менеджере
' менеджер создается в Public Sub New() ClassCalculation
' в конструктор менеджера передается строковый массив с именами файлов (массив состоит из строковых констант имен) и путь к БД
' в цикле создаются экземпляры классов и добавляются в коллекцию
' метод ("имя класса",X,Y) as double который инкапсулирует нахождение в коллекции экземпляра класса и запрос у него (X,Y) as double

Enum GraphType
    ГрафикТемпКоэфК1
    ГрафикВлажностиКоэфК2
End Enum

Friend Class ManagerGraphKof
    Implements IEnumerable

    ' внутренняя коллекция для управления формами
    Private mCollectionGraph2D As Dictionary(Of String, Graph2D)
    Private Property PathCatalog As String

    Public Sub New(ByVal inPathCatalog As String)
        Me.PathCatalog = inPathCatalog
        mCollectionGraph2D = New Dictionary(Of String, Graph2D)
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return mCollectionGraph2D.Count
        End Get
    End Property

    Public ReadOnly Property CollectionGraph2D() As Dictionary(Of String, Graph2D)
        Get
            Return mCollectionGraph2D
        End Get
    End Property

    Public ReadOnly Property Item(ByVal indexKey As String) As Graph2D
        Get
            Return mCollectionGraph2D.Item(indexKey)
        End Get
    End Property

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return mCollectionGraph2D.GetEnumerator
    End Function

    Public Sub Remove(ByVal indexKey As String)
        ' удаление по номеру или имени или объекту?
        ' если целый тип то по плавающему индексу, а если строковый то по ключу
        mCollectionGraph2D.Remove(indexKey)
    End Sub

    Public Sub Clear()
        mCollectionGraph2D.Clear()
    End Sub

    Protected Overrides Sub Finalize()
        mCollectionGraph2D = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub ДобавитьГрафики(ByVal arrГрафик() As String, ByVal ТипГрафика As GraphType, ByVal Режим As ClassCalculation.RegimeEngine)
        For Each sName As String In arrГрафик
            Add(sName, PathCatalog, ТипГрафика, Режим)
        Next
    End Sub

    Public Sub Add(ByVal name As String, ByVal inPathCatalog As String, ByVal ТипГрафика As GraphType, ByVal режим As ClassCalculation.RegimeEngine)
        If Not ПроверкаИмени(name) Then Exit Sub

        mCollectionGraph2D.Add(name, New Graph2D(name, inPathCatalog, ТипГрафика, режим))
    End Sub

    Private Function ПроверкаИмени(ByVal Name As String) As Boolean
        If mCollectionGraph2D.ContainsKey(Name) Then
            MessageBox.Show("График с именем " & Name & " уже существует!", "Ошибка добавления графика в коллекцию", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If
        Return True
    End Function

    Public Function СчитатьИнициализировать() As Boolean
        Dim success As Boolean = False
        Dim fileName As String = Nothing

        Try
            For Each tempGraph2D As Graph2D In mCollectionGraph2D.Values
                fileName = tempGraph2D.FileName
                tempGraph2D.Open()
            Next

            success = True
        Catch e As FileNotFoundException
            ' Console.WriteLine("The file {0} does not exist", fileName)
            MessageBox.Show("Файл {" & fileName & "} отсутствует", "Функция СчитатьИнициализировать", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch e As ColdCallFileFormatException
            ' Console.WriteLine("The file {0} appears to have been corrupted", fileName)
            ' Console.WriteLine("Details of problem are: {0}", e.Message)
            ' If e.InnerException IsNot Nothing Then
            '     Console.WriteLine("Inner exception was: {0}", e.InnerException.Message)
            ' End If
            MessageBox.Show("Файл {" & fileName & "} имеет неправильный формат" & vbLf & e.Message, "Функция СчитатьИнициализировать", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch e As Exception
            ' Console.WriteLine("Exception occurred:" & vbLf & e.Message)
            MessageBox.Show("Файл {" & fileName & "} вызвал исключение:" & vbLf & e.Message, "Функция СчитатьИнициализировать", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try
        Return success
    End Function

    Public ReadOnly Property КоэфПересчета(ByVal indexKey As String) As Double
        Get
            Return mCollectionGraph2D.Item(indexKey).КоэфПересчета
        End Get
    End Property

    ''' <summary>
    ''' В зависимости от ТипГрафика графика выбирается метод аппроксимации,
    ''' а в зависимости от режима выбираются те таблицы, которые соответствуют этому режиму.
    ''' </summary>
    ''' <param name="Тбокса"></param>
    ''' <param name="N1измеренные"></param>
    ''' <param name="ВлажностьОтн"></param>
    ''' <param name="Режим"></param>
    Public Sub ВычислитьВсеКоэфПересчета(ByVal Тбокса As Double, ByVal N1измеренные As Double, ByVal ВлажностьОтн As Double, ByVal Режим As ClassCalculation.RegimeEngine)
        For Each tempGraph2D As Graph2D In mCollectionGraph2D.Values
            Select Case tempGraph2D.ТипГрафика
                Case GraphType.ГрафикТемпКоэфК1
                    If tempGraph2D.Режим = Режим Then
                        tempGraph2D.ВычислитьКоэфПересчета(N1измеренные, Тбокса)
                    End If
                Case GraphType.ГрафикВлажностиКоэфК2
                    If tempGraph2D.Режим = Режим Then
                        tempGraph2D.ВычислитьКоэфПересчета(Тбокса, ВлажностьОтн)
                    End If
            End Select
        Next
    End Sub

End Class

''' <summary>
''' Класс считывания из файла таблицы
''' </summary>
''' <remarks></remarks>
Friend Class Graph2D
    Implements IDisposable

    ' Класс с методом считывания из файла таблицы, в случае ошибки выдать сообщение
    ' свойства:
    Public Property Name As String ' имя класса
    Public Property ПутьКаталога As String ' - путь с БД
    Public Property FileName As String
    Public Property ТипГрафика As GraphType
    Public Property Режим As ClassCalculation.RegimeEngine
    Public Property КоэфПересчета As Double ' метод(параметр X,Y) для выдачи интерполированного числа 

    Private fs As FileStream
    Private sr As StreamReader
    Private _ЧислоСтрок As Integer
    Private _ЧислоСтолбцов As Integer

    Private isDisposed As Boolean = False
    Private isOpen As Boolean = False

    Private КоэфПерBicubic As Bicubic
    Private Const N1Start As Integer = 75
    Private Const N1End As Integer = 105

    Public Sub New(ByVal Name As String, ByVal ПутьКаталога As String, ByVal ТипГрафика As GraphType, ByVal Режим As ClassCalculation.RegimeEngine)
        Me.Name = Name
        Me.ПутьКаталога = ПутьКаталога
        FileName = Path.Combine(ПутьКаталога, Name & ".txt")
        Me.Режим = Режим
        Me.ТипГрафика = ТипГрафика

        Select Case ТипГрафика
            Case GraphType.ГрафикТемпКоэфК1
                _ЧислоСтрок = 21
                _ЧислоСтолбцов = 19
            Case GraphType.ГрафикВлажностиКоэфК2
                _ЧислоСтрок = 19
                _ЧислоСтолбцов = 6
        End Select
    End Sub

    ''' <summary>
    ''' метод расчета внутренних коэф(,) (массив координат X(), Y(), Z(,)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Open()
        ' при создании New считывается файл как массив строк, строки разбираются как массивы слов, далее в цикле через шаг 2 производится разбор в X(),Y()
        ' для быстрого поиска применяются индексы, когда ищетсяи выходит за рамки диапазона выдает первое или последние значения массива
        ' Bicubic подходит для апроксимации влажности
        ' для К от температур нужно сделать равномерный квадрат по X - обороты, по Y -  температуры (они равномерны)
        ' считываются все массивы длиной 21 arrX(20) и arrY(20) по оборотам 
        '     Public Sub New(ByVal InputArray As Double(,), ByVal XIstart As Double, ByVal XIfinish As Double, ByVal YJstart As Double, ByVal YJfinish As Double)
        ' разобъём диапазон на N1End - N1Start  и создадим arrXNew(N1End - N1Start) и вычислим для каждого arrXNew(N1End - N1Start) новые значения в массиве ZNew(50) Spline3Interpolate   (для этого нужно предварительно вычислить таблицу методом Spline3BuildTable)
        ' тогда соединим все в массив ZNew(,)(массив коэффициентов) - где 20 это шкала температур Y, N1End - N1Star - ось оборотов X
        ' тогда можно вычислять так N1пересчитанныйКоэффициент = New Bicubic(ZNew(,), XIstart, XIfinish, YJstart, YJfinish)
        Dim delimiterTab As Char() = vbTab.ToCharArray
        Dim words() As String
        Dim I, J, L As Integer
        Dim strList As New List(Of String)
        Dim arrList() As String
        Dim line As String = Nothing

        Dim axisY() As Double
        Dim axisХ() As Double
        Dim Z(,) As Double

        If isDisposed Then
            Throw New ObjectDisposedException("Graph2D_" & _Name)
        End If
        If Not File.Exists(FileName) Then
            Throw New FileNotFoundException("Отсутствует файл " & FileName)
        End If

        fs = New FileStream(FileName, FileMode.Open)
        sr = New StreamReader(fs)

        Try
            ' считать первую строку
            Dim firstLine As String = sr.ReadLine()
            ' определить число шлейфов
            words = firstLine.Split(delimiterTab)

            ' предпоследняя цифра содержит число столбцов
            If words.Length / 2 <> _ЧислоСтолбцов OrElse Integer.Parse(words(words.Length - 2)) <> _ЧислоСтолбцов Then
                Throw New ColdCallFileFormatException("Неправильное число столбцов в файле")
            End If

            J = 0
            ' определить для какого диапазона шлейф
            'ReDim_axisY(_ЧислоСтолбцов - 1)
            Re.Dim(axisY, _ЧислоСтолбцов - 1)
            For I = 1 To words.Length - 1 Step 2
                axisY(J) = Double.Parse(words(I))
                J += 1
            Next

            ' остальные строки в цикле
            ' здесь считывание и проверка числа столбцов из первой строки
            Do While sr.Peek() >= 0
                line = sr.ReadLine()
                strList.Add(line)
            Loop
            sr.Close()
            fs.Close()
            fs = Nothing

            If strList.Count <> _ЧислоСтрок Then
                Throw New ColdCallFileFormatException("Неправильное число строк в файле")
            End If
            isOpen = True ' присвоить только после проверки

            'ReDim_Z(_ЧислоСтрок - 1, _ЧислоСтолбцов - 1)
            Re.Dim(Z, _ЧислоСтрок - 1, _ЧислоСтолбцов - 1)

            Select Case ТипГрафика
                Case GraphType.ГрафикТемпКоэфК1
                    'ReDim_axisХ(N1End - N1Start)
                    Re.Dim(axisХ, N1End - N1Start)
                    Dim arrОсьХ(_ЧислоСтрок - 1, _ЧислоСтолбцов - 1) As Double
                    Dim InputX(_ЧислоСтрок - 1), InputY(_ЧислоСтрок - 1) As Double ' для сплайн интерполяции
                    Dim InputX2(_ЧислоСтрок), InputY2(_ЧислоСтрок) As Double ' для линенйной интерполяции
                    Dim tblBuildTable(,) As Double = Nothing

                    Dim ZNew(N1End - N1Start, _ЧислоСтолбцов - 1) As Double
                    arrList = strList.AsEnumerable().Reverse().ToArray()

                    For I = 0 To _ЧислоСтрок - 1
                        words = arrList(I).Split(delimiterTab)
                        If words.Length / 2 <> _ЧислоСтолбцов Then
                            Throw New ColdCallFileFormatException("Неправильное число столбцов в файле")
                        End If

                        For J = 0 To words.Length \ 2 - 1 ' Step 2
                            arrОсьХ(I, J) = Double.Parse(words(J * 2))
                            Z(I, J) = Double.Parse(words(J * 2 + 1))
                        Next
                    Next

                    For J = 0 To _ЧислоСтолбцов - 1
                        For I = 0 To _ЧислоСтрок - 1
                            InputX(I) = arrОсьХ(I, J)
                            InputY(I) = Z(I, J)
                            InputX2(I + 1) = InputX(I)
                            InputY2(I + 1) = InputY(I)
                        Next
                        Spline3BuildTable(UBound(InputX) + 1, 2, InputX, InputY, 0, 0, tblBuildTable)

                        L = 0
                        For Обороты = N1Start To N1End
                            If Обороты < InputX(0) OrElse Обороты > InputX(_ЧислоСтрок - 1) Then
                                ' если вне диапазона то линейная интерполяция
                                ZNew(L, J) = InterpLine(InputX2, InputY2, Обороты)
                            Else ' сплайн интерполяции
                                ZNew(L, J) = Spline3Interpolate(UBound(tblBuildTable, 2) + 1, tblBuildTable, Обороты)
                            End If
                            L += 1
                        Next
                    Next

                    L = 0
                    For Обороты = N1Start To N1End
                        axisХ(L) = Обороты
                        L += 1
                    Next

                    КоэфПерBicubic = New Bicubic(ZNew, axisХ(0), axisХ(axisХ.Length - 1), axisY(0), axisY(_ЧислоСтолбцов - 1), 30)
                Case GraphType.ГрафикВлажностиКоэфК2
                    'ReDim_axisХ(_ЧислоСтрок - 1)
                    Re.Dim(axisХ, _ЧислоСтрок - 1)
                    arrList = strList.ToArray()

                    For I = 0 To _ЧислоСтрок - 1
                        words = arrList(I).Split(delimiterTab)
                        If words.Length / 2 <> _ЧислоСтолбцов Then
                            Throw New ColdCallFileFormatException("Неправильное число столбцов в файле")
                        End If

                        axisХ(I) = Double.Parse(words(0))
                        For J = 0 To words.Length \ 2 - 1 ' Step 2
                            Z(I, J) = Double.Parse(words(J * 2 + 1))
                        Next
                    Next

                    КоэфПерBicubic = New Bicubic(Z, axisХ(0), axisХ(_ЧислоСтрок - 1), axisY(0), axisY(_ЧислоСтолбцов - 1), 50)
            End Select
        Catch e As FormatException
            Throw New ColdCallFileFormatException("Неправильное число строк в файле", e)
        End Try
    End Sub

    Public Sub ВычислитьКоэфПересчета(ByVal X As Double, ByVal Y As Double)
        КоэфПересчета = КоэфПерBicubic.Calculate(X, Y)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If isDisposed Then
            Return
        End If

        isDisposed = True
        isOpen = False
        If fs IsNot Nothing Then
            sr.Close()
            fs.Close()
            fs = Nothing
        End If
    End Sub
End Class

Class ColdCallFileFormatException
    Inherits ApplicationException
    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
        MyBase.New(message, innerException)
    End Sub
End Class
