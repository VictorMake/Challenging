Imports System.Data.OleDb
Imports System.Windows.Forms
Imports BaseFormKT
Imports MathematicalLibrary
Imports MathematicalLibrary.Spline3

Public Class ClassCalculation
    Implements IClassCalculation

    ' событие для выдачи ошибки в вызывающую программу
    Public Event DataError(ByVal sender As IClassCalculation, ByVal e As IClassCalculation.DataErrorEventArgs) Implements IClassCalculation.DataError

    ' FireEventArgs: пользовательское событие наследуется от EventArgs.
    Public Class DataErrorEventArgs
        Inherits EventArgs
        Public ReadOnly Property Description() As String
        Public ReadOnly Property Message() As String

        Public Sub New(ByVal message As String, ByVal description As String)
            Me.Message = message
            Me.Description = description
        End Sub
    End Class

    Public Property Manager() As ProjectManager Implements IClassCalculation.Manager
    Private mManagerGraph As ManagerGraphKof
    Private counterUpdate As Integer

    ' Для М1
    ' 1-ПФ115 ' 2-МФ78 ' 3-М' 4-КР
    ' Для М2
    ' 1-F ' 3-M
    Enum RegimeEngine ' Режим - код Коды режимов: 1 - ПФ115; 3 - Макс
        Максимал
        Форсаж
    End Enum

    Private Regime As RegimeEngine

#Region "Константы и массивы"
    ''' <summary>
    ''' Константы для расчета
    ''' </summary>
    ''' <remarks></remarks>
    Private Const conКельвин As Double = 273.15 ' абс. ноль
    Private Const con288 As Double = 288.15
    Private Const con_1_033 As Double = 1.033227
    Private Const con_g As Double = 9.80665 ' ускорение силы тяжести м/сек.кв.
    Private Const кон735_6 As Double = 735.6 ' для барометра
    'Private Const Rв As Double = 29.27 ' универсальная газовая постоянная кГм/кг*К
    Private Const Порог As Double = 0.00001 ' измерительный шум
    Private Const НерПоляЗаКВД As Double = 0.99 ' КР2 = 0.99 - поправочный коэффициент, учитывающий неравномерность поля за КВД. 
    Private Const КоэфПолнСгорТоплКС As Double = 0.99 ' коэффициент полноты сгорания топлива в КС; ηКС = 0,99 в соответствии с «Инструкцией на приведение параметров и определение температуры газа при стендовых испытаниях изделия 99». 99.57 ИН [1].
    Private Const СтехиометрКоэфLo As Double = 14.95 ' 14.92914 '' где значение стехиометрического коэффициента Lo принято равным Lo =14.95, аналогично штатному изделию 99 [1]. Уточняется по фактическому сорту используемого топлива.
    Private Const ПерКоэфПрЧастотыРВД As Double = 13300 / 10200 ' переходный коэффициент для приведенной частоты РВД
    Private Const limitOverflow As Integer = 100 ' Счетчик Переп
    Private Const GR As Double = 118.0 ' Где GпроТВД  =118,0  при тарированном значении САТВДо= 305 [см2]
    Private Const САТВДо305 As Double = 305.0

    ' влажность воздуха
    ' зависимость Рп. насыщ.(Тн,К)
    '    Т,К    	253.15	263.15	273.1	283.15	    288.15	    293.15	303.15	313.15	323.15
    ' рП.Н,[Па] 	127.968	287.795	610.381	1227.693	1704.907	2332.75	4238.94	7371.49	12330.25
    Private ReadOnly ТвоздК() As Double = {0, 253.15, 263.15, 273.1, 283.15, 288.15, 293.15, 303.15, 313.15, 323.15} ' для расчета используется INTERP1 - массив начинается с 1
    Private ReadOnly PдавлНасПарВодыПа() As Double = {0, 127.968, 287.795, 610.381, 1227.693, 1704.907, 2332.75, 4238.94, 7371.49, 12330.25}
    ' по формуле ВлагосодВозд=0.622*PдавлНасПаровВодыИнтерпол*(Fi/100)/(101325-PдавлНасПаровВодыИнтерпол*(Fi/100))  ВлагосодВозд -влагосодержание воздуха - количество килограммов водяного пара, приходящегося на 1 кг сухого воздуха

    ' kv - коэффициент, учитывающий входной импульс потока воздуха на входе в изделие.
    Private ReadOnly arralRUD() As Double = {-1, 0, 12, 30, 67, 78, 105, 116} ' для расчета используется INTERP1 - массив начинается с 1
    Private ReadOnly arrKv_alRUD() As Double = {-1, 0, 1.011, 1.017, 1.019, 1.018, 1.015, 1.013}

    ' Массив A [l:6, 1:1 lj содержит информацию, необходимую для расчета термо¬динамических параметров газа и воздуха. В первых семи столбцах содержатся следующие данные:
    ' -	в первой строке - коэффициенты аппроксимации зависимости теплоемкости воздуха от температуры Cpв = ИНТЕГР( Т) полиномом 6-й степени (начиная со свобод¬ного члена);
    ' -	в третьей строке - коэффициенты аппроксимирующего полинома для зависи¬мости CpСО2 = ИНТЕГР( Т)  ;
    ' -	в четвертой строке - коэффициенты аппроксимирующего полинома для зави¬симости CpН2О = ИНТЕГР( Т)  
    ' -	в пятой строке - коэффициенты аппроксимирующего полинома для зависи¬мости CpО2 = ИНТЕГР( Т)  
    ' -	в шестой строке - коэффициенты аппроксимирующего полинома для зависи¬мости CpН2 = ИНТЕГР( Т)  
    ' Массив А приведен в рекомендуемом приложении 5.

    ' В практике расчетов ГТД наиболее часто рассматривается так -называемое "нормальное" углеводородное топливо, весовой состав компонентов которого сле¬дующий: С=0,85, Н = 0,15, 0 = 0.
    ' Случай использования этого топлива при отсутствии в воздухе воды рассмат¬ривается отдельно, коэффициенты полинома, аппроксимирующего зависимость Cpyса=f( Т), рассчитываются по формулам (21) и (22) с использованием
    ' ранее полученных коэффициентов полиномов CpСО2 CpН2О CpО2
    '  что позволяет уменьшить затраты машинного времени в этом
    ' наиболее типичном случае.
    ' Коэффициенты полиномаCpyса =f(Т),для "нормального* топлива содержат¬ся во второй строке массива А, начиная со свободного члена (см. рекомендуемое приложение 5).
    ' Элементы 8, 9 и 10 в каждой строке массива А являются резервными (на случай, если потребуется увеличить степени аппроксимирующих полиномов, см. рекомендуемое приложение 5).
    ' В одиннадцатом столбце массива А содержатся следующие данные. В ячей¬ках A [l, ll] , А[2, ll] и А[3, ll] находятся величины С, Н к О, характери¬зующие весовой состав топлива.
    ' Ячейка А[4, ll] содержит величину sigma, характеризующую влагосодержание воздуха. В ячейке А [5, ll] находится теплотворная способность топлива Ни (для сведения).
    ' Ячейка А[6, 1l] первоначально является свободной: в нее можно поместить, например, значение стехиометрического коэффициента L0 , характеризующее количество килограммов сухого воздуха, потребное для полного сгорания 1 кг топлива. 

    Private ReadOnly AполинТепл(,) As Double = New Double(,) {
        {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
        {0, 0.2521923, -0.1186612, 0.3360775, -0.3073812, 0.1382207, -0.03090246, 0.002745383, 0, 0, 0, 0.85},
        {0, 0.2079764, 1.211806, -1.464097, 1.291195, -0.6385396, 0.1574277, -0.01518199, 0, 0, 0, 0.15},
        {0, 0.1047056, 0.4234367, -0.3953561, 0.2249471, -0.07729786, 0.0146217, -0.001166819, 0, 0, 0, 0},
        {0, 0.4489375, -0.1088401, 0.4027652, -0.2638393, 0.07993751, -0.0115716, 0.000624195, 0, 0, 0, 0},
        {0, 0.2083632, -0.0112279, 0.2235868, -0.2732668, 0.1461334, -0.03687021, 0.003584204, 0, 0, 0, 10331},
        {0, 0.3070881, 0.2230734, -0.4909, 0.5321652, -0.2756533, 0.6851526, -0.06596988, 0, 0, 0, 0}
    }

    ' Массив С [l:6] содержит следующие величины:
    ' C[l] - теплоемкость при постоянном давлении; 
    ' С[2] - газовая постоянная; 
    ' С[4] - показатель адиабаты; 
    ' С[б] - стехиометрический коэффициент LОсух 
    ' Ячейки С [2] и С [б] являются резервными.
    Private C_output(6) As Double

    ' Массив Сr [l : 10] используется в подпрограммах расчета истечения из сопел.
    ' Он содержит следующие величины:
    ' CR[l] - импульс сопла, кгс (при измерении импульса в Ньютонах используется множитель 9,80665);
    ' CR[4] - коэффициент скорости;
    ' CR[5] - площадь сопла;
    ' Ск[б] - перепад давления в сопле ПИс
    ' ПИс = Рф*/Рс
    ' для сопла полного расширения:
    ' ПИс* = Рф*/РН
    ' СR[7] - скорость истечения из сопла; 
    ' СR[9] - давление на срезе; 
    ' СR[10] - коэффициент скорости ФИс.
    ' Ячейки СR[2], СR[3] , СR[8]  здесь не используются. Коэффициент скорости сопла определяется по формуле:
    ' ЛямcS = WCs/aKp,
    ' здесь
    ' aKp  =SQR(2*K/(K+1)*RT*)

    ' 4,2. Подпрограммы для расчета термодинамических параметров воздуха и продуктов сгорания
    ' Для решения перечисленных задач используются шесть подпрограмм -
    ' CPS, DI, TI, PIT, TPI, TKR.
    ' Начало всех подпрограмм идентичное: коэффициенты полинома (27), аппроксимирующего зависимость теплоемкости смеси от температуры, определяются по
    ' коэффициентам аппроксимирующих полиномов для отдельных компонентов смеси. Здесь рассматриваются следующие варианты:
    ' -	чистый воздух;
    ' -	продукты сгорания "нормального* углеводородного топлива ( С - 0,85; Н - 0,15; О - О) в воздухе;
    ' -	продукты сгорания углеводородного топлива с произвольным соотношением компонентов ( С, Н и О) в воздухе;
    ' -	продукты сгорания чистого водорода: α >= 1; α < 1.
    ' Назначение и заголовки указанных подпрограмм рассматриваются на примере процедур, написанных на языке АЛГОЛ-60.

    Private clAir As Air
#End Region

#Region "настроечные параметры"
    ' имена настроечных параметров через константы
    ' строковые значения константы должны совпадать с полем "Name" контролов для этапов 1-6
    ' префикс константы имени начинается control, а величина сохраняется в переменной с префиксом val

    ' Контролы Для Тип Изделия
    Private Const cntКНаклонаОси As String = "КНаклонаОси"
    ' Контролы Для НомерИзделия
    ' ---------------------------
    ' Контролы Для Номер Сборки
    Private Const cntТ4огр As String = "Т4огр"

    ' Контролы Для Номер Постановки
    Private Const cntНомерБокса As String = "НомерБокса"
    Private Const cntТипТоплива As String = "ТипТоплива"
    Private Const cntИПа1мин As String = "ИПа1мин"
    Private Const cntИПа1макс As String = "ИПа1макс"
    Private Const cntа1градмин As String = "а1градмин"
    Private Const cntа1градмакс As String = "а1градмакс"
    Private Const cntИПа2мин As String = "ИПа2мин"
    Private Const cntИПа2макс As String = "ИПа2макс"
    Private Const cntа2градмин As String = "а2градмин"
    Private Const cntа2градмакс As String = "а2градмакс"
    Private Const cntДкрделмин As String = "Дкрделмин"
    Private Const cntДкрделмакс As String = "Дкрделмакс"
    Private Const cntДкрминмм As String = "Дкрминмм"
    Private Const cntДкрмаксмм As String = "Дкрмаксмм"
    Private Const cntFсаТВД As String = "FсаТВД"
    Private Const cntFi As String = "Fi"

    ' Контролы Для Номер Запуска
    ' ---------------------------
    ' Контролы Для Номер КТ
    Private Const cntРежим As String = "Режим"
    Private Const cntТипИспытания As String = "ТипИспытания"
    Private Const cntNктМаксимала As String = "NктМаксимала"

    ' *************************************************************
    ' переменные настроечных параметров 
    ' *************************************************************
    ' Контролы Для Тип Изделия
    Private valКНаклонаОси As Double
    ' Контролы Для Номер Изделия
    ' ---------------------------
    ' Контролы Для Номер Сборки
    Private valТ4огр As Double
    ' Контролы Для Номер Постановки
    Private valНомерБокса As String
    ''' <summary>
    '''переменные зависящие от бокса
    ''' </summary>
    ''' <remarks></remarks>
    Private K1a0 As Double
    Private K1a1 As Double
    Private K1a2 As Double
    Private K1a3 As Double
    Private K2a0 As Double
    Private K2a1 As Double
    Private K2a2 As Double
    Private K2a3 As Double
    Private DM As Double ' диаметр в сечениии M
    Private DB As Double ' диаметр в сечениии B
    Private БазовоеДавлениеВКабине As Boolean
    Private Кпотери As Double

    ''' <summary>
    ''' Контролы Для Номер Постановки
    ''' </summary>
    ''' <remarks></remarks>
    Private valТипТоплива As String
    Private valИПа1мин As Double
    Private valИПа1макс As Double
    Private valа1градмин As Double
    Private valа1градмакс As Double
    Private valИПа2мин As Double
    Private valИПа2макс As Double
    Private valа2градмин As Double
    Private valа2градмакс As Double
    Private valДкрделмин As Double
    Private valДкрделмакс As Double
    Private valДкрминмм As Double
    Private valДкрмаксмм As Double
    Private valFсаТВД As Double
    Private valFi As Double

    ' Контролы Для НомерЗ апуска

    ' Контролы Для Номер КТ
    Private valРежим As String ' М;Ф
    Private valТипИспытания As String ' РПТ;Б;УБ
    Private valNктМаксимала As String

    ''' <summary>
    ''' имена констант, добавлять в массив для проверки в цикле по этим константам
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly arrНастроечныеПараметры As String() = {
            cntКНаклонаОси,
            cntТ4огр,
            cntНомерБокса, cntТипТоплива, cntТипТоплива, cntИПа1мин, cntИПа1макс, cntа1градмин, cntа1градмакс, cntИПа2мин, cntИПа2макс, cntа2градмин, cntа2градмакс, cntДкрделмин, cntДкрделмакс, cntДкрминмм, cntДкрмаксмм, cntFсаТВД, cntFi,
            cntРежим, cntТипИспытания, cntNктМаксимала
            }
#End Region

#Region "измеренные"
    ' 2.2	Параметры, получаемые по результатам непосредственных измерений (прямых или кос венных):
    ' Примечание.Замеры статических давлений Р6 и Р2 вводятся по спецуказанию и использу-ются при анализе параметров и уточнении характеристик узлов.

    '--- измеренные -----------------------------------------------------------
    ' строковые значения константы должны совпадать с полем "ИмяПараметра" таблицы ИзмеренныеПараметры
    ' префикс константы имени начинается cizm
    ' измеренные -  константы для определения ключа в таблице</summary>
    Private Const cizmAlpha_руд As String = "alpha руд" ' a руд , град - положение РУД .
    Private Const cizmВо As String = "Во" ' BA, мм.рт.ст., абс. - атмосферное давление (барометрическое); 
    Private Const cizmPбокса As String = "Pбокса" ' РБ, кгс/см2, абс. - давление воздуха в боксе; 
    Private Const cizmTб As String = "tб" ' tб = tвх °С	- температура воздуха в боксе (на входе в изделие);
    Private Const cizmtтопл As String = "tтопл" ' для вычисления расхода топлива GTKC
    Private Const cizmRизм As String = "Rизм" ' Rизм, кгс	- усилие, получаемое на измерительном устройстве;
    Private Const cizmN1f As String = "N1f" ' n1,%	- частота вращения ротора низкого давления;
    Private Const cizmN2f As String = "N2f" ' n2,%	- частота вращения ротора высокого давления;
    Private Const cizmGт_осн_кс As String = "Gт_осн_кс" ' GTKC, кг/ч	- расход топлива основной камеры сгорания (ОКС);
    ' Private Const cizmЧастота As String = "Частота"
    Private Const cizmGтопл As String = "Gтопл" ' GT, КГ/Ч	- расход топлива суммарный;
    ' Private Const cizmFca_твд As String = "Fca твд" ' площадь сопловых апаратов ТВД
    Private Const cizmT300_физ As String = "T300 физ" ' t2 (или t300), °С - температура торможения на входе в КС; 
    Private Const cizmU447 As String = "U447" ' U447, °С - температура рабочих лопаток турбины высокого давления ( ТВД);
    Private Const cizmT4 As String = "T4" ' T4* , °С - температура газа за ТНД;
    ' Private Const cizmT4огр As String = "T4огр"
    Private Const cizmH_б_с As String = "H(б.с.)" ' Рм, Рв, кгс/см2, абс. - полное давление воздуха в сечениях «М» и «В» расходомерного коллектора (РМК); 
    Private Const cizmB_м_с As String = "B(м.с.)"
    Private Const cizmP2a_квд As String = "P2a квд" ' Р2* (или Рзоо)> кгс/см2, абс. - полное давление воздуха на входе в КС;
    Private Const cizmP2б_квд As String = "P2б квд"
    Private Const cizmP2ст As String = "P2ст" ' P2 (или P301), кгс/см , абс. - статическое давление воздуха на входе в КС *;
    Private Const cizmP6 As String = "P6" ' Р*б кгс/см2, абс. - полное давление воздуха за КНД; 
    Private Const cizmP6ст As String = "P6ст" ' Рб, кгс/см2, абс. - статическое давление воздуха за КНД ); 
    Private Const cizmP4 As String = "P4" ' Р4* , кгс/см , абс. - полное давление газа за ТНД; 
    Private Const cizmP4ст As String = "P4ст"
    Private Const cizmT6 As String = "T6" ' T6 °С	- температура торможения за КНД;
    Private Const cizmA1дел As String = "a1дел" ' a1, а2	- углы ВНА КНД и НА КВД , соответственно;
    Private Const cizmA2дел As String = "a2дел"
    Private Const cizmDрcдел As String = "Dрcдел" ' DКР - диаметр критического сечения PC;
    Private Const cizmОхлаждение As String = "Охлаждение" ' Soxn - сигнальный параметр вкл.\ выкл. охлаждения турбины;
    ' Private Const cizmВлажность As String = "Влажность" ' Ф, % - относительная влажность воздуха;заменен на valFi

    ' для хранения измеренных величин
    Private Alpha_руд As Double
    Private Во As Double ' B0 иди Вк 
    Private Pb As Double ' Pb или Вб 
    Private Tбокса As Double ' tb 
    Private tтопл As Double ' tтопл
    Private Rизм As Double ' Rizm
    Private N1f As Double ' n1
    Private N2f As Double ' n2
    Private Gт_осн_кс As Double '  Private Частота As Double ' freq 
    Private Gтопл As Double ' Gtopl или Gt 
    Private T300_физ As Double ' t300fiz или Т2 ' t300fiz 
    Private U447 As Double ' u447 ' U447 = u447ввод + 273.15
    Private T4 As Double
    Private H_б_с As Double ' Hbc или Н(бст)
    Private B_м_с As Double ' Bmc или В(мст)
    Private P2a_квд As Double ' P2aKVD или Р2а 
    Private P2б_квд As Double ' P2bKVD  или Р2б
    Private P2ст As Double ' P2CT или Р2 (конст) 
    Private P6 As Double ' P6b или Р6полн 
    Private P6ст As Double ' P6CT или Р6стат (конст) 
    Private P4 As Double ' P4b 
    Private P4ст As Double ' P4CT (конст) 
    Private T6 As Double
    Private a1дел As Double ' del1 или  a1 
    Private a2дел As Double ' del2 или  a2 
    Private Dрcдел As Double ' del3 или  Дрс 
    Private Охлаждение As Integer '  n2ВклОхл 
    ' Private Влажность As Double ' Fi,% Относит. влажность воздуха, %
#End Region

#Region "Физические"
    ' 2.3 Основные параметры, получаемые расчетным путем с использованием результатов изме рений:

    ' GB сум, кг/с - суммарный расход воздуха (на входе в изделие);
    ' GB I КГ/С - расход воздуха через внутренний контур (на входе в КВД);
    ' GBII кг/с - расход воздуха через наружный контур (на входе во II контур);

    ' Hk сум	- адиабатический коэффициент полезного действия процесса сжатия в изделии ; 
    ' ТГ , К - температура торможения газа перед ТВД;

    ' 2.8	Индекс «М» соответствует максимальным бесфорсажным режимам работы.
    ' 2.9	Индекс «Ф» и «ОР» соответствует форсированным режимам работы.
    ' 2.10	Физические значения параметров - без индекса.

    '     ИмяПараметра	ОписаниеПараметра
    ' alfa пер	коэффициент полноты сгорания топлива
    ' Cr пер	удельный расход топлива пересчитаный к СА
    ' Gв пер	расход воздуха пересчитаный к СА
    ' Gт пер	расход топлива пересчитаный к СА
    ' m пер	    коэффициент двухконтурности двигателя
    ' n1 пер	обороты КНД пересчитаные по законам регулирования
    ' n2 пер	обороты КВД пересчитаные по законам регулирования
    ' Pi квд	степень повышения давления за КВД
    ' Pi кнд пер	степень повышения давления за КНД
    ' Pi т пер	степень понижения давления на турбине
    ' Pi т пр	степень понижения давления на турбине  
    ' PiKs пр	сумарная степень повышения давления перед КС
    ' R пер	тяга пересчитаная к СА
    ' T4* пер	температура газа за турбиной пересчитаная
    ' Wкнд	    показатель ГДУ . Отношение Pi кнд  к Gв-ха
    ' Т3*пер	температура газа перед ТВД пересчитаная к СА
    ' Т3физ*	температура газа перед ТВД 

    '--- для Физические величин -----------------------------------------------
    ' строковые значения константы должны совпадать с полем "ИмяПараметра" таблицы РасчетныеПараметры
    ' расчетные переменные служат для вычислений и их заносят в таблицу РасчетныеПараметры
    ' префикс константы имени начинается call

    Private Const callа1 As String = "а1"
    Private Const callа2 As String = "а2"
    Private Const callPh As String = "Ph" ' атмосферное давление
    Private Const callPb As String = "Pb" ' базовое давление
    Private Const calltb As String = "tb"
    Private Const callP_вх_КНД As String = "P_вх_КНД"
    Private Const callP6стфиз As String = "P6стфиз" ' ни где не участвует
    Private Const callР6полнФиз As String = "Р6полнФиз"
    Private Const callT6физ As String = "T6физ"
    Private Const callPi_КНД As String = "Pi_КНД" ' Piкнд - степень повышения давления в КНД;
    Private Const callКПД_КНД As String = "КПД_КНД" ' Hкнд - адиабатический коэффициент полезного действия процесса сжатия в КНД;
    Private Const callP2стфиз As String = "P2стфиз" ' ни где не участвует
    Private Const callР2полнФиз As String = "Р2полнФиз"
    Private Const callТ2физ As String = "Т2физ"
    Private Const callРiКВД As String = "РiКВД" ' Piквд - степень повышения давления в КВД;
    Private Const callКПД_КВД As String = "КПД_КВД" ' Hквд - адиабатический коэффициент полезного действия процесса сжатия в КВД. 
    Private Const callPiKсум As String = "PiKсум" ' Pik сум	- суммарная степень повышения давления в изделии;
    Private Const callКПДКсум As String = "КПДКсум"
    Private Const callT3физ As String = "T3физ"
    Private Const callT3ТДРфиз As String = "T3ТДРфиз" ' измеренная через расход ТДР
    Private Const callТ4физ As String = "Т4физ"
    Private Const callP4стфиз As String = "P4стфиз" ' ни где не участвует
    Private Const callР4полнФиз As String = "Р4полнФиз"
    Private Const callPiTсум As String = "PiTсум" ' PiT сум	- суммарная степень понижения давления в турбине;
    Private Const callGвоздВнКфиз As String = "GвоздВнКфиз"
    Private Const callGвоздСумФизВ As String = "GвоздСумФизВ"
    Private Const callRфиз As String = "Rфиз" ' R, кгс - тяга;
    Private Const callGтТДР As String = "GтТДР"
    Private Const callальф_физ As String = "альф_физ" ' aсум	- суммарный коэффициент избытка воздуха;
    Private Const callCR As String = "CR" ' CR, кг/(кгс*час) - удельный расход топлива;
    Private Const callm As String = "m" ' m	- степень двухконтурности;
    Private Const callDpc As String = "Dpc"
    Private Const callU447физ As String = "U447физ"
    Private Const callT4огрфиз As String = "T4огрфиз"

    Private а1 As Double ' alfa1
    Private а2 As Double ' alfa2
    Private Ph As Double
    Private Pбаз As Double
    Private tbфиз As Double
    Private P_вх_КНД As Double ' ppB
    Private P6стфиз As Double ' физ от P6ст '  P6CTb
    Private Р6полнФиз As Double ' P6bb ' 	 Р*6
    Private T6физ As Double ' T6b
    Private Pi_КНД As Double ' PiKND
    Private КПД_КНД As Double ' ethaKND  ' КПД кнд *
    Private P2стфиз As Double ' физ от P2ст
    Private Р2полнФиз As Double ' P2b
    Private Т2физ As Double ' T2b
    Private РiКВД As Double ' PiKVD
    Private КПД_КВД As Double ' ethaKVD
    Private PiKсум As Double ' PiKsum
    Private КПДКсум As Double ' ethaKsum
    Private T3физ As Double
    Private T3ТДРфиз As Double ' T3fizTDR
    Private Т4физ As Double ' T4B
    Private P4стфиз As Double ' физ от P4ст'  P4CTb
    Private Р4полнФиз As Double ' P4bb
    Private PiTсум As Double ' PiT
    Private GвоздВнКфиз As Double
    Private GвоздСумФизВ As Double ' GBphyzB
    Private Rфиз As Double ' Rfiz
    Private GтТДР As Double ' GtoplTDR
    Private альф_физ As Double ' alfafiz
    Private CR As Double
    Private _m As Double
    Private Dpc As Double
    Private U447физ As Double
    Private T4огрфиз As Double
#End Region

#Region "Приведенные"
    ' 2.4	Параметры, подлежащие приведению к САУ :
    ' n1, n2, R, GT, CR, Gbсум, Тг*, t4* , t i, P i, P i* , Gb1 . 
    ' ( t*i, P i, P i* - температура и давление в проточной части изделия в 1 сечении ).
    ' 2.6	Индексом "np" обозначены значения параметров, приведенных к САУ при условии FKp = const.

    '--- для Приведенных величин ----------------------------------------------
    ' строковые значения константы должны совпадать с полем "ИмяПараметра" таблицы РасчетныеПараметры
    ' расчетные переменные служат для вычислений и их заносят в таблицу РасчетныеПараметры
    ' префикс константы имени начинается call

    Private Const calln1пр As String = "n1пр"
    Private Const calln2пр As String = "n2пр"
    Private Const callRпр As String = "Rпр"
    Private Const callCRпр As String = "CRпр"
    Private Const callGтпр As String = "Gтпр"
    Private Const callGвоздСумпр As String = "GвоздСумпр"
    Private Const callальфаSпр As String = "альфаSпр"
    Private Const callУстКНДпр As String = "УстКНДпр"
    Private Const callУстКВДпр As String = "УстКВДпр"
    Private Const callT3пр As String = "T3пр"
    Private Const callT3ТДРпр As String = "T3ТДРпр"
    Private Const callT4пр As String = "T4пр"
    Private Const callu447пр As String = "u447пр"
    Private Const callP6полнПр As String = "P6полнПр"
    Private Const callP6стПр As String = "P6стПр"
    Private Const callT6пр As String = "T6пр"
    Private Const callP2полнПр As String = "P2полнПр"
    Private Const callP2стПр As String = "P2стПр"
    Private Const callT2пр As String = "T2пр"
    Private Const callP4полнПр As String = "P4полнПр"
    Private Const callP4стПр As String = "P4стПр"
    Private Const calln2КВДпр As String = "n2КВДпр"
    Private Const callGКВДпр As String = "GКВДпр"

    Private n1пр As Double
    Private n2пр As Double
    Private Rпр As Double
    Private CRпр As Double
    Private Gтпр As Double
    Private GвоздСумпр As Double
    Private альфаSпр As Double
    Private УстКНДпр As Double
    Private УстКВДпр As Double
    Private T3пр As Double
    Private T3ТДРпр As Double
    Private P4стПр As Double
    Private u447пр As Double
    Private P6полнПр As Double
    Private P6стПр As Double
    Private T6пр As Double
    Private P2полнПр As Double
    Private P2стПр As Double
    Private T2пр As Double
    Private P4полнПр As Double
    Private P4пр As Double
    Private n2КВДпр As Double
    Private GКВДпр As Double
#End Region

#Region "Пересчитанные"
    ' 2.5	Параметры, подлежащие пересчету к САУ на максимальных и форсированных режимах с учётом законов регулирования:
    ' п1, п2, R, GT, CR, Gbсум, aсум, Тг* , t4* , Pi КНд, Pik сум, PiT сум, m .
    ' 2.7	Индексом "пер" обозначены значения параметров, пересчитанных к САУ при условии поддержания заданных законов регулирования на максимальных и форсированных режимах.

    '--- для Пересчитанных величин --------------------------------------------
    ' строковые значения константы должны совпадать с полем "ИмяПараметра" таблицы РасчетныеПараметры
    ' расчетные переменные служат для вычислений и их заносят в таблицу РасчетныеПараметры
    ' префикс константы имени начинается call

    Private Const calln1пер As String = "n1пер"
    Private Const calln2пер As String = "n2пер"
    Private Const callRпер As String = "Rпер"
    Private Const callCRпер As String = "CRпер"
    Private Const callGт_пер As String = "Gт_пер"
    Private Const callGвоздСумпер As String = "GвоздСумпер"
    Private Const callальфа_пер As String = "альфа_пер"
    Private Const callmпер As String = "mпер"
    Private Const callPiКНДпер As String = "PiКНДпер"
    Private Const callPiКsпер As String = "PiКsпер"
    Private Const callT3пер As String = "T3пер"
    Private Const callT4пер As String = "T4пер"
    Private Const callPiКВДпер As String = "PiКВДпер"
    Private Const callGTфор_пер As String = "GTфор_пер"
    Private Const callT3ТДРпер As String = "T3ТДРпер"

    Private n1пер As Double
    Private n2пер As Double
    Private Rпер As Double
    Private CRпер As Double
    Private Gт_пер As Double
    Private GвоздСумпер As Double
    Private альфа_пер As Double
    Private mпер As Double
    Private PiКНДпер As Double
    Private PiКsпер As Double
    Private T3пер As Double
    Private T4пер As Double
    Private PiКВДпер As Double
    Private GTфор_пер As Double
    Private T3ТДРпер As Double
#End Region

#Region "переменные для файлов графиков приведения"
    ' имена графиков приведения для температуры коэф К1 режима М (ОР/Б/УБ)
    ' с номерами рисунков
    Private Const gAkn1 As String = "Akn1" ' 2
    Private Const gAkPiB As String = "AkPiB" ' 3
    Private Const gAkGB As String = "AkGB" ' 4
    Private Const gAkm As String = "Akm" ' 5
    Private Const gAkn2 As String = "Akn2" ' 6
    Private Const gAkPiKS As String = "AkPiKS" ' 7
    Private Const gAkGT0 As String = "AkGT0" ' 8
    Private Const gAkT3 As String = "AkT3" ' 9
    Private Const gAkT4 As String = "AkT4" ' 10
    Private Const gAkR As String = "AkR" ' 11
    Private Const gAkCR As String = "AkCR" ' 12

    ' имена графиков приведения для температуры коэф К1 режима Ф (ОР/Б/УБ)
    ' с номерами рисунков
    Private Const gFkn1 As String = "Fkn1" ' 13
    Private Const gFkPiB As String = "FkPiB" ' 14
    Private Const gFkGB As String = "FkGB" ' 15
    Private Const gFkm As String = "Fkm" ' 16
    Private Const gFkn2 As String = "Fkn2" ' 17
    Private Const gFkPiKVD As String = "FkPiKVD" ' 18
    Private Const gFkPiKS As String = "FkPiKS" ' 19
    Private Const gFkGT0 As String = "FkGT0" ' 20
    Private Const gFkT3 As String = "FkT3" ' 21
    Private Const gFkT4 As String = "FkT4" ' 22
    Private Const gFkGTF As String = "FkGTF" ' 23
    Private Const gFkaS As String = "FkaS" ' 24
    Private Const gFkR As String = "FkR" ' 25
    Private Const gFkCR As String = "FkCR" ' 26

    Private kn1 As Double
    Private kPiB As Double
    Private kGB As Double
    Private km As Double
    Private kn2 As Double
    Private kPiKS As Double
    Private kGT0 As Double
    Private kT3 As Double
    Private kT4 As Double
    Private kR As Double
    Private kCR As Double

    ' требуются только для форсажа
    Private kPiKVD As Double
    Private kGTF As Double
    Private kaS As Double

    ' имена графиков приведения для влажности К2 режима М (ОР/Б/УБ)
    ' с номерами рисунков
    Private Const gAkn1fi As String = "Akn1fi" ' 27
    Private Const gAkPiBfi As String = "AkPiBfi" ' 28
    Private Const gAkGBfi As String = "AkGBfi" ' 29
    Private Const gAkmfi As String = "Akmfi" ' 30
    Private Const gAkn2fi As String = "Akn2fi" ' 31
    Private Const gAkPiKSfi As String = "AkPiKSfi" ' 32
    Private Const gAkGT0fi As String = "AkGT0fi" ' 33
    Private Const gAkT3fi As String = "AkT3fi" ' 34
    Private Const gAkT4fi As String = "AkT4fi" ' 34a
    Private Const gAkRfi As String = "AkRfi" ' 35
    Private Const gAkCRfi As String = "AkCRfi" ' 36

    ' имена графиков приведения для влажности К2 режима Ф (ОР/Б/УБ)
    ' с номерами рисунков
    Private Const gFkn1fi As String = "Fkn1fi" ' 37
    Private Const gFkPiBfi As String = "FkPiBfi" ' 38
    Private Const gFkGBfi As String = "FkGBfi" ' 39
    Private Const gFkmfi As String = "Fkmfi" ' 40
    Private Const gFkn2fi As String = "Fkn2fi" ' 41
    Private Const gFkPiKVDfi As String = "FkPiKVDfi" ' 42
    Private Const gFkPiKSfi As String = "FkPiKSfi" ' 43
    Private Const gFkGT0fi As String = "FkGT0fi" ' 44
    Private Const gFkT3fi As String = "FkT3fi" ' 45
    Private Const gFkT4fi As String = "FkT4fi" ' 46
    Private Const gFkGTFfi As String = "FkGTFfi" ' 47
    Private Const gFkRfi As String = "FkRfi" ' 48
    Private Const gFkCRfi As String = "FkCRfi" ' 49

    Private kn1fi As Double
    Private kPiBfi As Double
    Private kGBfi As Double
    Private kmfi As Double
    Private kn2fi As Double
    Private kPiKSfi As Double
    Private kGT0fi As Double
    Private kT3fi As Double
    Private kT4fi As Double
    Private kRfi As Double
    Private kCRfi As Double

    ' требуются только для форсажа
    Private kPiKVDfi As Double
    Private kGTFfi As Double
#End Region

#Region "переменные для расчета"
    '--- переменные для расчета -----------------------------------------------
    Private Hu As Double ' теплотворная способность топлива зависит от valТипТоплива 
    ' Private B As Double ' барометр база
    ' Private k As Double ' коэф адиабаты
    ''' <summary>
    ''' Запомнить Значения
    ''' выставляется только в определенном диапазоне
    ''' </summary>
    Private isMemoryValue As Boolean = False
    ''' <summary>
    ''' Данные Для Макс Получены
    ''' </summary>
    Private isDataForMaximumCollected As Boolean = False
    Private КоэПриведенияTбокса As Double
    Private КоэПривP_вх_КНД As Double
    Private PдавлНасПаровВоды As Double ' давление насыщенных паров воды
    Private ВлагосодВозд As Double ' влагосодержание воздуха
    Private kv As Double ' коэффициент, учитывающий входной импульс потока воздуха на входе в изделие
    Private njuОхл As Double '(НЮ)коэфф отн расх возд на охл СА ТВД
    Private XпарамАльфа As Double 'связан с коэф избытка воздуха
    'Private Т0масшт, Тмасшт As Double
    '**************************
    ' расчёт основных параметров
    '**************************
    Private РстатМернСечМ, РполнМернСечМ, sigmaKC As Double
    Private КадиабВлВоздуха, КадиабСумм As Double ' коэф-т адиабаты влажного воздуха
    Private КадиабКНД As Double 'ТсреднТбокса_Т6,'ТсреднТбокс_Т2,
    Private ТсреднТ6_Т2, КадиабКВД As Double
    Private ТконечАдиаб, ДействРаботаКНД, АдиабРаботаКНД, ДействРаботаКВД, АдиабРаботаКВД, ДействРаботаКомпрессора, АдиабРаботаКомпрессора As Double
    Private GВоздВГорлеСА, Т3ГазаВГорлеСА As Double

    ' запоминание значения:
    Private u447Memory As Double ' u447 для М режима
    Private PiTMemory As Double ' PiT для М режима
    Private mФизMemory As Double ' m для М режима
    Private T3физMemory As Double ' T3физ для М режима
    Private GвоздВнКфизMemory As Double ' GвоздВнКфиз для М режима
#End Region

    Public Sub New(ByVal Manager As ProjectManager)
        MyBase.New()
        Me.Manager = Manager
        mManagerGraph = New ManagerGraphKof(gMainFomMdiParent.OwnCatalogue)
        clAir = New Air
        clAir.ИнициализацияПеременных()

        Dim arrГрафикиТемпКоэфК1Максимал As String() = {gAkn1, gAkPiB, gAkGB, gAkm, gAkn2, gAkPiKS, gAkGT0, gAkT3, gAkT4, gAkR, gAkCR}
        Dim arrГрафикиТемпКоэфК1Форсаж As String() = {gFkn1, gFkPiB, gFkGB, gFkm, gFkn2, gFkPiKVD, gFkPiKS, gFkGT0, gFkT3, gFkT4, gFkGTF, gFkaS, gFkR, gFkCR}
        Dim arrГрафикиВлажностиКоэфК2Максимал As String() = {gAkn1fi, gAkPiBfi, gAkGBfi, gAkmfi, gAkn2fi, gAkPiKSfi, gAkGT0fi, gAkT3fi, gAkT4fi, gAkRfi, gAkCRfi}
        Dim arrГрафикиВлажностиКоэфК2Форсаж As String() = {gFkn1fi, gFkPiBfi, gFkGBfi, gFkmfi, gFkn2fi, gFkPiKVDfi, gFkPiKSfi, gFkGT0fi, gFkT3fi, gFkT4fi, gFkGTFfi, gFkRfi, gFkCRfi}

        mManagerGraph.ДобавитьГрафики(arrГрафикиТемпКоэфК1Максимал, GraphType.ГрафикТемпКоэфК1, RegimeEngine.Максимал)
        mManagerGraph.ДобавитьГрафики(arrГрафикиТемпКоэфК1Форсаж, GraphType.ГрафикТемпКоэфК1, RegimeEngine.Форсаж)
        mManagerGraph.ДобавитьГрафики(arrГрафикиВлажностиКоэфК2Максимал, GraphType.ГрафикВлажностиКоэфК2, RegimeEngine.Максимал)
        mManagerGraph.ДобавитьГрафики(arrГрафикиВлажностиКоэфК2Форсаж, GraphType.ГрафикВлажностиКоэфК2, RegimeEngine.Форсаж)

        Try
            If mManagerGraph.СчитатьИнициализировать Then
                ' тест
                ' _Режим = РежимEnum.Максимал
                ' Tбокса = 45
                ' N1f = 105
                ' Влажность = 100
                ' ВычислитьРасчетныеПараметры()
            Else
                mManagerGraph = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ClassCalculation.New ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    'Private Function ПроверитьНаличиеЗаписиРасчетныйПараметр(ByVal NameRow As String) As Boolean
    '    'If dt.Columns.Contains(NameColumn) Then
    '    '    Return dt.Columns(NameColumn).Ordinal
    '    'Else
    '    '    Return -1
    '    'End If
    '    'For Each rowРасчетныйПараметр As BaseFormDataSet.РасчетныеПараметрыRow In Manager.РасчетныеПараметры.Rows
    '    '    If rowРасчетныйПараметр.ИмяПараметра = NameRow Then
    '    '        Return True
    '    '        Exit Function
    '    '    End If
    '    'Next
    '    'Return False

    '    Dim rowРасчетныйПараметр As BaseFormDataSet.РасчетныеПараметрыRow = Manager.РасчетныеПараметры.FindByИмяПараметра(NameRow)
    '    Return rowРасчетныйПараметр IsNot Nothing
    'End Function

    Public Sub ПолучитьЗначенияНастроечныхПараметров()
        If Me.Manager.ControlsForPhase.Count = 0 Then Exit Sub

        Dim phase As String
        Dim success As Boolean = False ' Параметр Найден

        ' Вначале проверяется наличие расчетных параметров в базе
        For Each itenName As String In arrНастроечныеПараметры
            success = False

            For I As Integer = 0 To Me.Manager.ControlsForPhase.Count - 1
                If Me.Manager.ControlsForPhase.Item(Me.Manager.StageNames(I)).Count = 0 Then Exit For

                For Each keyControl As MDBControlLibrary.IUserControl In Me.Manager.ControlsForPhase.Item(Me.Manager.StageNames(I)).Values
                    If keyControl.Name = itenName Then
                        success = True
                        Exit For
                    End If
                Next
                If success Then Exit For
            Next

            If success = False Then
                ' перенаправление встроенной ошибки
                RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs($"Настроечный параметр {itenName} среди настроечных контролов не найден!", $"Процедура: {NameOf(ПолучитьЗначенияНастроечныхПараметров)}")) ' не ловит в конструкторе
                ' MessageBox.Show(Message, Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Next

        ' Затем проверяется наличие в расчетном модуле переменных, соответствующих расчетным настроечным
        ' и присвоение им значений
        ' не делать, т.к. есть контролы значения которых не используются - описание
        'With Me.Manager
        '    For I As Integer = 0 To Me.Manager.ControlsForPhase.Count - 1
        '        For Each keyControl As MDBControlLibrary.IUserControl In Me.Manager.ControlsForPhase.Item(Me.Manager.ИменаЭтапов(I)).Values
        '            If arrНастроечныеПараметры.Contains(keyControl.Name) Then
        '                Exit For
        '            Else
        '                Dim Message As String = "Настроечный параметр " & keyControl.Name & " не имеет соответствующей переменной в модуле расчета!"
        '                'перенаправление встроенной ошибки
        '                Dim fireDataErrorEventArgs As New IClassCalculation.DataErrorEventArgs(Message, Description)
        '                RaiseEvent DataError(Me, fireDataErrorEventArgs) 'не ловит в конструкторе
        '                'MessageBox.Show(Message, Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '            End If
        '        Next
        '    Next
        'End With

        Try
            With Me.Manager
                ' Контролы Для Тип Изделия
                phase = "ТипИзделия"
                valКНаклонаОси = Double.Parse(.GetValueTuningParameter(phase, cntКНаклонаОси))

                ' Контролы Для Номер Изделия
                ' ТипЭтапа ="НомерИзделия"
                ' = cdbl(.ПолучитьЗначениеНастроечногоПараметра(ТипЭтапа, control))

                ' Контролы Для Номер Сборки
                phase = "НомерСборки"
                valТ4огр = Double.Parse(.GetValueTuningParameter(phase, cntТ4огр))

                ' Контролы Для Номер Постановки
                phase = "НомерПостановки"
                ' проверить значение с предыдущего запоминания, чтобы не делать лишние запросы
                If valНомерБокса <> .GetValueTuningParameter(phase, cntНомерБокса) Then
                    valНомерБокса = .GetValueTuningParameter(phase, cntНомерБокса)
                    ЗапроситьНастройкиДляСтенда()
                End If

                ' проверить значение с предыдущего запоминания, чтобы не делать лишние запросы
                If valТипТоплива <> .GetValueTuningParameter(phase, cntТипТоплива) Then
                    valТипТоплива = .GetValueTuningParameter(phase, cntТипТоплива)
                    ЗапроситьНастройкиДляТипТоплива()
                End If

                valИПа1мин = Double.Parse(.GetValueTuningParameter(phase, cntИПа1мин))
                valИПа1макс = Double.Parse(.GetValueTuningParameter(phase, cntИПа1макс))
                valа1градмин = Double.Parse(.GetValueTuningParameter(phase, cntа1градмин))
                valа1градмакс = Double.Parse(.GetValueTuningParameter(phase, cntа1градмакс))
                valИПа2мин = Double.Parse(.GetValueTuningParameter(phase, cntИПа2мин))
                valИПа2макс = Double.Parse(.GetValueTuningParameter(phase, cntИПа2макс))
                valа2градмин = Double.Parse(.GetValueTuningParameter(phase, cntа2градмин))
                valа2градмакс = Double.Parse(.GetValueTuningParameter(phase, cntа2градмакс))
                valДкрделмин = Double.Parse(.GetValueTuningParameter(phase, cntДкрделмин))
                valДкрделмакс = Double.Parse(.GetValueTuningParameter(phase, cntДкрделмакс))
                valДкрминмм = Double.Parse(.GetValueTuningParameter(phase, cntДкрминмм))
                valДкрмаксмм = Double.Parse(.GetValueTuningParameter(phase, cntДкрмаксмм))
                valFсаТВД = Double.Parse(.GetValueTuningParameter(phase, cntFсаТВД))
                valFi = Double.Parse(.GetValueTuningParameter(phase, cntFi))

                ' Контролы Для Номер Запуска
                phase = "НомерЗапуска"

                ' КонтролыДляНомерКТ
                phase = "НомерКТ"
                valРежим = .GetValueTuningParameter(phase, cntРежим)
                valТипИспытания = .GetValueTuningParameter(phase, cntТипИспытания)
                valNктМаксимала = .GetValueTuningParameter(phase, cntNктМаксимала)
            End With

            Manager.NeedToRewrite = False ' после получения значений настроечных параметров
        Catch ex As Exception
            ' перенаправление встроенной ошибки
            RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs(ex.Message, $"Процедура: {NameOf(ПолучитьЗначенияНастроечныхПараметров)}")) 'не ловит в конструкторе
            'MessageBox.Show(fireDataErrorEventArgs, Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    ' здесь индивидуальные настройки для класса

    ''' <summary>
    ''' вызывается из mФормаРегистраторЛокальнаяСсылка_frmMainAcquiredData
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calculate() Implements IClassCalculation.Calculate
        If Manager.IsCalculateEnable AndAlso counterUpdate >= Manager.CounterGraph Then
            ' Для Приведенных и Пересчитанных параметров входные единицы измерения
            ' только в единицах СИ, выходные единицы измерения - любого типа
            Manager.ErrorPanelVisible = False

            Try
                ' мне здесь пока не надо получать от контролов
                If Manager.NeedToRewrite Then ПолучитьЗначенияНастроечныхПараметров()
                ' Переводим в Си только измеренные пареметры
                Manager.СonversionToSiUnitMeasurementParameters()
                ' получение абсолютных давлений
                Manager.CalculationBasePressure()
                ' весь подсчет производится исключительно в единицах СИ
                ' извлекаем значения измеренных параметров
                ИзвлечьЗначенияИзмеренныхПараметров()
                ОпределитьРежим()
                ВычислитьРасчетныеПараметры()
                Manager.СonversionToTuningUnitCalculationParameters()
                Manager.AcquisitionValueMeasurementAndCalculateParameters()
            Catch ex As Exception
                ' перенаправление встроенной ошибки
                RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs(ex.Message, $"Процедура: {NameOf(Calculate)}"))
            End Try
            counterUpdate = 0
        End If
        counterUpdate += 1
    End Sub

    ''' <summary>
    ''' вызывается из frmСнятиеКТ.ToolStripButtonПересчетКТ_Click
    ''' так как с параметры можно менять, то режим определяется также пользователем, а не вычисляется
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RecalculationKT() Implements IClassCalculation.RecalculationKT
        Try
            ' мне здесь пока не надо получать от контролов
            If Manager.NeedToRewrite Then ПолучитьЗначенияНастроечныхПараметров()

            ' Переводим в Си только измеренные пареметры
            ' извлекаем значения измеренных параметров которые уже введны на лист и по кнопке пересчета считаны назад в rowИзмеренныйПараметр.ИзмеренноеЗначение
            Manager.СonversionToSiUnitMeasurementParameters()
            ' получение абсолютных давлений
            Manager.CalculationBasePressure()
            ' весь подсчет производится исключительно в единицах СИ
            ИзвлечьЗначенияИзмеренныхПараметров()
            ОпределитьРежимПересчетКТ()
            ВычислитьРасчетныеПараметры()
            Manager.СonversionToTuningUnitCalculationParameters()
        Catch ex As Exception
            ' перенаправление встроенной ошибки
            RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs(ex.Message, $"Процедура: {NameOf(RecalculationKT)}"))
        End Try
    End Sub

    Private Sub ИзвлечьЗначенияИзмеренныхПараметров()
        Try
            With Manager.MeasurementDataTable
                ' расчетные параметры
                ' Tбокса = .FindByИмяПараметра(conTбокса).ЗначениеВСИ ' 	температура в боксе
                ' Барометр = .FindByИмяПараметра(conБарометр).ЗначениеВСИ ' 	БРС1-М
                ' учет атмосферного давления - относительного давления воздуха
                ' B = Барометр / кон735_6
                ' ДавлениеВоздухаНаВходе = .FindByИмяПараметра(conДавлениеВоздухаНаВходе).ЗначениеВСИ + B

                Alpha_руд = .FindByИмяПараметра(cizmAlpha_руд).ЗначениеВСИ
                Во = .FindByИмяПараметра(cizmВо).ЗначениеВСИ
                Pb = .FindByИмяПараметра(cizmPбокса).ЗначениеВСИ
                Tбокса = .FindByИмяПараметра(cizmTб).ЗначениеВСИ
                tтопл = .FindByИмяПараметра(cizmtтопл).ЗначениеВСИ
                Rизм = .FindByИмяПараметра(cizmRизм).ЗначениеВСИ
                N1f = .FindByИмяПараметра(cizmN1f).ЗначениеВСИ
                N2f = .FindByИмяПараметра(cizmN2f).ЗначениеВСИ
                Gт_осн_кс = .FindByИмяПараметра(cizmGт_осн_кс).ЗначениеВСИ
                Gтопл = .FindByИмяПараметра(cizmGтопл).ЗначениеВСИ
                T300_физ = .FindByИмяПараметра(cizmT300_физ).ЗначениеВСИ
                U447 = .FindByИмяПараметра(cizmU447).ЗначениеВСИ
                T4 = .FindByИмяПараметра(cizmT4).ЗначениеВСИ
                H_б_с = .FindByИмяПараметра(cizmH_б_с).ЗначениеВСИ
                B_м_с = .FindByИмяПараметра(cizmB_м_с).ЗначениеВСИ
                P2a_квд = .FindByИмяПараметра(cizmP2a_квд).ЗначениеВСИ
                P2б_квд = .FindByИмяПараметра(cizmP2б_квд).ЗначениеВСИ
                P2ст = .FindByИмяПараметра(cizmP2ст).ЗначениеВСИ
                P6 = .FindByИмяПараметра(cizmP6).ЗначениеВСИ
                P6ст = .FindByИмяПараметра(cizmP6ст).ЗначениеВСИ
                P4 = .FindByИмяПараметра(cizmP4).ЗначениеВСИ
                P4ст = .FindByИмяПараметра(cizmP4ст).ЗначениеВСИ
                T6 = .FindByИмяПараметра(cizmT6).ЗначениеВСИ
                a1дел = .FindByИмяПараметра(cizmA1дел).ЗначениеВСИ
                a2дел = .FindByИмяПараметра(cizmA2дел).ЗначениеВСИ
                Dрcдел = .FindByИмяПараметра(cizmDрcдел).ЗначениеВСИ
                Охлаждение = CInt(.FindByИмяПараметра(cizmОхлаждение).ЗначениеВСИ)
                If Охлаждение > 0 Then Охлаждение = 1

                КоэПриведенияTбокса = System.Math.Sqrt(con288 / (Tбокса + conКельвин))
            End With
        Catch ex As Exception
            ' перенаправление встроенной ошибки
            RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs(ex.Message, $"Процедура: {NameOf(ИзвлечьЗначенияИзмеренныхПараметров)}"))
            Manager.ErrorPanelVisible = True
        End Try
    End Sub

    Private Sub ВозвратЗначенийВычисленныхПараметров()
        Try
            ' занести вычисленные значения
            With Manager.CalculatedDataTable
                ' "Физические"
                .FindByИмяПараметра(callа1).ВычисленноеЗначениеВСИ = а1
                .FindByИмяПараметра(callа2).ВычисленноеЗначениеВСИ = а2
                .FindByИмяПараметра(callPh).ВычисленноеЗначениеВСИ = Ph
                .FindByИмяПараметра(callPb).ВычисленноеЗначениеВСИ = Pбаз
                .FindByИмяПараметра(calltb).ВычисленноеЗначениеВСИ = tbфиз
                .FindByИмяПараметра(callP_вх_КНД).ВычисленноеЗначениеВСИ = P_вх_КНД
                .FindByИмяПараметра(callP6стфиз).ВычисленноеЗначениеВСИ = P6стфиз
                .FindByИмяПараметра(callР6полнФиз).ВычисленноеЗначениеВСИ = Р6полнФиз
                .FindByИмяПараметра(callT6физ).ВычисленноеЗначениеВСИ = T6физ
                .FindByИмяПараметра(callPi_КНД).ВычисленноеЗначениеВСИ = Pi_КНД
                .FindByИмяПараметра(callКПД_КНД).ВычисленноеЗначениеВСИ = КПД_КНД
                .FindByИмяПараметра(callP2стфиз).ВычисленноеЗначениеВСИ = P2стфиз
                .FindByИмяПараметра(callР2полнФиз).ВычисленноеЗначениеВСИ = Р2полнФиз
                .FindByИмяПараметра(callТ2физ).ВычисленноеЗначениеВСИ = Т2физ
                .FindByИмяПараметра(callРiКВД).ВычисленноеЗначениеВСИ = РiКВД
                .FindByИмяПараметра(callКПД_КВД).ВычисленноеЗначениеВСИ = КПД_КВД
                .FindByИмяПараметра(callPiKсум).ВычисленноеЗначениеВСИ = PiKсум
                .FindByИмяПараметра(callКПДКсум).ВычисленноеЗначениеВСИ = КПДКсум
                .FindByИмяПараметра(callT3физ).ВычисленноеЗначениеВСИ = T3физ
                .FindByИмяПараметра(callT3ТДРфиз).ВычисленноеЗначениеВСИ = T3ТДРфиз
                .FindByИмяПараметра(callТ4физ).ВычисленноеЗначениеВСИ = Т4физ
                .FindByИмяПараметра(callP4стфиз).ВычисленноеЗначениеВСИ = P4стфиз
                .FindByИмяПараметра(callР4полнФиз).ВычисленноеЗначениеВСИ = Р4полнФиз
                .FindByИмяПараметра(callPiTсум).ВычисленноеЗначениеВСИ = PiTсум
                .FindByИмяПараметра(callGвоздВнКфиз).ВычисленноеЗначениеВСИ = GвоздВнКфиз
                .FindByИмяПараметра(callGвоздСумФизВ).ВычисленноеЗначениеВСИ = GвоздСумФизВ
                .FindByИмяПараметра(callRфиз).ВычисленноеЗначениеВСИ = Rфиз
                .FindByИмяПараметра(callGтТДР).ВычисленноеЗначениеВСИ = GтТДР
                .FindByИмяПараметра(callальф_физ).ВычисленноеЗначениеВСИ = альф_физ
                .FindByИмяПараметра(callCR).ВычисленноеЗначениеВСИ = CR
                .FindByИмяПараметра(callm).ВычисленноеЗначениеВСИ = _m
                .FindByИмяПараметра(callDpc).ВычисленноеЗначениеВСИ = Dpc
                .FindByИмяПараметра(callU447физ).ВычисленноеЗначениеВСИ = U447физ
                .FindByИмяПараметра(callT4огрфиз).ВычисленноеЗначениеВСИ = T4огрфиз

                ' "Приведенные"
                .FindByИмяПараметра(calln1пр).ВычисленноеЗначениеВСИ = n1пр
                .FindByИмяПараметра(calln2пр).ВычисленноеЗначениеВСИ = n2пр
                .FindByИмяПараметра(callRпр).ВычисленноеЗначениеВСИ = Rпр
                .FindByИмяПараметра(callCRпр).ВычисленноеЗначениеВСИ = CRпр
                .FindByИмяПараметра(callGтпр).ВычисленноеЗначениеВСИ = Gтпр
                .FindByИмяПараметра(callGвоздСумпр).ВычисленноеЗначениеВСИ = GвоздСумпр
                .FindByИмяПараметра(callальфаSпр).ВычисленноеЗначениеВСИ = альфаSпр
                .FindByИмяПараметра(callУстКНДпр).ВычисленноеЗначениеВСИ = УстКНДпр
                .FindByИмяПараметра(callУстКВДпр).ВычисленноеЗначениеВСИ = УстКВДпр
                .FindByИмяПараметра(callT3пр).ВычисленноеЗначениеВСИ = T3пр
                .FindByИмяПараметра(callT3ТДРпр).ВычисленноеЗначениеВСИ = T3ТДРпр
                .FindByИмяПараметра(callT4пр).ВычисленноеЗначениеВСИ = P4стПр
                .FindByИмяПараметра(callu447пр).ВычисленноеЗначениеВСИ = u447пр
                .FindByИмяПараметра(callP6полнПр).ВычисленноеЗначениеВСИ = P6полнПр
                .FindByИмяПараметра(callP6стПр).ВычисленноеЗначениеВСИ = P6стПр
                .FindByИмяПараметра(callT6пр).ВычисленноеЗначениеВСИ = T6пр
                .FindByИмяПараметра(callP2полнПр).ВычисленноеЗначениеВСИ = P2полнПр
                .FindByИмяПараметра(callP2стПр).ВычисленноеЗначениеВСИ = P2стПр
                .FindByИмяПараметра(callT2пр).ВычисленноеЗначениеВСИ = T2пр
                .FindByИмяПараметра(callP4полнПр).ВычисленноеЗначениеВСИ = P4полнПр
                .FindByИмяПараметра(callP4стПр).ВычисленноеЗначениеВСИ = P4пр
                .FindByИмяПараметра(calln2КВДпр).ВычисленноеЗначениеВСИ = n2КВДпр
                .FindByИмяПараметра(callGКВДпр).ВычисленноеЗначениеВСИ = GКВДпр

                ' "Пересчитанные"
                .FindByИмяПараметра(calln1пер).ВычисленноеЗначениеВСИ = n1пер
                .FindByИмяПараметра(calln2пер).ВычисленноеЗначениеВСИ = n2пер
                .FindByИмяПараметра(callRпер).ВычисленноеЗначениеВСИ = Rпер
                .FindByИмяПараметра(callCRпер).ВычисленноеЗначениеВСИ = CRпер
                .FindByИмяПараметра(callGт_пер).ВычисленноеЗначениеВСИ = Gт_пер
                .FindByИмяПараметра(callGвоздСумпер).ВычисленноеЗначениеВСИ = GвоздСумпер
                .FindByИмяПараметра(callальфа_пер).ВычисленноеЗначениеВСИ = альфа_пер
                .FindByИмяПараметра(callmпер).ВычисленноеЗначениеВСИ = mпер
                .FindByИмяПараметра(callPiКНДпер).ВычисленноеЗначениеВСИ = PiКНДпер
                .FindByИмяПараметра(callPiКsпер).ВычисленноеЗначениеВСИ = PiКsпер
                .FindByИмяПараметра(callT3пер).ВычисленноеЗначениеВСИ = T3пер
                .FindByИмяПараметра(callT4пер).ВычисленноеЗначениеВСИ = T4пер
                .FindByИмяПараметра(callPiКВДпер).ВычисленноеЗначениеВСИ = PiКВДпер
                .FindByИмяПараметра(callGTфор_пер).ВычисленноеЗначениеВСИ = GTфор_пер
                .FindByИмяПараметра(callT3ТДРпер).ВычисленноеЗначениеВСИ = T3ТДРпер
            End With

        Catch ex As Exception
            ' перенаправление встроенной ошибки
            RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs(ex.Message, $"Процедура: {NameOf(ВозвратЗначенийВычисленныхПараметров)}"))
            Manager.ErrorPanelVisible = True
        End Try
    End Sub

    Private Sub ОпределитьРежим()
        Manager.ShowEngineRegime(Режим(Alpha_руд, N2f * КоэПриведенияTбокса))
        Select Case Regime
            Case RegimeEngine.Максимал
                Manager.Regime = "М"
            Case RegimeEngine.Форсаж
                Manager.Regime = "Ф"
        End Select
    End Sub

    Private Function Режим(ARUD As Double, N2привед As Double) As String
        Regime = RegimeEngine.Максимал ' по умолчанию
        isMemoryValue = False

        If ARUD <= 8 Then
            Return vbNullString
        ElseIf (ARUD > 8 AndAlso ARUD < 14) AndAlso N2привед < 80 Then
            Return "МГ" ' Малый газ
        ElseIf (ARUD >= 14 AndAlso ARUD < 65) AndAlso N2привед < 86 Then
            Return "КРппрод" ' Крейсерский продолжительный
        ElseIf ARUD < 65 AndAlso (N2привед >= 86 AndAlso N2привед < 90) Then
            Return "КР" ' Крейсерский
        ElseIf ARUD < 65 AndAlso (N2привед >= 90 AndAlso N2привед < 93.5) Then
            Return "92%" ' N2прив 92%
        ElseIf ARUD < 65 AndAlso (N2привед >= 93.5 AndAlso N2привед < 97) Then
            Return "95%" ' N2прив 95%
            'ElseIf (sngARUD >= 65 AndAlso sngARUD < 69) AndAlso sngN2привед >= 97 Then
        ElseIf (ARUD >= 65 AndAlso ARUD < 69) Then
            isMemoryValue = True
            Return "M" ' Максимальный режим
        ElseIf (ARUD > 69 AndAlso ARUD < 78) AndAlso N2привед >= 97 Then
            Regime = RegimeEngine.Форсаж
            Return "Фмин" ' Форсаж минимальный
        ElseIf (ARUD > 78 AndAlso ARUD < 113) AndAlso N2привед >= 97 Then
            Regime = RegimeEngine.Форсаж
            Return "Фпол" ' Форсаж полный
        ElseIf ARUD > 113 AndAlso N2привед >= 97 Then
            Regime = RegimeEngine.Форсаж
            Return "ОР" ' Особый режим
        Else
            Return vbNullString
        End If
    End Function

    Private Sub ОпределитьРежимПересчетКТ()
        ' для непрерывного сбора:
        ' ввести признак получения данных ДанныеДляМаксПолучены для режима максимал для данного запуска
        ' он устанавливается если выполнено условие максимала РежимМакс = True
        ' если РежимМакс = True то заполнить переменные хранящие нужные величины(пирометр) и установить ДанныеДляМаксПолучены=True
        ' если выполнено условие РежимФорсаж=True AND ДанныеДляМаксПолучены=True,  то можно сделать косвенный расчет. (иначе константы99999 - может прыгать шлейф) или ни чего не вычислять
        ' вообще - то условия РежимМакс и РежимФорсаж не обязаны идти вслед друг за другом, а могут только в определенном диапазоне
        ' в любом случае если РежимМакс = True, то запоминается последние нужные значения для расчета и выставляется ДанныеДляМаксПолучены=True

        ' индекс запуска (keyНомерЗапуска) где искать КТ в базе определить как самый последний в базе (ведь вначале должен быть сделан запуск, затем записаться хоть одна КТ с режимом Максимал)
        ' номер КТ максимала есть самый последний из всех записанных номеров (если только КТ с режимом Форсаж не самая первая из записанных в данном запуске)
        '   должен где-то запоминаться
        ' этот номер должен быть обязательно заполнен иначе при повторном пересчёте данные с режима Максимал с этим номером не считаются и пересчет Форсажа не будет произведен
        ' он запоминается и Private Sub cmdOKClickEvent(ByVal IndexPage As Integer)
        ' ЗапомнитьЗначения = True

        ' для пересчёта КТ:
        ' Если подсчитывается Форсаж - проверить ввод номера КТ для максимала
        ' считать настроечные параметры режима из записи по нему узнаем с каким режимом был произведен расчет и запись КТ
        Regime = RegimeEngine.Максимал ' по умолчанию

        Select Case valРежим
            Case "М"
                Regime = RegimeEngine.Максимал
            Case "Ф"
                Regime = RegimeEngine.Форсаж
        End Select

        If Regime = RegimeEngine.Форсаж AndAlso valNктМаксимала <> "0" Then
            isDataForMaximumCollected = ЗапроситьПараметрыДляРасчетаФорсажа(valNктМаксимала)
        End If
    End Sub

    Private Sub ВычислитьРасчетныеПараметры()
        ' вначале обязательно обнулить расчетные величины, чтобы они не запоминались при смене расчета максимал и форсаж
        Try
            ВычислитьКоэфПересчета()
            ВычислитьФизические()
            ' далее произвести пересчет в соответствии с коэффициентами
            ВычислитьПриведенные()
            ВычислитьПересчитанные()
            ВозвратЗначенийВычисленныхПараметров()

            ''занести вычисленные значения
            'With Manager
            '    ''по имени параметра strИмяПараметраГрафика определяем нужную функцию приведения
            '    ''("n1") 'который измеряет
            '    ''должна быть вызвана функция приведения параметра "n1" например
            '    'If n1ГПриводить Then
            '    '    'должна быть вызвана функция приведения параметра "n1" например
            '    '    .РасчетныеПараметры.FindByИмяПараметра(cn1Г).ВычисленноеЗначениеВСИ = Air.funПривестиN(.ИзмеренныеПараметры.FindByИмяПараметра("n1").ЗначениеВСИ, tm)
            '    'Else
            '    '    'приводить не надо, просто копирование
            '    '    .РасчетныеПараметры.FindByИмяПараметра(cn1Г).ВычисленноеЗначениеВСИ = .ИзмеренныеПараметры.FindByИмяПараметра("n1").ЗначениеВСИ
            '    'End If

            '    ''cGвМПолеДавлений
            '    '.РасчетныеПараметры.FindByИмяПараметра(cGвМПолеДавлений).ВычисленноеЗначениеВСИ = funВычислитьGвМПолеДавлений(GвМПолеДавленийПриводить)
            '    ''или так
            '    ''.РасчетныеПараметры.FindByИмяПараметра(cGвМПолеДавлений).ВычисленноеЗначениеВСИ = funВычислитьGвМПолеДавлений(.НастроечныеПараметры.FindByИмяПараметра("GвМПолеДавленийПриводить").ЛогическоеЗначение)

            '    ''cnИГ_03Г
            '    ''который измеряет
            '    'nИГ_03 = .ИзмеренныеПараметры.FindByИмяПараметра("nИГ-03").ЗначениеВСИ / 46325 'коэф. перевода  n=1 при N=45190
            'End With
        Catch ex As Exception
            ' перенаправление встроенной ошибки
            RaiseEvent DataError(Me, New IClassCalculation.DataErrorEventArgs(ex.Message, $"Процедура: {NameOf(ВычислитьРасчетныеПараметры)}"))
            Manager.ErrorPanelVisible = True
        End Try
    End Sub

    Private Sub ВычислитьКоэфПересчета()
        mManagerGraph.ВычислитьВсеКоэфПересчета(Tбокса, N1f, valFi, Regime)
        Select Case Regime
            Case RegimeEngine.Максимал
                ' температура
                kn1 = mManagerGraph.КоэфПересчета(gAkn1) ' 2
                kPiB = mManagerGraph.КоэфПересчета(gAkPiB) ' 3
                kGB = mManagerGraph.КоэфПересчета(gAkGB) ' 4
                km = mManagerGraph.КоэфПересчета(gAkm) ' 5
                kn2 = mManagerGraph.КоэфПересчета(gAkn2) ' 6
                kPiKS = mManagerGraph.КоэфПересчета(gAkPiKS) ' 7
                kGT0 = mManagerGraph.КоэфПересчета(gAkGT0) ' 8
                kT3 = mManagerGraph.КоэфПересчета(gAkT3) ' 9
                kT4 = mManagerGraph.КоэфПересчета(gAkT4) ' 10
                kR = mManagerGraph.КоэфПересчета(gAkR) ' 11
                kCR = mManagerGraph.КоэфПересчета(gAkCR) ' 12

                ' требуются только для форсажа
                kPiKVD = 1 ' 14
                kGTF = 1 ' 15
                kaS = 1 ' 13

                ' влажность
                kn1fi = mManagerGraph.КоэфПересчета(gAkn1fi) ' 2
                kPiBfi = mManagerGraph.КоэфПересчета(gAkPiBfi) ' 3
                kGBfi = mManagerGraph.КоэфПересчета(gAkGBfi) ' 4
                kmfi = mManagerGraph.КоэфПересчета(gAkmfi) ' 5
                kn2fi = mManagerGraph.КоэфПересчета(gAkn2fi) ' 6
                kPiKSfi = mManagerGraph.КоэфПересчета(gAkPiKSfi) ' 7
                kGT0fi = mManagerGraph.КоэфПересчета(gAkGT0fi) ' 8
                kT3fi = mManagerGraph.КоэфПересчета(gAkT3fi) ' 9
                kT4fi = mManagerGraph.КоэфПересчета(gAkT4fi) ' 10
                kRfi = mManagerGraph.КоэфПересчета(gAkRfi) ' 11
                kCRfi = mManagerGraph.КоэфПересчета(gAkCRfi) ' 12

                ' требуются только для форсажа
                kPiKVDfi = 1
                kGTFfi = 1
            Case RegimeEngine.Форсаж
                ' температура
                kn1 = mManagerGraph.КоэфПересчета(gFkn1)
                kPiB = mManagerGraph.КоэфПересчета(gFkPiB)
                kGB = mManagerGraph.КоэфПересчета(gFkGB)
                km = mManagerGraph.КоэфПересчета(gFkm)
                kn2 = mManagerGraph.КоэфПересчета(gFkn2)
                kPiKS = mManagerGraph.КоэфПересчета(gFkPiKS)
                kGT0 = mManagerGraph.КоэфПересчета(gFkGT0)
                kT3 = mManagerGraph.КоэфПересчета(gFkT3)
                kT4 = mManagerGraph.КоэфПересчета(gFkT4)
                kR = mManagerGraph.КоэфПересчета(gFkR)
                kCR = mManagerGraph.КоэфПересчета(gFkCR)

                ' требуются только для форсажа
                kPiKVD = mManagerGraph.КоэфПересчета(gFkPiKVD)
                kGTF = mManagerGraph.КоэфПересчета(gFkGTF)
                kaS = mManagerGraph.КоэфПересчета(gFkaS)

                ' влажность
                kn1fi = mManagerGraph.КоэфПересчета(gFkn1fi)
                kPiBfi = mManagerGraph.КоэфПересчета(gFkPiBfi)
                kGBfi = mManagerGraph.КоэфПересчета(gFkGBfi)
                kmfi = mManagerGraph.КоэфПересчета(gFkmfi)
                kn2fi = mManagerGraph.КоэфПересчета(gFkn2fi)
                kPiKSfi = mManagerGraph.КоэфПересчета(gFkPiKSfi)
                kGT0fi = mManagerGraph.КоэфПересчета(gFkGT0fi)
                kT3fi = mManagerGraph.КоэфПересчета(gFkGT0fi)
                kT4fi = mManagerGraph.КоэфПересчета(gFkT4fi)
                kRfi = mManagerGraph.КоэфПересчета(gFkRfi)
                kCRfi = mManagerGraph.КоэфПересчета(gFkCRfi)

                ' требуются только для форсажа
                kPiKVDfi = mManagerGraph.КоэфПересчета(gFkPiKVDfi)
                kGTFfi = mManagerGraph.КоэфПересчета(gFkGTFfi)
        End Select
    End Sub

    Private Sub ВычислитьФизические()
        ' влажность воздуха
        PдавлНасПаровВоды = InterpLine(ТвоздК, PдавлНасПарВодыПа, Tбокса + conКельвин) ' давление насыщенных паров воды
        ВлагосодВозд = 0.622 * PдавлНасПаровВоды * valFi / 100 / (101325 - PдавлНасПаровВоды * valFi / 100) ' влагосодержание воздуха 

        If (a1дел <> 0) Then а1 = valа1градмакс + (a1дел - valИПа1макс) * (valа1градмакс - valа1градмин) / (valИПа1макс - valИПа1мин)
        If (a2дел <> 0) Then а2 = valа2градмакс + (a2дел - valИПа2макс) * (valа2градмакс - valа2градмин) / (valИПа2макс - valИПа2мин)
        If (Dрcдел <> 0) Then Dpc = valДкрминмм + (Dрcдел - valДкрделмин) * (valДкрмаксмм - valДкрминмм) / (valДкрделмакс - valДкрделмин)

        U447физ = U447 + conКельвин
        tbфиз = Tбокса + conКельвин
        Ph = Во / кон735_6

        If БазовоеДавлениеВКабине Then
            ' - базовым явл давл в боксе:  Pбокса
            Pбаз = Pb / кон735_6    ' Рbaz- базовое
        Else '-  базовым явл давл в кабине В0
            Pбаз = Ph
        End If

        ' 4 ОПРЕДЕЛЕНИЕ ПАРАМЕТРОВ ПО ТРАКТУ ИЗДЕЛИЯ 
        ' 4.1 Расходомерный коллектор (РМК). Температура воздуха t*ex на входе в изделие определяется как среднее арифметическое из Nt измерений t*вx/i , производимых приемниками, установленными на входной сетке:
        ' t*вx =1/N*Сумма(t*вxi) (обычно Nt = 4); 
        ' T*вx =conКельвин + t*вx:	(4.1.1)
        ' Алгоритмы расчета суммарного расхода воздуха Gbсум и полных давлений Рм* Рв*  в сечениях «М» и «В» перед двигателем приведены в ПРИЛОЖЕНИИ I.
        РстатМернСечМ = Pбаз - Math.Abs(H_б_с / кон735_6) ' статическое давл в мерн сеч М
        РполнМернСечМ = Pбаз - Math.Abs(B_м_с / кон735_6) ' полное давл в мерн сеч М

        XпарамАльфа = 0 ' для чистого воздуха
        CPS(C_output, tbфиз, XпарамАльфа, AполинТепл, ВлагосодВозд) ' расчёт нового коэф адиабаты
        КадиабВлВоздуха = C_output(4) ' коэф-т адиабаты влажного воздуха 
        ПараметрыРасходомерногоКоллектора(tbфиз, РстатМернСечМ, РполнМернСечМ, GвоздСумФизВ, P_вх_КНД, КадиабВлВоздуха) ' параметры воздуха с учётом влажности

        ' 4.2 Параметры КНД. 4.2.1 Степень повышения давления воздуха в КНД :
        ' Pi КНд= Рб* / Рв*  	(4.2.1)
        If (P6ст < Порог) Then ' измеренные
            P6стфиз = 0   ' -если отсутствует замер P6ст '  P6стфиз- стат давл за КНД
        Else
            P6стфиз = Pбаз + P6ст '  P6стфиз-истинное значение стат давления
            '	P6ст -значение стат давления, измеренное диффер датчиком
        End If
        Р6полнФиз = Pбаз + P6 ' Р6* - полное давл за КНД

        If (T6 < Порог) Then
            T6физ = 0
        Else
            T6физ = T6 + conКельвин ' это Т*6 за КНД
        End If

        Pi_КНД = Р6полнФиз / P_вх_КНД ' (4.2.1)

        ' Параметры КНД: 
        ' ТсреднТбокса_Т6 = 0.5 * (Tбокса + T6) ' t*6,C
        КадиабКНД = Kadiabat((Tбокса + T6) / 2)

        If (T6физ = tbфиз OrElse T6 < Порог) Then
            КПД_КНД = 0
        Else
            ' 4.2.2 Адиабатический коэффициент полезного действия КНД :
            ' Hкнд is  = Lk_АД  / Lk =  (Ib(T*6ад) - Ib(T*вх))/(Ib(T*6) - Ib(T*вх))
            ' [11] (4.21,)
            ' где Lk -действительная работа сжатия 1 кг воздуха в КНД, определяемая по разности энтальпий на выходе и входе:
            ' Lk=IB(T*6)- IB(Т*вх), кДж/кг	(4.2.2.2)
            ' адиабатическая работа сжатия 1 кг воздуха в КНД, определяемая в изоэнтропическом процессе сжатия по разности энтальпий на выходе и входе:
            ' Lk ад=Iв(Т* 6ад)- Iв(Т*вх), КДж/кг                     (4.2.2.3)
            ' Iв(Т *6ад)" энтальпия воздуха в конце изоэнтропического процесса повышения давления воздуха от полного давления на входе в КНД Р*b до полного давления за КНД Р*6Ср, кДж/кг
            ' Т*бад - температура в конце изоэнтропического процесса повышения давления воздуха в КНД определяется из уравнения адиабаты:
            ' SB(T*6aд)= Sb(T* BX)+RB Ln(P* 6 / P*B)	(4.2.2.4) 
            ' старая формула
            ' КПД_КНД = tbфиз * (Pi_КНД ^ ((КадиабКНД - 1) / КадиабКНД) - 1) / (T6физ - tbфиз)    ' (4.2.2)

            ' кпд КНД через энтальпию с учётом влажности воздуха k=k(d)
            XпарамАльфа = 0.0 '  для воздуха весовая доля топлива равна нулю                    
            ТконечАдиаб = TPI(Pi_КНД, tbфиз, XпарамАльфа, AполинТепл, C_output, ВлагосодВозд) ' адиабатическая температура 

            ' работа сжатия 1 кг воздуха в КНД равна разности энтальпий на выходе и входе в КНД
            XпарамАльфа = 0.0
            ДействРаботаКНД = DI(tbфиз, T6физ, XпарамАльфа, AполинТепл, ВлагосодВозд)

            ' располагаемая адиабатическая работа турбины 1 кг газа
            XпарамАльфа = 0.0
            АдиабРаботаКНД = DI(tbфиз, ТконечАдиаб, XпарамАльфа, AполинТепл, ВлагосодВозд)
            КПД_КНД = АдиабРаботаКНД / ДействРаботаКНД ' кпд КНД                      
        End If

        If (P2ст < Порог) Then
            P2стфиз = 0
        Else
            P2стфиз = P2ст + Pбаз
        End If

        ' Избавление от нулевого замера  
        If (P2a_квд < Порог) Then P2a_квд = P2б_квд
        If (P2б_квд < Порог) Then P2б_квд = P2a_квд

        Р2полнФиз = Pбаз + (P2a_квд + P2б_квд) * НерПоляЗаКВД / 2 'полное давл за КВД
        Т2физ = conКельвин + T300_физ 'температура перед КС

        ' Параметры КВД:
        If (Р6полнФиз < Порог) Then
            РiКВД = 0
        Else
            РiКВД = Р2полнФиз / Р6полнФиз
        End If

        ТсреднТ6_Т2 = (T300_физ + T6) / 2 ' T300_физ=Т_2-за КВД ,t6b- за КНД  
        ' КадиабКВД = Kadiabat(ТсреднТ6_Т2)' убрал т.к. нужен влажный воздух
        T6физ = T6 + conКельвин

        If (T6 < Порог) Then
            КПД_КВД = 0
            T6физ = 0.0
        Else
            ' Учёт влажности воздуха k=k(d) КВД
            XпарамАльфа = 0 '  для чистого воздуха
            CPS(C_output, ТсреднТ6_Т2 + conКельвин, XпарамАльфа, AполинТепл, ВлагосодВозд)
            КадиабКВД = C_output(4) '  коэф-т адиабаты влажного воздуха

            ' 4.3 Параметры КВД . 4.3.1 Степень повышения давления воздуха в КВД :
            ' Pi квд =P*2ср / Рб*	(4.3.1)
            ' где: Р*2ср = Р*2 * КР2 среднее значение давления в сечении "2" ;	(4.3.2)
            ' КР2 = 0.99 - поправочный коэффициент, учитывающий неравномерность поля за КВД. 
            ' 4.3.2 Адиабатический коэффициент полезного действия КВД :
            ' Hквд is =  L квд _АД  / Lквд =  
            ' 111  (4ЛЗ)
            ' где Lквд - действительная работа сжатия 1 кг воздуха в КВД, определяемая по разности энтальпий на выходе и входе:
            ' Lквд =Iв(Т*2)-Iв(Т*б), кДж/кг	(4.3.3.1)
            ' адиабатическая работа сжатия 1 кг воздуха в КВД, определяемая в изоэнтропическом процессе сжатия по разности энтальпий на выходе и входе:
            ' Lквд_АД=Iв(Т* 2ад) - Iв(Т *6), кДж/кг	(4.3.3.2)
            ' Iв(Т* 2ад) энтальпия конце изоэнтропического процесса повышения давления воздуха от полного давления на входе в КВД Р* бср до полного давления за компрессором Р* 2сР, Дж/кг
            ' Т*2ад - температура в конце изоэнтропического процесса повышения давления воздуха в КВД определяется из уравнения адиабаты:
            ' SB(T*2a) = SB(T*6) + RBLn(P*2cp / P*6)	(4.3.3.3)
            ' старая формула
            ' КПД_КВД = T6физ * (РiКВД ^ ((КадиабКВД - 1) / КадиабКВД) - 1) / (Т2физ - T6физ)

            ' кпд КВД через энтальпию с учётом влажности воздуха k=k(d)
            XпарамАльфа = 0.0 '  для воздуха весовая доля топлива равна нулю                   
            ТконечАдиаб = TPI(РiКВД, T6физ, XпарамАльфа, AполинТепл, C_output, ВлагосодВозд) ' адиабатическая температура    
            ' работа сжатия 1 кг воздуха в КВД равна разности энтальпий на выходе и входе в КВД
            XпарамАльфа = 0.0
            ДействРаботаКВД = DI(T6физ, Т2физ, XпарамАльфа, AполинТепл, ВлагосодВозд)
            ' располагаемая адиабатическая работа турбины 1 кг газа
            XпарамАльфа = 0.0
            АдиабРаботаКВД = DI(T6физ, ТконечАдиаб, XпарамАльфа, AполинТепл, ВлагосодВозд)
            КПД_КВД = АдиабРаботаКВД / ДействРаботаКВД
        End If

        ' 4.4 Суммарные параметры компрессора.
        ' 4.4.1	Суммарная степень повышения давления воздуха в изделии :
        ' Pik сумм = P*2ср / Рв*  	(4-4.1)
        PiKсум = Р2полнФиз / P_вх_КНД
        ' ТсреднТбокс_Т2 = 0.5 * (Tбокса + T300_физ)
        КадиабСумм = Kadiabat((Tбокса + T300_физ) / 2)

        ' 4.4.2	Адиабатический коэффициент полезного действия процесса сжатия в изделии:
        ' Hk сум is =  L сум _АД  / Lксум  =
        ' 	 (4.4.2)
        ' где L Lксум  - действительная работа сжатия 1 кг воздуха в изделии, определяемая по разности энтальпий на выходе и входе:
        ' Lксум =Iв(Т*2)- Iв(Т*вх), кДж/кг	(4.4.2.1)
        ' адиабатическая работа сжатия 1 кг воздуха в КВД, определяемая в изоэнтропическом процессе сжатия по разности энтальпий на выходе и входе:
        ' L сум _АД  =Iв(Т*2ад)- Iв(Т*вх), кДж/кг	(4.4.2.2)
        ' Iв(Т*2ад) = энтальпия воздуха в конце изоэнтропического процесса повышения давления воздуха от полного давления на входе в КНД Р* в до полного давления за компрессором P*2ср, кДж/кг
        ' Т*2ад - температура в конце изоэнтропического процесса повышения давления воздуха в изделии определяется из уравнения адиабаты:
        ' SB(Т*2ад ) = SB(T*BX)+ RB Ln (P*2/P*B)	(4.4.2.3)
        ' старая формула
        ' КПДКсум = tbфиз * (PiKсум ^ ((КадиабСумм - 1) / КадиабСумм) - 1) / (Т2физ - tbфиз)  '   (4.4.2)

        ' суммарный кпд компрессора через энтальпию с учётом влажности воздуха k=k(d)
        XпарамАльфа = 0.0 ' для воздуха весовая доля топлива равна нулю  
        ТконечАдиаб = TPI(PiKсум, tbфиз, XпарамАльфа, AполинТепл, C_output, ВлагосодВозд) ' адиабатическая температура
        ' работа сжатия 1 кг воздуха в КВД равна разности энтальпий на выходе КВД и входе в компрессор КНД
        XпарамАльфа = 0.0
        ДействРаботаКомпрессора = DI(tbфиз, Т2физ, XпарамАльфа, AполинТепл, ВлагосодВозд)
        ' располагаемая адиабатическая работа турбины 1 кг газа
        XпарамАльфа = 0.0
        АдиабРаботаКомпрессора = DI(tbфиз, ТконечАдиаб, XпарамАльфа, AполинТепл, ВлагосодВозд)
        КПДКсум = АдиабРаботаКомпрессора / ДействРаботаКомпрессора ' кпд Компрессора

        ' bКС = 0.95 - на режимах ОР, Б и УБ (М и Ф);
        ' bКС = 0.0995 + 0.905 * π(λ2) - на дроссельных режимах, где: π(λ2) = Р2/Р2*
        ' В отсутствии прямого замера статического давления Р2 его значение может быть определено следующим эмпирическим соотношением:
        ' Р2 = Р2* - dP(P2*,Soxл) ;
        ' dP(P2*, Soxn) = [ 0.0224 + 0.07487 • Р2 *- 0.0006163 • (Р2* )^2] *K(Soxл);	(4.5.1)
        ' K(Soxл) =  1.0 , при включенном охлаждении турбины (Soxл=1);
        ' 1.083 , при частично выключенном охлаждении (Soxл =0 ); 
        sigmaKC = КоэВосстПолнДавлКС(Regime, Охлаждение, Р2полнФиз)

        ' ************************      T3ТДРфиз        *******************
        If Gт_осн_кс <= 0 Then
            GтТДР = 0
            T3ТДРфиз = 0
        Else
            ' GтТДР=aGtoplTDR+bGtoplTDR*freq   используется такая формула в предыдущем подсчете
            ' GтТДР = A(17) + A(18) * freq - формула апроксмимации полиномом 1-го порядка для приведения штатного замера расхода топлива в кг в замер расхода ТДР в тех-же кг а не в литрах
            ' GтТДР = Gт_осн_кс так в КБПР

            ' Gт_осн_кс расход в л/час, GтТДР  расход кг/час
            ' РасходТопливаКамерыСгоранияКгСек = РасходТопливаКамерыСгорания / 3600.0# ' перевод л/час в кг/сек камеры сгорания
            GтТДР = Gт_осн_кс * clAir.КоэфУчитывающийПлотность(tтопл) ' ввести учет Т топлива
            Т3ПоПропСпособнТурбВД(sigmaKC, Hu, valFсаТВД, Р2полнФиз, Т2физ, GтТДР, GВоздВГорлеСА, Т3ГазаВГорлеСА, AполинТепл, C_output, ВлагосодВозд)
            T3ТДРфиз = Т3ГазаВГорлеСА
        End If

        T4огрфиз = valТ4огр + conКельвин
        Т4физ = T4 + conКельвин ' -температура газов за ТНД

        If (P4ст < Порог) Then
            P4стфиз = 0
        Else
            P4стфиз = Pбаз + P4ст
        End If

        Р4полнФиз = Pбаз + P4

        ' 4.5.5 Полное давление на выходе из КС :
        ' P3* = Pср* bКС	(4.5.5) bКС - коэф восстановления рисунок 1.3
        ' 4.6.4	Степень понижения полного давления в турбине :
        ' PiT сумм = P3* / Р4*	(4.6.4) 
        PiTсум = Р2полнФиз * sigmaKC / Р4полнФиз

        ' 3 ОПРЕДЕЛЕНИЕ СУММАРНЫХ ХАРАКТЕРИСТИК ИЗДЕЛИЯ ПО РЕЗУЛЬТАТАМ ИЗМЕРЕНИЙ
        ' Определение ряда параметров, характеризующих работу двигателя, производится на основании замеров по следующим соотношениям.
        ' 3.1	Тяга физическая :
        ' R = Rи3M*kb*kv ,	(3.1)
        ' где: кb = 1.01 - коэффициент, учитывающий угол наклона оси сопла по отношению к продольной оси изделия; 
        ' valКНаклонаОси
        ' kv - коэффициент, учитывающий входной импульс потока воздуха на входе в изделие.
        ' Для стендов ММПП «Салют» коэффициент kv определяется в зависимости от режима следующим образом ( рис. 1.1 ) [1, 6, 7, 8] :
        ' -	kv = 1.011 на режиме МГ (аРУД =12°);
        ' -	kv = 1.017 - 1.019 - на бесфорсажных режимах линейно изменяется в диапазоне
        ' {aруд = 30° - 67° } ;
        ' -	kv= 1.019 - на режиме «М» { аРУД = 67° } ;
        ' -	kv - 1.018 - 1.014 - на форсажных режимах линейно изменяется по аРУД в диапазоне aруд = 78° -116°) от МФБ,УБ до ОР.
        kv = InterpLine(arralRUD, arrKv_alRUD, Alpha_руд)
        Rфиз = valКНаклонаОси * Rизм * kv

        ' Максимал, Дросс, Мг:       
        If Regime = RegimeEngine.Максимал Then
            ' расчёт по режиму Максимал

            ' определение температуры газа перед ТВД в кр сеч СА ТВД и расх воздуха через СА ТВД. 
            ' Для ПФ и Ф вышеупомянутые параметры, а также m(степ двухконт) вычисляются по формулам (4.6.5) и (4.6.6)

            ' 4.5 Параметры газа перед турбиной (ТВД). 
            ' 4.5.1 Определение температуры газа Т*г перед турбиной высокого давления (ТВД) в критическом сечении соплового аппарата (СА) и расхода воздуха через СА ТВД производится с использованием следующих величин:
            ' •	Р *2ср - среднего значения полного давления в сечении "2" (ф-ла 4.3.2);
            ' •	T2*	- температуры воздуха на входе в КС (измеряется); 
            ' •	GTKC	- расхода топлива через КС (измеряется);
            '    FСА твд	- геометрической площади критического сечения СА ТВД (измеряется);
            ' •	Нu	- фактической теплотворной способности топлива;
            ' •	H кс	- коэффициента полноты сгорания в КС;
            ' •	bКС	- коэффициента восстановления полного давления в КС (рис. 1.3 ), где

            ' По замерам FСА твд	рассчитывается значение фактической пропускной способности СА ТВД :
            ' GпрТВД = GпроТВД  * (FСАТВД / FСАТВДо )

            ' Где GпроТВД  =118,0  при тарированном значении САТВДо= 305 [см2]
            ' Температура газа Т*ГСА в горле соплового аппарата турбины высокого давления (СА ТВД) и расход газа через СА ТВД GГca находятся из решения системы уравнений: 

            ' (4.5.3)
            ' 4.5.3 Пропускная способность СА турбины ВД 
            ' GГСА= (GпрТВД * Р*ГСА) / SQR(Т*ГСА)

            ' (4.5.4)
            ' 4.5.4 Уравнение баланса тепла в камере сгорания (до горла СА) 
            ' GГСА * IГ(Т*ГСА) = (GГСА – GТКС)*IB(T*K) + GТКС*(Hu*ηКС + IT(To))

            ' где:
            ' Р*ГСА  = P*2 *  sigmaКС - среднее значение полного давления в горле СА ТВД, [кгс/см2]; 
            ' Т*ГСА - средняя температура газа в горле СА ТВД;
            ' GГСА - расход газа в горле СА ТВД;
            ' GТКС - расход топлива в основной камере сгорания (по данным измерений) кг/с; 
            ' FСАТВД - геометрическая площадь критического сечения СА ТВД (измеряется);
            ' IГ(Т*ГСА) = f(Т*ГСА , qТСА) - энтальпия газа в горле СА, [ кДж/кг];
            ' IB(T*K) = f(T*K) энтальпия воздуха выходе из КВД, [ кДж/кг];
            ' GВСА = (GГСА  - GT кс) - расход воздуха в горле СА;
            ' qТСА  =  GT кс / GВСА - относительный расход топлива в горле СА;
            ' ηКС - коэффициент полноты сгорания топлива в КС;
            ' Hu - теплотворная способность топлива (по данным хим. лаборатории).
            ' IT(To) - начальная энтальпия топлива при температуре То=293.15 К.
            ' ηКС = 0,99 в соответствии с «Инструкцией на приведение параметров и определение температуры газа при стендовых испытаниях изделия 99». 99.57 ИН [1].
            ' Т.о. система (4.5.1) - (4.5.4) является замкнутой и позволяет определить параметры в критическом сечении СА ТВД: GВСА, qT и Т*Г СА с учетом фактического значения Ни.
            ' Термодинамические параметры входящие в систему уравнений (4.5.1) * (4.5.4) определяются на оснований [11].
            ' В отсутствие фактических оценок параметра Huб в качестве расчётных данных допускается применение значения по ГОСТ 10227-86, что для обычных марок топлива для реактивных деигателей составляет величину Ни = 10300 ккал/кг (низшая теплота сгорания) [1, 10]. 

            Т3ПоПропСпособнТурбВД(sigmaKC, Hu, valFсаТВД, Р2полнФиз, Т2физ, Gтопл, GВоздВГорлеСА, Т3ГазаВГорлеСА, AполинТепл, C_output, ВлагосодВозд)

            If (Gтопл < Порог) Then
                Gтопл = 0.0
                T3физMemory = 0.0
                альф_физ = 0.0
                CR = 0.0
            Else
                ' 3.2	Физический удельный расход топлива :
                ' CR= GT/R	3.2)  
                CR = Gтопл / Rфиз
                ' 3.3	Физический суммарный коэффициент избытка воздуха :
                ' (3.3)
                ' aсум =3600* Gbсум / (L0 • GT)
                ' где значение стехиометрического коэффициента Lo принято равным Lo =14.95, аналогично штатному изделию 99 [1]. Уточняется по фактическому сорту используемого топлива.
                альф_физ = 3600 * GвоздСумФизВ / Gтопл / СтехиометрКоэфLo
            End If

            ' 4.6 Оценка степени двухконтурности и параметров турбины .
            ' 4.6.1	Суммарный расход воздуха через внутренний контур ( на входе в KBД) GIВсум . Вычисляется по GBСА с учетом относительного расхода воздуха через СА ТВД:
            ' GIВсум = GBСА  / vJCA;	(4.6.1)
            ' где vJCA - коэффициент относительного расхода воздуха через СА ТВД. В соответствии с расчетными данными и Инструкцией [1] :
            ' vJCA = 0.88 - при включенном охлаждении и П2> 90%;
            ' vJCA = 0.89 - при включенном охлаждении и П2< 80%;
            ' vJCA = 0.94 - при частично выключенном охлаждении . (4.6.1)	 
            ' 4.6.2	Суммарный расход воздуха через контур II. Вычисляется как разность:
            ' GIIВсум = GВсум - GIВсум	(4.6.2)
            ' 4.6.3	Степень двухконтурности. По определению, с учетом (4.6.1), (4.6.2):
            ' m = GIIВсум / GIВсум	(4.6.3)

            ' охлаждение СА ТВД
            n2пр = N2f * КоэПриведенияTбокса '  Math.Sqrt(288.15 / tb)
            ' Охлаждение=0(отсутств сигнала), Охлаждение=1(частич охл), Охлаждение=5(полное охл)
            ' -njuOHL- (НЮ)коэфф отн расх возд на охл СА ТВД
            If (Охлаждение = 0) Then
                njuОхл = 0.94
            Else ' Охлаждение >= 1
                If N2f > 91.5 Then
                    njuОхл = 0.88
                ElseIf n2пр < 80.0 Then
                    njuОхл = 0.89
                End If
            End If

            GвоздВнКфиз = GВоздВГорлеСА / njuОхл ' расх возд через 1-й конт  GIВсум = GBСА  / vJCA
            ' GIIВсум = GвоздФизВ - GвоздФиз ' GIIВсум = GВсум - GIВсум
            ' If _Режим = РежимEnum.Максимал Then
            ' _m = GIIВсум / GвоздФиз' или сокращая без использования GIIВсум
            _m = (GвоздСумФизВ - GвоздВнКфиз) / GвоздВнКфиз
            T3физ = Т3ГазаВГорлеСА

            If isMemoryValue Then
                ' Запомнить переменные для последующего расчета форсажа
                u447Memory = U447физ ' запоминание значения u447 для М режима
                PiTMemory = PiTсум ' запоминание значения PiT для М режима
                T3физMemory = T3физ
                mФизMemory = _m
                GвоздВнКфизMemory = GвоздВнКфиз
                isDataForMaximumCollected = True
                isMemoryValue = False
            End If
        ElseIf Regime = RegimeEngine.Форсаж AndAlso isDataForMaximumCollected Then
            ' расчёт по режиму Форсаж
            ' Примечание:
            ' 1. На режимах "Ф" в отсутствие прямых замеров расхода топлива в основной КС определение температуры газа перед СА ТВД и степени двухконтурности производится путем сравнения с максимальным режимом по следующим эмпирическим формулам [1]:
            '  T*ГФ = T* Гмах + 1.4*(U447 - U447max) ; 	(4.6.5)
            ' Mф = mmax + 0.23 * (PiT сумм ф - PiT сумм max)	(4-6.6)
            ' 2. С целью накопления статистики и уточнения метода определения Т*ГФ факультативно рассчитывать Т*ГФ по значению  GT ОСН КС 
            ' 2.1	GT ОСН КС = GT СУМ – GT Ф	(4.6.7) где GT СУМ -суммарный расход топлива на режиме"Ф"?
            ' GT Ф является суммарным расходом топлива по 5 коллекторам 
            ' (в форсажной камере наверно). 
            ' 2.2	GT ОСН КС   (наверно имеется в виду GT Ф) определяется по показанию датчика ТДР-14.

            GвоздВнКфиз = GвоздВнКфизMemory
            _m = mФизMemory + 0.23 * (PiTсум - PiTMemory)

            If (Gтопл < Порог) Then
                Gтопл = 0.0
                T3физMemory = 0.0
                T3физ = 0.0
                альф_физ = 0.0
                CR = 0.0
            Else
                T3физ = T3физMemory + 1.4 * (U447физ - u447Memory)
                альф_физ = 3600 * GвоздСумФизВ / Gтопл / СтехиометрКоэфLo
                CR = Gтопл / Rфиз
            End If
        End If

        If (Rфиз < Порог) Then CR = 0.0
    End Sub

    Private Sub ВычислитьПриведенные()
        ' 5. ПРИВЕДЕНИЕ К САУ ОСНОВНЫХ ПАРАМЕТРОВ ИЗДЕЛИЯ НА ДРОССЕЛЬНЫХ РЕЖИМАХ
        ' 5.1 Приведение к САУ основных параметров изделия при FKp= const производится к параметрам на входе в изделие (сечение «В» РМК) по следующим формулам:
        ' T*В  = T*вх
        ' n1пр = n1* SQR(288.15/ T*В  )
        ' n2пр = n2* SQR(288.15/ T*В  )
        ' Rпр = R*1.0332 / P*В
        ' GTпр = GT * (1.0332 / P*В) * SQR(288.15/ T*В  )
        ' T*Гпр = T*Г *288.15/ T*В  
        ' U447 = (U447 + 273.15) * 288.15/ T*В  - 273.15      °C
        ' Для температуры и давления в проточной части:
        ' T*i\пр = ( t*i  + 273.15) * 288.15/ T*В  
        ' Pi\пр = Pi * 1.0332 / P*В
        ' P*i\пр = P*i * 1.0332 / P*В
        ' Примечание. Приведение параметра п2пр ко входу в двигатель не связано с критерием подобия, а используется только для регулирования НА КВД.
        ' 5.3 Приведение параметров КВД ко входу в КВД производится по спецуказанию при поиске неисправностей или анализе параметров [1]:
        ' n2пр\6 = n2* SQR(288.15/ T*6  )*Kn2
        ' GВквд\пр\6 = Gв сум * (1.0332 / P*6) * SQR(T*6  / 288.15) нет ли под корнем ошибки

        ' где G' Bсум рассчитывается по ф-ле (4.6.1).
        ' Kn2 - переходный коэффициент для приведенной частоты РВД [1]
        ' Kn2 = N2max/N1max = 13300/10200 = 1.3039
        Dim Коэфtb As Double = con288 / tbфиз

        n1пр = N1f * КоэПриведенияTбокса
        n2пр = N2f * КоэПриведенияTбокса

        If T3ТДРфиз <= Порог Then
            T3ТДРпр = 0
        Else
            T3ТДРпр = T3ТДРфиз * Коэфtb
        End If

        Rпр = Rфиз * КоэПривP_вх_КНД
        Gтпр = Gтопл * КоэПривP_вх_КНД * КоэПриведенияTбокса

        CRпр = Gтпр / Rпр
        GвоздСумпр = GвоздСумФизВ * КоэПривP_вх_КНД / КоэПриведенияTбокса
        альфаSпр = 3600 * GвоздСумпр / Gтпр / СтехиометрКоэфLo

        T2пр = Т2физ * Коэфtb
        T3пр = T3физ * Коэфtb
        P4стПр = Т4физ * Коэфtb
        T6пр = T6физ * Коэфtb
        u447пр = con288 * U447физ / tbфиз

        ' 4.2.3 Показатель устойчивости КНД ( с учетом соотношения (1.10), Приложение I):
        ' Wкнд=Piкнд / Gbсум_пр	(4.2.4)
        УстКНДпр = Pi_КНД / GвоздСумпр
        P2полнПр = Р2полнФиз * КоэПривP_вх_КНД
        P4полнПр = Р4полнФиз * КоэПривP_вх_КНД
        P6полнПр = Р6полнФиз * КоэПривP_вх_КНД
        P2стПр = P2стфиз * КоэПривP_вх_КНД
        P4пр = P4стфиз * КоэПривP_вх_КНД
        P6стПр = P6стфиз * КоэПривP_вх_КНД

        If (P6ст = 0) Then P6стПр = 0

        If (T6физ <= 0) Then
            n2КВДпр = 0
            GКВДпр = 0
            УстКВДпр = 0
        Else
            n2КВДпр = N2f * ПерКоэфПрЧастотыРВД * Math.Sqrt(con288 / T6физ) ' 1.234
            GКВДпр = GвоздВнКфиз * con_1_033 / Р6полнФиз * Math.Sqrt(T6физ / con288)

            ' 4.3.3 Показатель устойчивости КВД :
            ' Wквд = Piквд  / Gв_квд\пр\6 	(4.3.4)
            ' где: GB КВД\пр\6 рассчитывается по ф-ле (5.3-2).
            ' Примечание. Значения степеней сжатия Piкнд и Piквд,  коэффициентов полезного действия Hкнд , Hквд и показателей устойчивости Wкнд, Wквд, определенные по одно-точечным замерам Р6* и t*6 , - могут быть использованы лишь для ориентировочной оценки.
            УстКВДпр = РiКВД / (GвоздВнКфиз * con_1_033 / Р6полнФиз * Math.Sqrt(T6физ / con288))
        End If

        If КПД_КНД < 0 Then КПД_КНД = 0
        If КПД_КВД < 0 Then КПД_КВД = 0
    End Sub

    Private Sub ВычислитьПересчитанные()
        ' ПЕРЕСЧЕТ К САУ ОСНОВНЫХ ПАРАМЕТРОВ ИЗДЕЛИЯ
        ' НА МАКСИМАЛЬНЫХ И ФОРСИРОВАННЫХ РЕЖИМАХ
        ' 6.1	Пересчет основных параметров изделия к САУ на максимальных и форсированных режимах производится с использованием коэффициентов пересчета.
        ' 6.2	Коэффициенты пересчета К1 представляют собой относительные величины параметров, полученные по результатам математического моделирования термогазодинамических параметров изделия 99М2 в диапазоне температур воздуха на входе от -45 °С до +45 °С при конкретно выполненной настройке, поддержании заданных законов регулирования и с учетом экспериментальных данных по среднестатистической оценке характеристик составных элементов изделия, (см. рис. 2-26).
        ' 6.3	Коэффициенты пересчета К2 представляют собой относительные величины параметров, полученные по результатам математического моделирования термогазодинамических параметров изделия 99М2 в диапазоне температур воздуха на входе от -45 °С до +45 °С при относительной влажности воздуха Fi, изменяющейся в диапазоне от 0 до 100%. Для приведения параметров к условиям стандартной (нулевой) влажности окружающего воздуха рассчитан-ные значения необходимо дополнительно умножить на коэффициент пересчёта по влажности, (см. рис. 27-49).
        ' 6.4	Пересчет параметров изделия к САУ на максимальных (ОРМ, БМ, УБМ) и форсированных режимах (ОР, ПФ(Б) и УБФ) производится по формулам таблицы 6.

        n1пер = kn1 * kn1fi * N1f
        n2пер = kn2 * kn2fi * N2f
        Rпер = kR * kRfi * Rфиз * КоэПривP_вх_КНД
        CRпер = kCR * kCRfi * CR
        Gт_пер = Gтопл * kGT0 * kGT0fi * КоэПривP_вх_КНД
        T3пер = kT3 * kT3fi * T3физ
        T3ТДРпер = kT3 * kT3fi * T3ТДРфиз
        T4пер = kT4 * kT4fi * Т4физ
        GвоздСумпер = kGB * kGBfi * GвоздСумФизВ * КоэПривP_вх_КНД
        mпер = km * kmfi * _m
        PiКНДпер = kPiB * kPiBfi * Pi_КНД
        PiКsпер = kPiKS * kPiKSfi * PiKсум

        ' Select Case _Режим
        '    Case РежимEnum.Максимал
        '        alfaSпер = 0
        '        PiKVDпер = 0
        '        GTFпер = 0.0
        '    Case РежимEnum.Форсаж
        альфа_пер = kaS * альф_физ
        PiКВДпер = kPiKVD * kPiKVDfi * РiКВД
        GTфор_пер = kGTF * kGTFfi * Gтопл
        'End Select
    End Sub

    Private Sub ЗапроситьНастройкиДляСтенда()
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim pathDBase As String = gMainFomMdiParent.OwnCatalogue & ".mdb" ' & "\" & gvarfrmMain.Tag
        Dim strSQL As String = "SELECT * FROM(НомерБокса) WHERE (((НомерБокса.НомерБокса)= '" & valНомерБокса & "'));"

        cn = New OleDbConnection(BuildCnnStr(PROVIDER_JET, pathDBase))
        cn.Open()
        cmd = cn.CreateCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandText = strSQL

        Try
            Dim rdr As OleDbDataReader = cmd.ExecuteReader
            If rdr.Read Then
                K1a0 = CDbl(rdr("K1a0"))
                K1a1 = CDbl(rdr("K1a1"))
                K1a2 = CDbl(rdr("K1a2"))
                K1a3 = CDbl(rdr("K1a3"))
                K2a0 = CDbl(rdr("K2a0"))
                K2a1 = CDbl(rdr("K2a1"))
                K2a2 = CDbl(rdr("K2a2"))
                K2a3 = CDbl(rdr("K2a3"))
                DM = CDbl(rdr("DM"))
                DB = CDbl(rdr("DB"))
                БазовоеДавлениеВКабине = CBool(rdr("БазовоеДавлениеВКабине"))
                Кпотери = CDbl(rdr("Кпотери"))
            End If
            rdr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString, $"Процедура {NameOf(ЗапроситьНастройкиДляСтенда)}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            If (cn.State = ConnectionState.Open) Then
                cn.Close()
            End If
        End Try
    End Sub

    Private Sub ЗапроситьНастройкиДляТипТоплива()
        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim pathDBase As String = gMainFomMdiParent.OwnCatalogue & ".mdb" '& "\" & gvarfrmMain.Tag
        Dim strSQL As String = "SELECT * FROM(ТипТоплива) WHERE (((ТипТоплива.ТипТоплива)= '" & valТипТоплива & "'));"

        cn = New OleDbConnection(BuildCnnStr(PROVIDER_JET, pathDBase))
        cn.Open()
        cmd = cn.CreateCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandText = strSQL

        Try
            Dim rdr As OleDbDataReader = cmd.ExecuteReader
            If rdr.Read Then
                Hu = CDbl(rdr("Hu"))
            End If
            rdr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString, $"Процедура {NameOf(ЗапроситьНастройкиДляТипТоплива)}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            If (cn.State = ConnectionState.Open) Then
                cn.Close()
            End If
        End Try
    End Sub

    Private Function ЗапроситьПараметрыДляРасчетаФорсажа(ByVal НомерКТМаксимала As String) As Boolean
        ' известно valNктМаксимала с которым была записана данная точка с режимом Форсаж
        ' во время пересчета проверяется наличие в базе КТ с данным номером  в данном запуске и чтобы эта КТ была с режимом Максимал
        ' если все в порядке считываются нужные параметры и производится расчет
        ' индекс запуска (keyНомерЗапуска) где искать КТ с данным номером в базе должен быть такой же как и для данной пересчитываемой КТ режима Максимал
        ' иначе в расчет константы99999 или выдать ошибку

        Dim cn As OleDbConnection
        Dim cmd As OleDbCommand
        Dim pathDBase As String = gMainFomMdiParent.OwnCatalogue & ".mdb" '& "\" & gvarfrmMain.Tag
        Dim rdr As OleDbDataReader
        Dim strSQL As String

        Dim success As Boolean = False ' Параметры Для Расчета Форсажа Считаны
        Dim keyНомерЗапуска, keyНомерКТ As Integer
        Dim customerValue As String = Nothing ' Значение Пользователя

        cn = New OleDbConnection(BuildCnnStr(PROVIDER_JET, pathDBase))
        cn.Open()
        cmd = cn.CreateCommand
        cmd.CommandType = CommandType.Text
        ' узнаем keyНомерЗапуска
        strSQL = "SELECT [5НомерЗапуска].keyНомерЗапуска " &
        "FROM 5НомерЗапуска RIGHT JOIN 6НомерКТ ON [5НомерЗапуска].keyНомерЗапуска = [6НомерКТ].keyНомерЗапуска " &
        "WHERE ((([6НомерКТ].keyНомерКТ)= " & Manager.KeyNumberKT.ToString & " ));"
        cmd.CommandText = strSQL

        Try
            rdr = cmd.ExecuteReader
            If rdr.Read Then keyНомерЗапуска = CInt(rdr("keyНомерЗапуска"))
            rdr.Close()

            If keyНомерЗапуска <> 0 Then
                ' по keyНомерЗапуска узнаем keyНомерКТ для контрола с именем "НомерКТ" и значением = valNктМаксимала
                strSQL = "SELECT [6НомерКТ].keyНомерКТ " &
                "FROM 5НомерЗапуска RIGHT JOIN (КонтролыДляНомерКТ RIGHT JOIN (6НомерКТ RIGHT JOIN ЗначенияКонтроловДляНомерКТ ON [6НомерКТ].keyНомерКТ = ЗначенияКонтроловДляНомерКТ.keyУровень) ON КонтролыДляНомерКТ.keyКонтролДляУровня = ЗначенияКонтроловДляНомерКТ.keyКонтролДляУровня) ON [5НомерЗапуска].keyНомерЗапуска = [6НомерКТ].keyНомерЗапуска " &
                "WHERE ((([5НомерЗапуска].keyНомерЗапуска)= " & keyНомерЗапуска & " ) AND ((КонтролыДляНомерКТ.Name)='НомерКТ') AND ((ЗначенияКонтроловДляНомерКТ.ЗначениеПользователя)= '" & НомерКТМаксимала & "' ));"

                cmd.CommandText = strSQL
                rdr = cmd.ExecuteReader
                If rdr.Read Then keyНомерКТ = CInt(rdr("keyНомерКТ"))
                rdr.Close()

                If keyНомерКТ <> 0 Then
                    ' keyНомерКТ это keyУровень для таблицы ЗначенияКонтроловДляНомерКТ по нему узнаем ЗначениеПользователя для контрола с именем ="Режим" и он должен быть равен "М"
                    strSQL = "SELECT ЗначенияКонтроловДляНомерКТ.ЗначениеПользователя " &
                    "FROM КонтролыДляНомерКТ RIGHT JOIN ЗначенияКонтроловДляНомерКТ ON КонтролыДляНомерКТ.keyКонтролДляУровня = ЗначенияКонтроловДляНомерКТ.keyКонтролДляУровня " &
                    "WHERE (((КонтролыДляНомерКТ.Name)='Режим') AND ((ЗначенияКонтроловДляНомерКТ.keyУровень)= " & keyНомерКТ & " ));"

                    cmd.CommandText = strSQL
                    rdr = cmd.ExecuteReader
                    If rdr.Read Then customerValue = CStr(rdr("ЗначениеПользователя"))
                    rdr.Close()

                    If customerValue = "М" Then
                        ' после всех проверок выполнить перекрёстный запрос на выборку данной КТ максимала и извлечения необходимых параметров
                        strSQL = "TRANSFORM First([7ЗначенияПараметровКТ].Значение) AS [First-Значение] " &
                        "SELECT [6НомерКТ].НомерКТ " &
                        "FROM 6НомерКТ RIGHT JOIN (РасчетныеПараметры INNER JOIN 7ЗначенияПараметровКТ ON РасчетныеПараметры.ИмяПараметра = [7ЗначенияПараметровКТ].ИмяПараметра) ON [6НомерКТ].keyНомерКТ = [7ЗначенияПараметровКТ].keyНомерКТ " &
                        "WHERE ((([7ЗначенияПараметровКТ].Значение)<>0) AND (([6НомерКТ].keyНомерКТ)= " & keyНомерКТ.ToString & " )) " &
                        "GROUP BY [6НомерКТ].НомерКТ " &
                        "PIVOT РасчетныеПараметры.ИмяПараметра;"

                        cmd.CommandText = strSQL
                        rdr = cmd.ExecuteReader

                        If rdr.Read Then
                            u447Memory = CDbl(rdr(callU447физ)) ' U447физ
                            PiTMemory = CDbl(rdr(callPiTсум)) ' PiTсум
                            T3физMemory = CDbl(rdr(callT3физ)) ' T3физ
                            mФизMemory = CDbl(rdr(callm)) ' _m
                            GвоздВнКфизMemory = CDbl(rdr(callGвоздВнКфиз)) ' GвоздВнКфиз
                            success = True
                        End If

                        rdr.Close()
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString, $"Процедура {NameOf(ЗапроситьПараметрыДляРасчетаФорсажа)}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            If (cn.State = ConnectionState.Open) Then
                cn.Close()
            End If
        End Try
        If Not success Then
            MessageBox.Show("КТ с номером " & НомерКТМаксимала & " режима <Максимал>" & vbCrLf &
                            "с которой производится пересчет отсутствует!" & vbCrLf &
                            "Пересчет не будет произведен.", "Процедура ЗапроситьПараметрыДляРасчетаФорсажа", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Return success
    End Function

    Private Sub ПараметрыРасходомерногоКоллектора(ByVal Tb As Double, ByVal РстатМернСечМ As Double, ByVal РполнМернСечМ As Double, ByRef GBфизB As Double, ByRef P_вх_КНД As Double, ByVal k As Double)
        Dim Kminus1dKplus1, Kminus1_delK, One_delKminus1, Kplus1_del2, Kplus1_del2_Step_One_delKminus1 As Double
        Dim k1, k2, LamM, LamB, PiM, PMcp, PiMcp, qLamMcp, PBст, yLamB, PiLb, qLb, Pbторм, SсечВ As Double ', GвозПрB
        ' Приложение 1
        ' Алгоритм расчета параметров в расходомерном коллекторе (РМК). 
        ' Статическое и полное давление воздуха в мерном сечении «М»:
        ' Рм = (В0 –dРM)/735.6, [кг/см2];		(1.1)
        ' Р*м = (В0 –dР*M)/735.6, [кг/см2];		(1.1)
        ' Здесь: dРм, dР*м [мм рт.ст.] - значения избыточного статического и полного давлений,
        ' осредненные по нескольким замерам;
        ' Во [мм рт.ст.] - базовое давление для дифференциальных датчиков давления (Во = Ва или Во= Рб - в зависимости от способа установки датчиков).
        ' Газодинамическая функция в сечении «М»:
        ' πM = π(λM ) =  Рм / Р*м 
        ' Для РМК ММПП «Салют» (диаметром DM =905 мм, DM =907 мм или DM =924мм) коэффициент поля полного давления в сечении «М» :
        ' кi = -0.05019 + 3.3523 * πM -3.5814* πM 2+1.2794 * πM3	(1.3)
        ' ГДФ ПИ в мерн сечении М:
        PiM = РстатМернСечМ / РполнМернСечМ '  -(1.2) -нумерация формул см стр 13 ПРИЛОЖЕНИЯ 1
        ' для входного устр-ва коэфф k1 поля полного давления в сечении "М": 
        ' k1 = -0.05019 + 3.3523 * PiM - 3.5814 * PiM * PiM + 1.2794 * PiM * PiM * PiM   ' - (1.3)
        ' k1 = -0.050191 + 3.352285 * PiM  -3.581382 * PiM * PiM + 1.27943 * PiM * PiM * PiM   ' - (1.3)
        k1 = K1a0 + K1a1 * PiM + K1a2 * PiM * PiM + K1a3 * PiM * PiM * PiM

        ' НомерВходногоУстройства		К1ПолнДавл	К2ПолнДавл	К3ПолнДавл	К4ПолнДавл	К1Потерь	К2Потерь	К3Потерь	К4Потерь	Дм	Дв
        ' 6464-4203		                -0.050191	3.352285	-3.581382	1.27943	    -0.464944	4.628497	-4.947336	1.784405	905	905

        ' Среднее значение полного давления в сечении «М»:
        ' Р*м|ср = Р*м * k1 ; πM|ср = Рм / Р*м|ср	(1-4)
        ' среднее значение полного давления в сечении "М":
        PMcp = РполнМернСечМ * k1 ' (1.4)
        ' среднее значение ГДФ ПИ в сечении "М":
        PiMcp = РстатМернСечМ / PMcp ' (1.4)

        ' Исходя из соотношений (1.2) - (1.5) по значению газодинамической функции πM|ср приведенная плотность потока воздуха в сечении q(λ)м\ср определяется на основании решения системы уравнений :
        ' Pi(Lam) = (1- (k-1)/(k+1)*Lam^2)^(k/(k-1))		
        ' k=1.4 (для воздуха при t*вх < 50 град С
        ' q(Lam) = (Lam * ((k+1)/2)^(1/(k-1))) * (1- ((k-1)/(k+1)) * Lam^2  )^(1/(k-1))
        ' πM|ср выводится из q(λ)м\ср

        Kminus1dKplus1 = (k - 1) / (k + 1)
        Kminus1_delK = (k - 1) / k
        Kplus1_del2 = (k + 1) / 2
        One_delKminus1 = 1 / (k - 1)
        Kplus1_del2_Step_One_delKminus1 = Kplus1_del2 ^ One_delKminus1

        LamM = Math.Sqrt((1 - PiMcp ^ Kminus1_delK) / Kminus1dKplus1) ' аналитич решение 1-го уравнения (1.5)
        qLamMcp = LamM * (Kplus1_del2 * (1 - Kminus1dKplus1 * LamM * LamM)) ^ One_delKminus1 ' второе уравнение (1.5)

        ' Коэффициент пересчета статического давления от сечения «М» к сечению «В» к2=Рв/Рм для входных каналов длиной Lк = 1902 мм (с диаметрами DM =905 мм, DB =905 мм) или (DM =924 мм, DB =924 мм) определяется в функции параметра πM:
        ' к2 = 0.2803+2.161 * πM - 2.215 * πM2 + 0.7738 * πM3 ; 
        ' ак=1; РB =РМ * к2	(1.6)
        ' Для входного канала длиной Lк = 1902 мм (с диаметрами DM =907 мм, DB =905 мм) к2=Рв/Рм определяется в функции параметра πM:
        ' к2 = 0.5359 + 1.285 * πM - 1.2048 * πM 2 + 0. 3843 * πM3  
        ' ак=1.015; РB =РМ * к2	(1.6а)
        ' коэффициент k2 пересчёта статич давл от сеч "М" к сеч "В":
        k2 = K2a0 + K2a1 * PiM + K2a2 * PiM * PiM + K2a3 * PiM * PiM * PiM
        PBст = РстатМернСечМ * k2 '  -стат давл в сечении В

        ' Значение термогазодинамической функции у (Lam)в  в сечении «В»:
        ' у (Lam)в  =(q(λ)м\ср * Р*м|ср) / Рв * Fм/FВ
        ' Здесь: FM, Fв - значения площадей сечений «М» и «В», соответственно. 
        yLamB = qLamMcp * PMcp / PBст * DM * DM / DB / DB ' -(1.7)  '  Учёт наличия конусности

        ' По значению у (Lam)в  определяется значение функции Pi(Lam)в и полное давление PB*:
        ' у (Lam)в  => Pi(Lam)в: Pi(Lam)в = q(Lam)B  / у (Lam)в ; PB*= PB / Pi(Lam)в	(1.8)

        LamB = (Math.Sqrt(Kplus1_del2_Step_One_delKminus1 * Kplus1_del2_Step_One_delKminus1 + 4 * Kminus1dKplus1 * yLamB * yLamB) - Kplus1_del2_Step_One_delKminus1) / 2 / Kminus1dKplus1 / yLamB ' -аналит решение уравнения yLamB(LamB)=q(LamB)/Pi(LamB) относительно LamB,
        ' где LamB определяется как корень квадратного уравнения после преобразования

        PiLb = (1 - Kminus1dKplus1 * LamB * LamB) ^ (1 / Kminus1_delK) ' первое уравнение (1.5)
        qLb = LamB * (Kplus1_del2 * (1 - Kminus1dKplus1 * LamB * LamB)) ^ One_delKminus1 ' второе уравнение (1.5)

        ' Значение суммарного расхода воздуха на входе в изделие GBсум определяется:
        ' GBсум =	((mB *  PB* ) / sqr(T*B)) * q(Lam)B  * FВ  * 10000 * ак;    Tв = Tб ;	(1.9)
        ' где тв = 0.3964 (для воздуха при t*вх < 50 °С), множитель ак выбирается по соотношениям (1.6), (1.6а).
        Pbторм = PBст / PiLb ' -(1.8)
        SсечВ = Math.PI * DB * DB / 4.0
        ' Приведенное ко входу в двигатель (сечение «B») значение расхода воздуха:
        ' GBсумпр = GBсум *  (1.0332 / P*B) * SQR(T*B  / 288.15)
        GBфизB = 0.3964 * Pbторм / Math.Sqrt(Tb) * qLb * SсечВ * 10000.0 * Кпотери ' -(1.9)

        ' GвозПрB = GBфизB * con_1_033 / Pbторм * Math.Sqrt(Tb / con288)

        P_вх_КНД = PBст / PiLb '  P_вх_КНД это полное давление Р*_В
        КоэПривP_вх_КНД = con_1_033 / P_вх_КНД
    End Sub

    Private Function КоэВосстПолнДавлКС(ByVal _Режим As RegimeEngine, ByVal Охлаждение As Integer, ByVal P2b As Double) As Double
        Dim KsOHL, sigKC, dP, PP2 As Double

        If Охлаждение = 0 Then
            KsOHL = 1.083
        Else
            KsOHL = 1.0
        End If

        If _Режим = RegimeEngine.Максимал OrElse _Режим = RegimeEngine.Форсаж Then
            sigKC = 0.95
        Else ' Крейсерский(дроссельный)
            dP = (0.0224 + 0.07487 * P2b - 0.0006163 * P2b * P2b) * KsOHL
            PP2 = P2b - dP
            sigKC = 0.0995 + 0.905 * PP2 / P2b
        End If

        Return sigKC
    End Function

    Private Function Kadiabat(ByVal Tgrad As Double) As Double
        Dim K As Double = 1.4

        If (Tgrad <= 50) Then
            K = 1.4
        ElseIf (Tgrad > 50 AndAlso Tgrad < 500) Then
            K = 1.4035 - 5.05 * 0.00001 * Tgrad - 8.5 * 0.00000001 * Tgrad * Tgrad    '  (4.2.3)
        End If

        Return K
    End Function

    ''' <summary>
    ''' Процедура CPS предназначена для определения теплоемкости, газовой постоянной и показателя адиабаты по заданной температуре Т и коэффициенту избытка воздуха α. Попутно определяется стехиометрический коэффициент L0 .
    ''' Заголовок процедуры: 'PROCEDURE' CPS(C, Т, X, А).
    ''' C(1) - Cp    C(2) -R    C(4) - k     C(6) - L0
    ''' </summary>
    ''' <param name="C"></param>
    ''' <param name="T"></param>
    ''' <param name="X"></param>
    ''' <param name="A"></param>
    ''' <param name="ВлагоСодерж"></param>
    ''' <remarks></remarks>
    Private Sub CPS(ByRef C() As Double,
                       ByVal T As Double,
                       ByVal X As Double,
                       ByVal A(,) As Double,
                       ByVal ВлагоСодерж As Double)

        ' Параметры Т и X занесены в список значений.
        ' Назначение параметров:
        ' T - значение температуры;
        ' x= 1/α , применение параметра x удобнее, чем α, так как для чистого воздуха α= ∞ (X = О).
        ' Содержимое массивов А и С описано ранее; параметры Т , X и А являются входными, массив С - результат.

        Dim P(10) As Double
        Dim D, QT, AU, BU As Double
        Dim M, N As Integer
        ' почти всегда X = 0

        '  	C(1) - Cp    C(2) -R    C(4) - k     C(6) - L0    
        C(6) = 28.966 / 0.2095 * (A(1, 11) / 12.01 + A(2, 11) / 4.032 - A(3, 11) / 32.0)
        D = ВлагоСодерж
        QT = (1.0 - D) * X / C(6)
        M = 6
        N = M + 1

        If (QT < 0.00001 AndAlso D < 0.00001) Then
            C(3) = 29.27
            For I As Integer = 1 To N
                P(I) = A(1, I)
            Next
        Else
            If QT < 0.00001 Then
                C(3) = 29.27 * (1.0 + 0.60779 * D)
                For I As Integer = 1 To N
                    P(I) = A(1, I) * (1.0 - D) + A(4, I) * D
                Next
            End If

            If A(2, 11) < 0.999 Then
                If (Math.Abs(A(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(A(2, 11) - 0.15) < 0.001 AndAlso D < 0.000001) Then
                    C(3) = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
                    For I As Integer = 1 To N
                        P(I) = (A(2, I) * QT + A(1, I)) / (1.0 + QT)
                    Next
                Else
                    C(3) = 29.27 * (1.0 + 28.966 * (A(2, 11) / 4.032 + A(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
                    For I As Integer = 1 To N
                        AU = (44.01 * A(3, I) - 32.0 * A(5, I)) / 12.01
                        BU = (9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008
                        P(I) = ((A(1, 11) * AU + A(2, 11) * BU + A(3, 11) * A(5, I)) * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                    Next
                End If
            Else
                If X < 1.0 Then
                    C(3) = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
                    For I As Integer = 1 To N
                        P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                    Next
                Else
                    C(3) = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
                    For I As Integer = 1 To N
                        P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * (1.0 - D) / C(6) + D * A(4, I) + (1.0 - D) * A(1, I) + (QT - (1.0 - D) / C(6)) * A(6, I)) / (1.0 + QT)
                    Next
                End If
            End If
        End If

        T = T * 0.001
        C(1) = P(M + 1)

        For I As Integer = 1 To M
            C(1) = C(1) * T + P(M + 1 - I)
        Next

        C(4) = C(1) / (C(1) - C(3) / 426.9)
        T = T * 1000
    End Sub

    ''' <summary>
    ''' Процедура - функции DI предназначена для определения приращения энтальпии от температуры Т0 до температуры Т при заданном α АЛЬФА(Х=1/АЛЬФА). Приращение энтальпии вычисляется в единицах измерения ккал/кг.
    ''' Заголовок процедуры: 'PROCEDURE' DI (ТО, Т, X, А).
    ''' </summary>
    ''' <param name="T0"></param>
    ''' <param name="T"></param>
    ''' <param name="X"></param>
    ''' <param name="A"></param>
    ''' <param name="ВлагоСодерж"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DI(ByVal T0 As Double,
                           ByVal T As Double,
                           ByVal X As Double,
                           ByVal A(,) As Double,
                           ByVal ВлагоСодерж As Double) As Double

        ' Параметры ТО, Т и X занесены в список значений.
        Dim L0, I0, I1, D, QT, AU, BU As Double
        Dim P(10) As Double
        Dim M, N As Integer

        L0 = 28.966 / 0.2095 * (A(1, 11) / 12.01 + A(2, 11) / 4.032 - A(3, 11) / 32.0)
        D = ВлагоСодерж
        QT = (1.0 - D) * X / L0
        M = 6
        N = M + 1

        If (QT < 0.00001 AndAlso D < 0.00001) Then
            For I As Integer = 1 To N
                P(I) = A(1, I)
            Next
        Else
            If QT < 0.00001 Then
                For I As Integer = 1 To N
                    P(I) = A(1, I) * (1.0 - D) + A(4, I) * D
                Next
            Else
                If A(2, 11) < 0.999 Then
                    If (Math.Abs(A(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(A(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
                        For I As Integer = 1 To N
                            P(I) = (A(2, I) * QT + A(1, I)) / (1.0 + QT)
                        Next
                    Else
                        For I As Integer = 1 To N
                            AU = (44.01 * A(3, I) - 32.0 * A(5, I)) / 12.01
                            BU = (9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008
                            P(I) = ((A(1, 11) * AU + A(2, 11) * BU + A(3, 11) * A(5, I)) * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                        Next
                    End If
                Else
                    If X < 1.0 Then
                        For I As Integer = 1 To N
                            P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                        Next
                    Else
                        For I As Integer = 1 To N
                            P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * (1.0 - D) / L0 + D * A(4, I) + (1.0 - D) * A(1, I) + (QT - (1.0 - D) / L0) * A(6, I)) / (1.0 + QT)
                        Next
                    End If
                End If
            End If
        End If

        T0 = T0 * 0.001
        T = T * 0.001
        I0 = P(M + 1) / (M + 1)
        I1 = I0

        For I As Integer = 1 To M
            I0 = I0 * T0 + P(M + 1 - I) / (M + 1 - I)
            I1 = I1 * T + P(M + 1 - I) / (M + 1 - I)
        Next

        Return (I1 * T - I0 * T0) * 1000.0
        ' T = T * 1000.0
        ' T0 = T0 * 1000.0
    End Function

    ''' <summary>
    ''' Процедура - функция TI предназначена для определения температуры Тконечн по известным значениям начальной температуры Т0 и приращения энтальпии DI от этой температуры до температуры T  при заданном α АЛЬФА(Х=1/АЛЬФА). DI задается в ккал/кг.
    ''' Заголовок процедуры:TI (DI, TO, X, А, С).
    ''' </summary>
    ''' <param name="DI"></param>
    ''' <param name="T0"></param>
    ''' <param name="X"></param>
    ''' <param name="A"></param>
    ''' <param name="C"></param>
    ''' <param name="ВлагоСодерж"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TI(ByVal DI As Double,
                           ByVal T0 As Double,
                           ByVal X As Double,
                           ByVal A(,) As Double,
                           ByRef C() As Double,
                           ByVal ВлагоСодерж As Double) As Double

        ' Параметры DI , TO , X занесены в список значений.
        ' Наряду с определением температуры T, в процессе работы этой подпрограммы массив С заполняется значениями параметров, соответствующими этой температуре.
        ' В процессе определения температуры T выполнялись последовательные приближения по методу касательной, точность определения T = 0,001 К.

        Dim I0, I1, T1, T2, D, QT, AU, BU As Double
        Dim P(10) As Double
        Dim N, M As Integer
        Dim counterOverflow As Integer ' Счетчик Переп

        C(6) = 28.966 / 0.2095 * (A(1, 11) / 12.01 + A(2, 11) / 4.032 - A(3, 11) / 32.0)
        D = ВлагоСодерж
        QT = (1.0 - D) * X / C(6)
        M = 6
        N = M + 1

        If (QT < 0.00001 AndAlso D < 0.00001) Then
            C(3) = 29.27
            For I As Integer = 1 To N
                P(I) = A(1, I)
            Next
        Else
            If QT < 0.00001 Then
                C(3) = 29.27 * (1.0 + 0.60779 * D)
                For I As Integer = 1 To N
                    P(I) = A(1, I) * (1.0 - D) + A(4, I) * D
                Next
            Else
                If A(2, 11) < 0.999 Then
                    If (Math.Abs(A(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(A(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
                        C(3) = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
                        For I As Integer = 1 To N
                            P(I) = (A(2, I) * QT + A(1, I)) / (1.0 + QT)
                        Next
                    Else
                        C(3) = 29.27 * (1.0 + 28.966 * (A(2, 11) / 4.032 + A(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
                        For I As Integer = 1 To N
                            AU = (44.01 * A(3, I) - 32.0 * A(5, I)) / 12.01
                            BU = (9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008
                            P(I) = ((A(1, 11) * AU + A(2, 11) * BU + A(3, 11) * A(5, I)) * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                        Next
                    End If
                Else
                    If X < 1.0 Then
                        C(3) = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
                        For I As Integer = 1 To N
                            P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                        Next
                    Else
                        C(3) = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
                        For I As Integer = 1 To N
                            P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * (1.0 - D) / C(6) + D * A(4, I) + (1.0 - D) * A(1, I) + (QT - (1.0 - D) / C(6)) * A(6, I)) / (1.0 + QT)
                        Next
                    End If
                End If
            End If
        End If

        T0 = T0 * 0.001
        T1 = T0 + DI / 0.25 * 0.001
        I0 = P(M + 1) / (M + 1)

        For I As Integer = 1 To M
            I0 = I0 * T0 + P(M + 1 - I) / (M + 1 - I)
        Next

        I0 = I0 * T0 * 1000.0
        T2 = T1 ' начальное присваивание для последующей рокировки

        Do
            T1 = T2
            C(1) = P(M + 1)
            I1 = P(M + 1) / (M + 1)

            For I As Integer = 1 To M
                C(1) = C(1) * T1 + P(M + 1 - I)
                I1 = I1 * T1 + P(M + 1 - I) / (M + 1 - I)
            Next

            I1 = I1 * T1 * 1000.0
            T2 = T1 - (I1 - I0 - DI) / (1000.0 * C(1))
            counterOverflow += 1

            If counterOverflow > limitOverflow Then Exit Do
        Loop Until Math.Abs(T2 - T1) <= 0.000001 ' условие выхода

        C(4) = C(1) / (C(1) - C(3) / 426.9)
        ' TI = T2 * 1000.0
        Return T2 * 1000.0
        ' T0 = T0 * 1000.0
    End Function

    ''' <summary>
    ''' TPI предназначена для определения конечной температуры в адиабатическом процессе по заданным значениям начальной температуры TI и отношения конечного давления к начальному P2/P1 при известном АЛЬФА (Х=1/АЛЬФА).
    ''' TPI (PR, TI, X, A, C), PR - это P2/Р1
    ''' </summary>
    ''' <param name="PR"></param>
    ''' <param name="T1"></param>
    ''' <param name="X"></param>
    ''' <param name="A"></param>
    ''' <param name="C"></param>
    ''' <param name="ВлагоСодерж"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TPI(ByVal PR As Double,
                         ByVal T1 As Double,
                         ByVal X As Double,
                         ByVal A(,) As Double,
                         ByRef C() As Double,
                         ByVal ВлагоСодерж As Double) As Double

        ' Параметры PR, TI, X внесены в список значений.
        ' Попутно заполняется массив С (соответствует T2 ).
        ' Точность определения T2 в процессе последовательных приближений равна 0,001 К.

        Dim I1, I2, T2, T3, D, QT, AU, BU, DF As Double
        Dim M, N, L As Integer
        Dim P(10) As Double
        Dim counterOverflow As Integer ' Счетчик Переп

        C(6) = 28.966 / 0.2095 * (A(1, 11) / 12.01 + A(2, 11) / 4.032 - A(3, 11) / 32.0)
        D = ВлагоСодерж
        QT = (1.0 - D) * X / C(6)
        M = 6
        N = M + 1
        L = M - 1

        If (QT < 0.00001 AndAlso D < 0.00001) Then
            C(3) = 29.27
            For I As Integer = 1 To N
                P(I) = A(1, I)
            Next
        Else
            If QT < 0.00001 Then
                C(3) = 29.27 * (1.0 + 0.60779 * D)
                For I As Integer = 1 To N
                    P(I) = A(1, I) * (1.0 - D) + A(4, I) * D
                Next
            Else
                If A(2, 11) < 0.999 Then
                    If (Math.Abs(A(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(A(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
                        C(3) = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
                        For I As Integer = 1 To N
                            P(I) = (A(2, I) * QT + A(1, I)) / (1.0 + QT)
                        Next
                    Else
                        C(3) = 29.27 * (1.0 + 28.966 * (A(2, 11) / 4.032 + A(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
                        For I As Integer = 1 To N
                            AU = (44.01 * A(3, I) - 32.0 * A(5, I)) / 12.01
                            BU = (9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008
                            P(I) = ((A(1, 11) * AU + A(2, 11) * BU + A(3, 11) * A(5, I)) * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                        Next
                    End If
                Else
                    If X < 1.0 Then
                        C(3) = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
                        For I As Integer = 1 To N
                            P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * QT + A(1, I) * (1.0 - D) + A(4, I) * D) / (1.0 + QT)
                        Next
                    Else
                        C(3) = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
                        For I As Integer = 1 To N
                            P(I) = ((9.008 * A(4, I) - 8.0 * A(5, I)) / 1.008 * (1.0 - D) / C(6) + D * A(4, I) + (1.0 - D) * A(1, I) + (QT - (1.0 - D) / C(6)) * A(6, I)) / (1.0 + QT)
                        Next
                    End If
                End If
            End If
        End If

        T1 = T1 * 0.001
        T2 = T1 * PR ^ 0.25
        I1 = P(M + 1) / M

        For I As Integer = 1 To L
            I1 = I1 * T1 + P(M + 1 - I) / (M - I)
        Next

        I1 = I1 * T1
        T3 = T2

        Do
            T2 = T3 ' начальное присваивание для последующей рокировки
            DF = P(M + 1)
            I2 = P(M + 1) / M

            For I As Integer = 1 To L
                DF = DF * T2 + P(M + 1 - I)
                I2 = I2 * T2 + P(M + 1 - I) / (M - I)
            Next

            DF = DF + P(1) / T2
            I2 = I2 * T2 - I1 + P(1) * Math.Log(T2 / T1)
            T3 = T2 - (I2 - C(3) * Math.Log(PR) / 426.9) / DF
            counterOverflow += 1

            If counterOverflow > limitOverflow Then Exit Do
        Loop Until Math.Abs(T3 - T2) <= 0.000001

        C(1) = DF * T3
        C(4) = C(1) / (C(1) - C(3) / 426.9)
        Return T3 * 1000.0
        ' T1 = T1 * 1000.0
    End Function


    ''' <summary>
    ''' Нахождение Т*3 по пропускной способности турбины высокого лавления
    ''' </summary>
    ''' <param name="sigmaKC"></param>
    ''' <param name="Hu"></param>
    ''' <param name="valFсаТВД"></param>
    ''' <param name="Р2полнФиз"></param>
    ''' <param name="Т2физ"></param>
    ''' <param name="Gтопл"></param>
    ''' <param name="GВоздВГорлеСА"></param>
    ''' <param name="Т3ГазаВГорлеСА"></param>
    ''' <param name="A"></param>
    ''' <param name="C"></param>
    ''' <param name="ВлагосодВозд"></param>
    ''' <remarks></remarks>
    Private Sub Т3ПоПропСпособнТурбВД(ByVal sigmaKC As Double,
                                ByVal Hu As Double,
                                ByVal valFсаТВД As Double,
                                ByVal Р2полнФиз As Double,
                                ByVal Т2физ As Double,
                                ByVal Gтопл As Double,
                                ByRef GВоздВГорлеСА As Double,
                                ByRef Т3ГазаВГорлеСА As Double,
                                ByVal A(,) As Double,
                                ByRef C() As Double,
                                ByVal ВлагосодВозд As Double)

        Dim Gts, QT, XпарамАльфа As Double
        Dim HU0, GпропСпособСАТВД, TG0, ЭнтальпияВоздуха, ЭнтальпияТоплНач, IГазаГорлаСА As Double
        Dim ТконечнТопл As Double
        Dim counterOverflow As Integer ' Счетчик Переп

        HU0 = КоэфПолнСгорТоплКС * Hu
        Gts = Gтопл / 3600.0 ' секундный расход топлива в ОКС 
        GпропСпособСАТВД = GR * valFсаТВД / САТВДо305
        TG0 = 1000.0 ' начальная
        QT = Gts / GпропСпособСАТВД ' относительный расход топлива на единицу расхода воэдуха
        ТконечнТопл = tтопл + conКельвин

        ' АльфаКамеры = Gвоздуха_в_горении / (14.95 * РасходТопливаКамерыСгоранияКгСек)
        ' ЭнтальпияТоплНач = DI(T0, T, XпарамАльфа, A, ВлагосодВозд)   '  энтальпия топлива равна энтальпии воздуха и сокращается
        ' IГазаГорлаСА = ((GВоздВГорлеСА - Gts) * ЭнтальпияВоздуха + Gts * (HU0 + ЭнтальпияТоплНач)) / (GВоздВГорлеСА * (1 + QT))' энтальпия газа в горле СА
        ' решение задачи баланса тепла в камере сгорания

        Do
            ' уравнение 4.5.3
            ' вычислить новое GВоздВГорлеСА при новом TG0
            GВоздВГорлеСА = GпропСпособСАТВД * sigmaKC * Р2полнФиз / (1 + QT) / Math.Sqrt(TG0) ' газ в горле СА
            ' XпарамАльфа = 1 / (GВоздВГорлеСА / Gts / СтехиометрКоэфLo) поэтому преобразуем
            XпарамАльфа = (Gts * СтехиометрКоэфLo) / GВоздВГорлеСА ' α АЛЬФА(Х=1/АЛЬФА)
            ЭнтальпияТоплНач = DI(tbфиз, ТконечнТопл, XпарамАльфа, A, 0)   '  энтальпия топлива равна энтальпии воздуха и сокращается
            ЭнтальпияВоздуха = DI(tbфиз, Т2физ, XпарамАльфа, A, ВлагосодВозд) ' энтальпия воздуха
            ' IГазаГорлаСА = (ЭнтальпияВоздуха * GВоздВГорлеСА + Gts * HU0) / (GВоздВГорлеСА * (1 + QT)) ' энтальпия газа в горле СА
            IГазаГорлаСА = ((GВоздВГорлеСА - Gts) * ЭнтальпияВоздуха + Gts * (HU0 + ЭнтальпияТоплНач)) / (GВоздВГорлеСА * (1 + QT)) ' энтальпия газа в горле СА
            Т3ГазаВГорлеСА = TI(IГазаГорлаСА, tbфиз, XпарамАльфа, A, C, ВлагосодВозд)

            If Math.Abs(Т3ГазаВГорлеСА - TG0) < 0.01 Then Exit Do

            TG0 = Т3ГазаВГорлеСА
            QT = XпарамАльфа / СтехиометрКоэфLo ' восстановить QT из XпарамАльфа
            counterOverflow += 1

            If counterOverflow > limitOverflow Then Exit Do
        Loop
    End Sub
End Class

' ''' <summary>
' ''' Нахождение Т*3 по пропускной способности турбины высокого лавления
' ''' </summary>
' ''' <param name="sigmaKC"></param>
' ''' <param name="Hu"></param>
' ''' <param name="valFсаТВД"></param>
' ''' <param name="Р2полнФиз"></param>
' ''' <param name="Т2физ"></param>
' ''' <param name="Gтопл"></param>
' ''' <param name="GВоздВГорлеСА"></param>
' ''' <param name="Т3ГазаВГорлеСА"></param>
' ''' <param name="A"></param>
' ''' <param name="C"></param>
' ''' <param name="ВлагосодВозд"></param>
' ''' <remarks></remarks>
'Private Sub Т3ПоПропСпособнТурбВД(ByVal sigmaKC As Double,
'                            ByVal Hu As Double,
'                            ByVal valFсаТВД As Double,
'                            ByVal Р2полнФиз As Double,
'                            ByVal Т2физ As Double,
'                            ByVal Gтопл As Double,
'                            ByRef GВоздВГорлеСА As Double,
'                            ByRef Т3ГазаВГорлеСА As Double,
'                            ByVal A(,) As Double,
'                            ByRef C() As Double,
'                            ByVal ВлагосодВозд As Double)

'    Dim Gts, QT, XпарамАльфа As Double
'    Dim HU0, GпропСпособСАТВД, TG0, ЭнтальпияВоздуха, IГазаГорлаСА As Double 'ЭнтальпияТоплНач,
'    Dim СчетчикПереп As Integer
'    Const ТначТопл As Double = 293.15

'    HU0 = КоэфПолнСгорТоплКС * Hu
'    Gts = Gтопл / 3600.0 'секундный расход топлива в ОКС 
'    GпропСпособСАТВД = GR * valFсаТВД / САТВДо305
'    TG0 = 500.0 'начальная
'    QT = Gts / GпропСпособСАТВД 'относительный расход топлива на единицу расхода воэдуха
'    Т0масшт = ТначТопл 
'    Тмасшт = Т2физ 

'    'АльфаКамеры = Gвоздуха_в_горении / (14.94 * РасходТопливаКамерыСгоранияКгСек)
'    'ЭнтальпияТоплНач = DI(T0, T, XпарамАльфа, A, ВлагосодВозд)   ' энтальпия топлива равна энтальпии воздуха и сокращается
'    'IГазаГорлаСА = ((GВоздВГорлеСА - Gts) * ЭнтальпияВоздуха + Gts * (HU0 + ЭнтальпияТоплНач)) / (GВоздВГорлеСА * (1 + QT))'энтальпия газа в горле СА
'    'решение задачи баланса тепла в камере сгорания
'    Do
'        'уравнение 4.5.3
'        'вычислить новое GВоздВГорлеСА при новом TG0
'        GВоздВГорлеСА = GпропСпособСАТВД * sigmaKC * Р2полнФиз / (1 + QT) / Math.Sqrt(TG0) 'газ в горле СА

'        'XпарамАльфа = 1 / (GВоздВГорлеСА / Gts / СтехиометрКоэфLo) поэтому преобразуем
'        XпарамАльфа = (Gts * СтехиометрКоэфLo) / GВоздВГорлеСА 'α АЛЬФА(Х=1/АЛЬФА)
'        ЭнтальпияВоздуха = DI(Т0масшт, Тмасшт, XпарамАльфа, A, ВлагосодВозд) 'энтальпия воздуха
'        IГазаГорлаСА = (ЭнтальпияВоздуха * GВоздВГорлеСА + Gts * HU0) / (GВоздВГорлеСА * (1 + QT)) 'энтальпия газа в горле СА

'        Т3ГазаВГорлеСА = TI(IГазаГорлаСА, Т0масшт, XпарамАльфа, A, C, ВлагосодВозд)

'        If Math.Abs(Т3ГазаВГорлеСА - TG0) < 0.01 Then Exit Do
'        TG0 = Т3ГазаВГорлеСА

'        QT = XпарамАльфа / СтехиометрКоэфLo 'восстановить QT из XпарамАльфа
'        СчетчикПереп += 1
'        If СчетчикПереп > conСчПереп Then Exit Do
'    Loop
'End Sub

'Public Function ТестРасчетаБиблиотеки() As System.Data.DataSet Implements BaseFormKT.IClassCalculation.ТестРасчетаБиблиотеки
'    Dim myLinearAlgebra As New LinearAlgebra
'    With myLinearAlgebra
'        .matrixADataTextBox = "4.00, 2.00, -1.00; 1.00, 4.00, 1.00; 0.10, 1.00, 2.00;"
'        .matrixBDataTextBox = "2.00; 12.00; 10.00;"
'        '.operationsComboBox = Global.MathematicalLibrary.LinearAlgebra.EnumOperationsComboBox.SolveLinearEquations_AxB
'        .operationsComboBox = LinearAlgebra.EnumOperationsComboBox.SolveLinearEquations_AxB
'        .Compute()
'        ТестРасчетаБиблиотеки = .data
'    End With
'End Function

'Private Sub Иммитатор(ByVal blnSinTime As Boolean)
'    Static TimeIndex As Double
'    Dim I As Integer = 1
'    If blnSinTime Then
'        For Each rowРасчетныйПараметр As BaseFormDataSet.РасчетныеПараметрыRow In varProjectManager.РасчетныеПараметры.Rows
'            rowРасчетныйПараметр.ВычисленноеЗначениеВСИ = System.Math.Sin(TimeIndex / 20.0# * Math.PI * 2) * 10 + I '* I / 50 + I / 50
'            TimeIndex += 0.01
'            I += 1
'        Next
'    Else
'        For Each rowРасчетныйПараметр As BaseFormDataSet.РасчетныеПараметрыRow In varProjectManager.РасчетныеПараметры.Rows
'            rowРасчетныйПараметр.ВычисленноеЗначениеВСИ = I
'            I += 1
'        Next
'    End If
'End Sub


'    Private Sub ВычислитьРасчетныеПараметры()
'        'Dim Ps As Double
'        Try
'            ''***********************************
'            ''расчет Расхода воздуха на входе в КС
'            ''***********************************
'            'Рс_ср = (Рс1 + Рс2) / 2
'            'dРс_ср = (dРс1 + dРс2) / 2
'            'Тс = (Тс1 + Тс2) / 2
'            ''Показатель адиабаты
'            'If (Тс) < 50 Then
'            '    k = 1.4
'            'Else
'            '    k = Kadiobaty(Тс)
'            'End If
'            'Алин = (15.6 + 8.3 * Тс * 0.001 - 6.5 * (Тс * 0.001) ^ 2) * 0.000001
'            'Kt = 1 + Алин * (Тс - 20)
'            'Kт_поправка = Kt * Kt
'            't = Math.Abs(1 - dРс_ср / Рс_ср)

'            'e = (k - t ^ (2 / k)) / (k - 1)
'            'e = e * (1 - md * md * md * md) / (1 - md * md * md * md * t ^ (2 / k))
'            'e = e * ((1 - t) ^ ((k - 1) / k)) / (1 - t)
'            'e = Math.Sqrt(e)

'            'Gb = 51.556 * Math.Sqrt(Рс_ср * dРс_ср / (Тс + conКельвин)) * e * Kт_поправка

'            ''***********************************
'            ''расчет среднего расхода топлива камеры сгорания и подогрева
'            ''***********************************

'            ''найти коэф. учитывающий плотность
'            'Ps = clAir.КоэфУчитывающийПлотность(Тт_маг)
'            ''выведем расход в кг час ((Gт1 + Gт2) * Ps = вычисление л/час в кг/час)(Было литрах в час)
'            ''вывести на индикатор суммарный расход топлива в кг/час
'            'GтSum = (Gт1 + Gт2) * Ps + Gт_кп 'Gт1 + Gт2 + Gт_кп
'            'Gt_кс_ср = (Gт1 + Gт2) * Ps / 3600.0#  'перевод л/час в кг/сек камеры сгорания

'            ''Ps = Spline3Interpolate(UBound(TblКоэффициентыПлотностиТоплива, 2) + 1, TblКоэффициентыПлотностиТоплива, Тт_кп)
'            ''Gt_кп = Gт_кп * Ps / 3600.0#  'перевод л/час в кг/сек камеры подогрева
'            'Gt_кп = Gт_кп / 3600.0# 'перевод кг/час в кг/сек камеры подогрева

'            ''***********************************
'            ''расчет коэф. избытка воздуха
'            ''***********************************
'            'Gв_сум = Gb + Gt_кп 'суммарный расход воздуха после камеры сгорания -11
'            'АльфаКамеры = Gв_сум / (14.94 * Gt_кс_ср)

'            '***********************************
'            'расчет абсолютного полного давления воздуха на входе КС
'            '***********************************
'            'Рст_абс_ср = Math.Abs((Р311_1 + Р311_2 + Р311_3 + Р311_4 + Р311_5 + Р311_6) / 6) 'барометр уже добавлен
'            'к дифференциальным замерам добавить базовое давление
'            'Dim ИмяПояса As String
'            'For Зонд As Integer = 1 To 3
'            '    For Пояс As Integer = 1 To 5
'            '        ИмяПояса = "dР310-" & Зонд.ToString & "-" & Пояс.ToString
'            '        'перепад скорее всего показывает избыточное, значит в тарировке он с плюсом
'            '        Manager.ИзмеренныеПараметры.FindByИмяПараметра(ИмяПояса).ЗначениеВСИ += Рст_абс_ср
'            '    Next
'            'Next

'            'MathematicalLibrary.PlotSurface.PlotSurface(Рст_абс_ср)

'            '***********************************
'            'среднее значение температуры торможения -19
'            '***********************************
'            'T309 = (Т309_1 + Т309_2 + Т309_3 + Т309_4 + Т309_5 + Т309_6) / 6

'            '***********************************
'            'приведенная скорость газового потока -20
'            ''***********************************
'            'Lamda = LamdaFun(Тс, T309, Gb, Fdif, Рст_абс_ср, Рв_вх_абс_полн_ср)

'            With Manager
'                ''по имени параметра strИмяПараметраГрафика определяем нужную функцию приведения
'                ''("n1") 'который измеряет
'                ''должна быть вызвана функция приведения параметра "n1" например
'                'If n1ГПриводить Then
'                '    'должна быть вызвана функция приведения параметра "n1" например
'                '    .РасчетныеПараметры.FindByИмяПараметра(cn1Г).ВычисленноеЗначениеВСИ = Air.funПривестиN(.ИзмеренныеПараметры.FindByИмяПараметра("n1").ЗначениеВСИ, tm)
'                'Else
'                '    'приводить не надо, просто копирование
'                '    .РасчетныеПараметры.FindByИмяПараметра(cn1Г).ВычисленноеЗначениеВСИ = .ИзмеренныеПараметры.FindByИмяПараметра("n1").ЗначениеВСИ
'                'End If

'                ''cGвМПолеДавлений
'                '.РасчетныеПараметры.FindByИмяПараметра(cGвМПолеДавлений).ВычисленноеЗначениеВСИ = funВычислитьGвМПолеДавлений(GвМПолеДавленийПриводить)
'                ''или так
'                ''.РасчетныеПараметры.FindByИмяПараметра(cGвМПолеДавлений).ВычисленноеЗначениеВСИ = funВычислитьGвМПолеДавлений(.НастроечныеПараметры.FindByИмяПараметра("GвМПолеДавленийПриводить").ЛогическоеЗначение)

'                ''cGвМПито
'                '.РасчетныеПараметры.FindByИмяПараметра(cGвМПито).ВычисленноеЗначениеВСИ = funВычислитьGвМПито(GвМПитоПриводить)

'                ''cПиК
'                '.РасчетныеПараметры.FindByИмяПараметра(cПиК).ВычисленноеЗначениеВСИ = funВычислитьПиК()

'                ''cКПДадиабат
'                '.РасчетныеПараметры.FindByИмяПараметра(cКПДадиабат).ВычисленноеЗначениеВСИ = funВычислитьКПДадиабат()

'                ''cnИГ_03Г
'                ''который измеряет
'                'nИГ_03 = .ИзмеренныеПараметры.FindByИмяПараметра("nИГ-03").ЗначениеВСИ / 46325 'коэф. перевода  n=1 при N=45190

'                'If nИГ_03ГПриводить Then
'                '    'должна быть вызвана функция приведения параметра "n1" например
'                '    .РасчетныеПараметры.FindByИмяПараметра(cnИГ_03Г).ВычисленноеЗначениеВСИ = Air.funПривестиN(nИГ_03, tm)
'                'Else
'                '    'приводить не надо, просто копирование
'                '    .РасчетныеПараметры.FindByИмяПараметра(cnИГ_03Г).ВычисленноеЗначениеВСИ = nИГ_03
'                'End If

'                ''cnПиК_GвМПД
'                ''должна быть вызвана функция вычисления параметра "GПиК/GвМПД" например
'                '.РасчетныеПараметры.FindByИмяПараметра(cnПиК_GвМПД).ВычисленноеЗначениеВСИ = funВычислитьПиК_GвМПД()
'            End With


'            '*********************************************

'            'If blnИзмерениеПоТемпературам Then
'            '***********************************
'            'расчет интегральной температура газа по мерному сечению на поясе
'            'внимание этот расчет для равномерного положения термопар относительно друг друга
'            '**********************************
'            Dim Ystart As Double, Yend As Double
'            НайтиЗначенияТемпературНаСтенках(КоординатыТермопар, arrТекущаяПоПоясам, ШиринаМерногоУчастка, Ystart, Yend)
'            y(0) = Ystart
'            y(ЧислоТермопар + 1) = Yend

'            T_интегр = ИнтегрированиеРадиальнойЭпюрыНаПроизвольныхКоординатах(КоординатыТермопар, arrТекущаяПоПоясам, ШиринаМерногоУчастка)

'            '***********************************
'            'расчет Качество 
'            '**********************************
'            mКачество = КачествоFun(T_интегр, Тсредняя_газа_на_входе, Тг_расчет)

'            'после замены на координатнике термопар ттемпература на выходе измеряется как средняя из 6
'            'Т3г_сред = (T3_1 + T3_2 + T3_3 + T3_4 + T3_5 + T3_6) / 6

'            '0.3964 коэф. от адиабаты
'            'gFun_Lamda = Gb * Math.Sqrt(T309 + conКельвин) / (0.3964 * Рв_вх_абс_полн_ср * Fk)







'Барометр-барометр
'gp-расход топлива камеры подогрева
'gc-расход топлива камеры сгорания
'Gвоздуха-расход Gв основной
'go-расход Go отбора
'Pp-плотность топлива от температуры
'Ps-плотность топлива от температуры
'gg-расход газа
'gv-расход воздуха участвующий в горении
'gs-суммарный расход топлива через к.с. и к.п.
'Gотбора-относительный расход отбираемого газа
'АльфаКамеры-коэффициент избытка воздуха
'ls-суммарный коэффициент избытка воздуха
'Лямда-приведенная скорость газового потока на входе в К.С.
'Dim temp, Отн_расход_отбир_газа, КплотностьКС, КплотностьКП, Gтопл_в_отбир_газе As Double
'Dim Gасход_газа, Gвоздуха_в_горении, Сумм_коэф_избытка_воздуха As Double
'Dim funGотА As Double
'Dim Gрасход_отбора, РасходТопливаКамерыПодогреваКгСек, РасходТопливаКамерыСгоранияКгСек As Double

'найти коэф. учитывающий плотность
'КплотностьКС = clAir.КоэфУчитывающийПлотность(ТтопливаКС)
'КплотностьКП = clAir.КоэфУчитывающийПлотность(ТтопливаКП)
'расход топлива камеры сгорания
'РасходТопливаКамерыСгоранияКгСек = РасходТопливаКамерыСгорания * КплотностьКС / 3600.0# 'перевод л/час в кг/сек камеры сгорания
''расход топлива камеры подогрева
'РасходТопливаКамерыПодогреваКгСек = РасходТопливаКамерыПодогрева * КплотностьКП / 3600.0# 'перевод л/час в кг/сек камеры подогрева
'расход Gв основной
'Gвоздуха = Расход(ДавлениеВоздухаНаВходе, ПерепадДавленияВоздухаНаВходе, T3мерн_участка, D20отвОсн, D20трубОсн, Ks, 1)
'расход Go отбора
'Gрасход_отбора = Расход(ДавлениеМагистралеОтбора, ПерепадДавленияВоздухаОтбора, Тотбора, D20отвОтб, D20трубОтб, Ks, 2)
'расход газа gg
'Gасход_газа = Gвоздуха + РасходТопливаКамерыПодогреваКгСек
'относительный расход отбираемого газа
'Отн_расход_отбир_газа = Gрасход_отбора / Gасход_газа
'кол. топлива в отбираемом газе
'Gтопл_в_отбир_газе = РасходТопливаКамерыПодогреваКгСек * Отн_расход_отбир_газа
'расход воздуха участвующий в горении
'Gвоздуха_в_горении = Gвоздуха - Gрасход_отбора + Gтопл_в_отбир_газе
''суммарный расход топлива через к.с. и к.п.
'Gсум_расход_топливаКС_КП = РасходТопливаКамерыСгоранияКгСек + РасходТопливаКамерыПодогреваКгСек - Gтопл_в_отбир_газе
'Gсум_расход_топливаКС_КП_кг_час = Gсум_расход_топливаКС_КП * 3600
''коэффициент избытка воздуха
'АльфаКамеры = Gвоздуха_в_горении / (14.94 * РасходТопливаКамерыСгоранияКгСек)
''относительный расход отбираемого газа
'Gотбора_относительный = Gрасход_отбора * 100.0# / Gасход_газа
'суммарный коэффициент избытка воздуха ни где не участвует
'Сумм_коэф_избытка_воздуха = Gвоздуха_в_горении / (Gсум_расход_топливаКС_КП * 14.94)
'temp = System.Math.Sqrt(1.358 * (2 / 2.358) ^ (2.358 / 0.358)) * System.Math.Sqrt(g / 29.4)
'вычисление g(a)
'funGотА = Gасход_газа * System.Math.Sqrt(Тсредняя_газа_на_входе + conКельвин) / (temp * Р310полное_воздуха_на_входе_КС * Fdif)
'приведенная скорость газового потока 
'Лямда = clAir.ПриведеннаяСкорость(funGотА)

'Dim T3b, Gtksb, Gbksb, АльфаСум As Double
'T3b = 308.7121+conКельвин
'Gtksb = 1473.427 / 3600
'Gbksb = 12.49772
'АльфаСум = 2.043875
'***********************************
'расчетная температура газа в Цельсий
'**********************************
'Т3 температура на входе в мерный участок
'Gтопл суммарный
'Gвоздкс
'Тг_расчет = clAir.РасчётнаяТемпература(T3мерн_участка + conКельвин, Gсум_расход_топливаКС_КП, Gвоздуха_в_горении, АльфаКамеры)
'        Catch ex As Exception
'            Description = "Процедура: ВычислитьРасчетныеПараметры"
'            'перенаправление встроенной ошибки
'            Dim fireDataErrorEventArgs As New IClassCalculation.DataErrorEventArgs(ex.Message, Description)
'            '  Теперь вызов события с помощью вызова делегата. Проходя в
'            '   object которое инициирует  событие (Me) такое же как FireEventArgs. 
'            '  Вызов обязан соответствовать сигнатуре FireEventHandler.
'            RaiseEvent DataError(Me, fireDataErrorEventArgs)
'            mФормаРодителя.varfrmСнятиеКТ.tsTextBoxОшибкаРасчета.Visible = True

'        End Try
'        mФормаРодителя.varfrmСнятиеКТ.tsTextBoxОшибкаРасчета.Visible = CBool(Тг_расчет = 0 - conКельвин)
'    End Sub



'Private action As System.Action(Of FileInfo)
'Private match As System.Predicate(Of FileInfo)
'Private fileList As New List(Of FileInfo)
'Private fileArray() As FileInfo

'Private Sub outputWindowRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles outputWindowRadioButton.CheckedChanged
'    action = New System.Action(Of FileInfo) _
'     (AddressOf DisplayInOutputWindow)
'End Sub

'Private Sub forEachListButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles forEachListButton.Click
'    ResetListBox()
'    fileList.ForEach(action)
'End Sub

'Private Sub DisplayInOutputWindow(ByVal file As FileInfo)
'    Debug.WriteLine(String.Format("{0} ({1} bytes)", _
'     file.Name, file.Length))
'End Sub

'Private Sub smallFilesRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles smallFilesRadioButton.CheckedChanged
'    match = New System.Predicate(Of FileInfo) _
'     (AddressOf IsSmall)
'End Sub
'Private Sub findAllListButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles findAllListButton.Click
'    ResetListBox()

'    ' Create a list containing matching files,
'    ' and then take the appropriate action.
'    Dim subList As List(Of FileInfo) = fileList.FindAll(match)
'    subList.ForEach(action)
'End Sub

'Private Function IsSmall( _
'ByVal file As FileInfo) As Boolean

'    ' Return True if the file's length is less than 500 bytes.
'    Return file.Length < 500
'End Function


''Dim A(26) As Double по Янкину
''массив начинается с 0
'Private A() As Double = New Double() {-100, 29.27, 0.0, 1.0775667, 0.0, СтехиометрКоэфLo, 10331, _
'                      0.24582008, -0.035655427, -0.04983327, 0.57930396, _
'                      -1.0122514, 0.86642789, -0.42615499, 0.12290562, _
'                      -0.019373617, 0.0012910064, 0.16192593, 1.434004, _
'                      -1.5125178, -0.049143299, 2.4095975, -2.9655719, _
'                      1.7457616, -0.56350004, 0.095991782, -0.0067651442}
''common
'Private M, M1, ij, ii As Integer
'Private M0, R0, R, AR, Q, T2, T1, _DI, KT As Double
''common
'  из Private SuВычислитьФизические()
' удаленные
'Dim Thhot, GBPHYZca As Double
'Dim GBCA, T3f As Double
'где расчет T3ТДРфиз
'Т3_ТВД(sigmaKC, Hu, valFсаТВД, Р2полнФиз, Т2физ, GтТДР, GBPHYZca, Thhot) 'виктор
'T3физТДР = Thhot '-режим Мах 'виктор
'где Уравнение баланса тепла
'Т3_ТВД(sigmaKC, Hu, valFсаТВД, Р2полнФиз, Т2физ, Gтопл, GBPHYZca, Thhot)
'T3fizmem = Thhot  'TODO:хуйня    ' - режим Мах, Дросс, Мг:
'Т3_Я(sigmaKC, Hu, valFсаТВД, Р2полнФиз, Т2физ, Gтопл, GBCA, T3f)
'
'


' ''' <summary>
' ''' 
' ''' </summary>
' ''' <param name="sigmaKC"></param>
' ''' <param name="Hu"></param>
' ''' <param name="valFсаТВД"></param>
' ''' <param name="Р2полнФиз"></param>
' ''' <param name="Т2физ"></param>
' ''' <param name="Gтопл"></param>
' ''' <param name="GВоздВГорлеСА"></param>
' ''' <param name="T3f"></param>
' ''' <remarks></remarks>
'Private Sub Т3_ТВД(ByVal sigmaKC As Double,
'                ByVal Hu As Double,
'                ByVal valFсаТВД As Double,
'                ByVal Р2полнФиз As Double,
'                ByVal Т2физ As Double,
'                ByVal Gтопл As Double,
'                ByRef GВоздВГорлеСА As Double,
'                ByRef T3f As Double)

'    Dim QT, Tg3, dTg7, Tg1, dTg10, Tg10, qT1, GпропСпособСАТВД, dTg1, Tg, dTg0, Tg0 As Double 'Y1,
'    Dim СчетчикПереп As Integer

'    QT = Gтопл / 3600.0 / 60.0

'    Tg0 = 40.347 + 0.928819 * Т2физ + 41830.4 * QT - 6.27948 * Т2физ * QT + 0.000029283 * Т2физ * Т2физ - 160994.0 * QT * QT
'    dTg0 = -0.00000003574 * Tg0 * Tg0 * Tg0 + 0.0001307 * Tg0 * Tg0 - 0.15396 * Tg0 + 58.27
'    Tg = Tg0 - dTg0

'    dTg1 = -0.353 + 0.0093184 * Tg - 0.0086214 * Т2физ
'    Tg = Tg + dTg1 * (Hu - 10250) / 100 'почему за базу Hu=10250
'    GпропСпособСАТВД = GR * valFсаТВД / САТВДо305

'    GВоздВГорлеСА = GпропСпособСАТВД * sigmaKC * Р2полнФиз / (1 + QT) / Math.Sqrt(Tg)
'    qT1 = Gтопл / 3600.0 / GВоздВГорлеСА
'    Do
'        Tg10 = 40.347 + 0.928819 * Т2физ + 41830.4 * qT1 - 6.27948 * Т2физ * qT1 + 0.000029283 * Т2физ * Т2физ - 160994 * qT1 * qT1
'        dTg10 = -0.00000003574 * Tg10 * Tg10 * Tg10 + 0.0001307 * Tg10 * Tg10 - 0.15396 * Tg10 + 58.27
'        Tg1 = Tg10 - dTg10
'        dTg7 = -0.353 + 0.0093184 * Tg1 - 0.0086214 * Т2физ
'        Tg3 = Tg1 + dTg7 * (Hu - 10250) / 100 'почему за базу Hu=10250

'        GВоздВГорлеСА = GпропСпособСАТВД * sigmaKC * Р2полнФиз / (1 + qT1) / Math.Sqrt(Tg3)
'        'Y1 = Math.Abs(Tg1 - Tg)
'        If Math.Abs(Tg1 - Tg) <= 0.01 Then Exit Do
'        qT1 = Gтопл / 3600.0 / GВоздВГорлеСА
'        Tg = Tg1
'        СчетчикПереп += 1
'        If СчетчикПереп > conСчПереп Then Exit Do
'    Loop

'    T3f = Tg3
'End Sub

' ''' <summary>
' ''' Нахождение Т*3 по Янкину с помощью Р400 и Р403
' ''' </summary>
' ''' <param name="sigmaKC"></param>
' ''' <param name="Hu"></param>
' ''' <param name="valFсаТВД"></param>
' ''' <param name="Р2полнФиз"></param>
' ''' <param name="Т2физ"></param>
' ''' <param name="Gтопл"></param>
' ''' <param name="GВоздВГорлеСА"></param>
' ''' <param name="T3f"></param>
' ''' <remarks></remarks>
'Private Sub Т3_Я(ByVal sigmaKC As Double,
'                 ByVal Hu As Double,
'                 ByVal valFсаТВД As Double,
'                 ByVal Р2полнФиз As Double,
'                 ByVal Т2физ As Double,
'                 ByVal Gтопл As Double,
'                 ByRef GВоздВГорлеСА As Double,
'                 ByRef T3f As Double)

'    Dim KH, QTT1, AL1, GВоздВГорлеСАнач, GпропСпособСАТВД, QT3, QTT, TG0, HU0, AL, VH, MH As Double
'    Dim СчетчикПереп As Integer

'    Dim AG As Double = 8377.74
'    Q = 0
'    Call P400(9)
'    T1 = 785
'    Call P400(5)
'    KH = KT
'    'RH = R
'    MH = 0.5
'    VH = MH * Math.Sqrt(KH * con_g * A(1) * Т2физ)
'    _DI = VH * VH / AG
'    Call P400(2)
'    Call P400(4)
'    AL = GR / Gтопл / A(5) * 3600 'ВЫЧИСЛЕНИЕ AL ПО GT
'    HU0 = КоэфПолнСгорТоплКС * Hu

'    Call P403(Т2физ, HU0, 0, AL, QT3, TG0, Q, 1)
'    'Р403-при S=1-ОПРЕДЕЛЕНИЕ температуры на выходе из камеры 
'    'сгорания по заданным коэффициенту -AL и т-ре на вх в КС-Тн0
'    T3f = TG0
'    QTT = QT3
'    GпропСпособСАТВД = GR * valFсаТВД / САТВДо305
'    GВоздВГорлеСАнач = GпропСпособСАТВД * sigmaKC * Р2полнФиз / (1 + QTT) / Math.Sqrt(TG0)

'    AL1 = GВоздВГорлеСАнач / Gтопл / A(5) * 3600.0 ' ВЫЧИСЛЕНИЕ AL1 В ПРОЦЕССЕ ИТЕРАЦИЙ ПРИ РЕШ ЗАД ПО GT
'    Do
'        Call P403(Т2физ, HU0, 0, AL1, QT3, TG0, Q, 1)
'        QTT1 = QT3
'        GВоздВГорлеСА = GпропСпособСАТВД * sigmaKC * Р2полнФиз / (1 + QTT1) / Math.Sqrt(TG0)

'        If Math.Abs(TG0 - T3f) <= 0.1 Then Exit Do

'        AL1 = GВоздВГорлеСА / Gтопл / A(5) * 3600.0
'        T3f = TG0

'        СчетчикПереп += 1
'        If СчетчикПереп > conСчПереп Then Exit Do
'    Loop
'End Sub


'    'FUNCTION DI(T0,T,X,A)  ( P400(1))
'    'FUNCTION TI(DI,T0,X,A,C) ( P400(2))
'    'TPI(T1,T2,X,A)   Р400(7): Т1,PI=P2/P1,qT || Т2
'    Private Sub P400(ByVal S As Integer)
'        Dim N, L, K, J, IW, NN, Z, F, E As Integer
'        Dim T(11) As Double
'        Dim TKP As Double
'        '     N-порядок аппроксимирующего полинома  стр 96 Янкина.
'        N = 9   '-ЯНКИН
'        T(4) = 1.0 + Q

'        'реально вызываются 1,2,4,5,9
'        Select Case S
'            Case 1
'                K = 1
'                T(8) = 0
'                Z = 7
'                L = 1
'                GoTo 31
'            Case 2
'                T(1) = _DI
'                T2 = 100.0
'                K = 5

'                T(8) = 0
'                Z = 7
'                L = 1
'                GoTo 31

'                'Case 3
'                '    K = 1
'                '    R = R0 * (1 + M0 * Q) / T(4)
'                '    AR = R / 427
'                '    IW = K
'                '    GoTo 902
'            Case 4
'                F = 1
'                T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Math.Log(T2 / T1)
'                Z = 8
'                L = 2
'                GoTo 31
'            Case 5
'                J = 1
'                Call Блок55(T, NN, A, K, IW)
'                GoTo 902
'                'Case 6
'                '    If Math.Abs(T2 - T1) <= 0.5 Then
'                '        J = 11
'                '        Call Блок55(T, NN, A, K, IW)
'                '        GoTo 902
'                '    Else
'                '        K = 3
'                '        T(8) = 0
'                '        Z = 7
'                '        L = 1
'                '        GoTo 31
'                '    End If
'                'Case 7
'                '    T(11) = PI
'                '    T2 = 100
'                '    F = 9
'                '    T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Math.Log(T2 / T1)
'                '    Z = 8
'                '    L = 2
'                '    GoTo 31
'                'Case 8
'                '    T2 = 0.8 * T1
'                '    T(11) = T2
'                '    E = -1
'                '    T(10) = T1
'                '    T1 = T2
'                '    J = 12
'                '    Call Блок55(T, NN, A, K, IW)
'                '    GoTo 902
'            Case 9
'                If (Q < 1.0 / A(5) OrElse A(2) = 0.0) Then
'                    M = 0
'                    R0 = A(1)
'                    M0 = A(3)
'                Else
'                    M = 2 * (N + 1)
'                    R0 = A(2)
'                    M0 = A(4)
'                End If

'                M1 = M + N + 1
'                Exit Sub
'            Case Else
'                Throw New Exception("P400(ByVal S As Integer) S= " & S.ToString)
'        End Select


'522:    T2 = T(9) + (100.0 - T(9)) * (T(1) - T(8)) / (T(10) - T(8))
'        If Math.Abs(T2 - T(9)) < 0.1 Then
'            If S = 2 Then
'                _DI = T(1)
'            Else
'                'PI = T(11)
'            End If
'            Exit Sub
'        Else
'            T(9) = T2
'            If S = 2 Then
'                K = 2
'                T(8) = 0
'                Z = 7
'                L = 1
'            Else
'                F = 2
'                T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Math.Log(T2 / T1)
'                Z = 8
'                L = 2
'            End If
'        End If

'31:
'        T(5) = 1
'        T(6) = 1
'        T(7) = 0
'        T(2) = T2 / 1000
'        T(3) = T1 / 1000
'        NN = N + 7

'        For I As Integer = Z To NN
'            T(5) = T(5) * T(2)
'            T(6) = T(6) * T(3)
'            T(7) = T(7) + 1
'            ii = I + M
'            ij = I + M1
'            T(8) = T(8) + (T(5) - T(6)) * (A(ii) + Q * A(ij)) / (T(7) * T(4))
'        Next

'        If L = 1 Then
'            _DI = T(8) * 1000
'            T(8) = _DI
'            IW = K
'            GoTo 902
'        Else
'            If F = 2 Then
'                GoTo 522
'            Else
'                K = 6
'                R = R0 * (1 + M0 * Q) / T(4)
'                AR = R / 427
'                IW = K
'                GoTo 902
'            End If
'        End If
'        Exit Sub



'        '902  goto (999,522,563,581,34,531,532,564,571,582,561,58),IW
'902:
'        Select Case IW
'            Case 1
'                'Exit Sub
'            Case 2
'                GoTo 522
'            Case 3
'                '563:
'                T(1) = _DI / (T2 - T1)
'                K = 8
'                R = R0 * (1 + M0 * Q) / T(4)
'                AR = R / 427
'                IW = K
'                GoTo 902
'            Case 4
'                T(1) = _DI + AR * KT * T2 / 2.0
'                If E <= 0 Then
'                    T(9) = T(1)
'                    T2 = T1
'                    E = 1
'                    T1 = T2
'                    J = 12
'                    Call Блок55(T, NN, A, K, IW)
'                    GoTo 902
'                Else
'                    TKP = (T2 * T(9) - T(11) * T(1)) / (T(9) - T(1))
'                    If Math.Abs(T2 - TKP) < 0.1 Then
'                        T2 = TKP
'                        F = 10
'                        T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Math.Log(T2 / T1)
'                        Z = 8
'                        L = 2
'                        GoTo 31
'                    Else
'                        T2 = TKP
'                        T1 = T2
'                        J = 12
'                        Call Блок55(T, NN, A, K, IW)
'                        GoTo 902
'                    End If
'                End If
'            Case 5
'                '34:
'                T(10) = T(8)
'                T(8) = 0.0
'                T(9) = T1
'                GoTo 522
'            Case 6
'                '531:
'                'PI = Exp(T(8) / AR)
'                IW = F
'                GoTo 902
'            Case 7
'                '532:
'                KT = T(3) / (T(3) - AR)
'                IW = J
'                GoTo 902
'            Case 8
'                '564:
'                'KC = T(1) / (T(1) - AR)
'                'Exit Sub
'            Case 9
'                '571
'                T(1) = AR * Math.Log(T(11))
'                T(10) = T(8)
'                T(8) = 0.0
'                T(9) = T1
'                GoTo 522
'            Case 10
'                '582:
'                'PKP = PI
'                'AKP = Math.Sqrt(g * KT * R * TKP)
'            Case 11
'                '561:
'                'KC = KT
'                'Exit Sub
'            Case 12
'                '58:
'                T1 = T(10)
'                K = 4
'                T(8) = 0
'                Z = 7
'                L = 1
'                GoTo 31
'            Case Else
'                Throw New Exception("Select Case IW = " & IW.ToString)
'        End Select
'    End Sub

'    Private Sub Блок55(ByRef T() As Double, ByRef N As Double, ByVal A() As Double, ByRef K As Integer, ByRef IW As Integer)
'        Dim I, NN As Integer
'        T(1) = T1 / 1000
'        T(2) = 1
'        T(3) = (A(M + 7) + Q * A(M1 + 7)) / T(4)
'        NN = N + 7

'        For I = 8 To NN
'            T(2) = T(1) * T(2)
'            ii = I + M
'            ij = I + M1
'            T(3) = T(3) + T(2) * (A(ii) + Q * A(ij)) / T(4)
'        Next
'        K = 7

'        R = R0 * (1 + M0 * Q) / T(4)
'        AR = R / 427
'        IW = K
'    End Sub

'    ''' <summary>
'    ''' Р403- тепловой расчет КС(топливо керосин или водород)
'    ''' </summary>
'    ''' <param name="X1"></param>
'    ''' <param name="X2"></param>
'    ''' <param name="X3"></param>
'    ''' <param name="X4"></param>
'    ''' <param name="Y1"></param>
'    ''' <param name="Y2"></param>
'    ''' <param name="Y3"></param>
'    ''' <param name="S"></param>
'    ''' <remarks></remarks>
'    Private Sub P403(ByVal X1 As Double, ByVal X2 As Double, ByVal X3 As Double, ByRef X4 As Double, ByRef Y1 As Double, ByRef Y2 As Double, ByRef Y3 As Double, ByVal S As Integer)
'        Dim V(9) As Double
'        Dim K, L As Integer

'        Q = X3
'        Call P400(9)

'        V(4) = 1.0 + Q
'        L = -1
'        If S <> 0 Then
'            V(2) = X4
'            X4 = X1
'            V(3) = 200
'            L = 1
'        End If

'        Do
'            If S <> 0 Then X4 = X4 + V(3)
'            T1 = X1
'            T2 = X4
'            Call P400(1)
'            V(1) = _DI
'            T1 = 293
'            V(5) = Q
'            K = M
'            Q = 0
'            M = M1
'            Call P400(1)
'            Q = V(5)
'            M = K
'            Y1 = V(4) * V(1) / (X2 - _DI)
'            Y2 = 1 / ((Y1 + Q) * A(5))

'            If L < 0 Then ' выполнится когда If S = 0 Then
'                Y3 = Q + Y1
'                Exit Do
'            Else
'                If V(3) < 0.1 Then
'                    Y2 = X4
'                    X4 = V(2)
'                    Y3 = Q + Y1
'                    Exit Do
'                Else
'                    If (Y2 - V(2)) < 0 Then
'                        X4 = X4 - V(3)
'                        V(3) = V(3) / 2
'                    End If
'                End If
'            End If
'        Loop
'    End Sub




'Процедура - функция PIT предназначена для определения отношения конечного давления в адиабатическом процессе сжатия или расширения к начальному
'P1 по заданным значениям начальной T1 и конечной T2  температур при известном α.
'Заголовок процедуры: 'REAL' 'PROCEDURE'- PIT (TI, T2, X, A).
'Параметры Т1 , Т2 и X внесены в список значений.



'Процедура - функция TKR предназначена для определения критической температуры Ткр по заданной температуре торможения Т1 при известном α.
'Заголовок процедуры: REAL' 'PROCEDURE'- TKR (TI, X, А, С).
'TI и X внесены в список значений.
'Попутно заполняется массив С (соответствует Tkp Точность определения Tkp в процессе последовательных приближений равна 0,001 К.

'Процедуры CPS, DI, TI, PIT, TPI, TKR позволяют вычислять термодинамические параметры в следующих единицах измерения: Ср - ккая/(кг*К), i - ккал/кг, R – кгС*М/(кг*К).
'Для вычисления этих же величин в системе СИ используются аналогичные процедуры- CPSI, DISI, TISI, TPISI, TKRSI (см. рекомендуемое приложение 2).

'4.3. Подпрограммы для расчета параметров истечения из сопел
'Процедура FULEXP предназначена для расчета параметров истечения из сопла Лаваля полного расширения.

'Заголовок процедуры: 'PROCEDURE' FULEXP (GGAS, FER, LAT, TF, XF, PF, PH, SIG, PHI, A, G, GR).
'Параметр XF внесен в список значений ('Value'). Назначение формальных параметров:
'GGAS – Gф
'FKR - Fскр
'LAT – λскрs
'TF – ТФ*
'XF – 1/αсум
'PF – РФ*
'PH – PH*

'SI G - управляющий сигнал; при SIG - О значение phic  равно величине фактического параметра, соответствующёго формальному параметру PHI ;
'при SIG = 1 phic  определяется по формуле: phic  = 1-0,012*SQR(Fc/FCкр)	. 
'в этом случае значение PHI может быть произвольным. 
'Массивы A, G. CR описаны выше.
'Параметры GGAS, TF, XF, PF, РН, SIG, PHI, А являются входными.
'FKR, LAT, С, CR - результаты, содержимое массива С соответствует температуре Тскр .

'Процедура CONNOZ предназначена для расчета параметров истечения из сужающегося сопла.
'Заголовок процедуры: 'PROCEDURE CONNOZ (GGAS,FKP, LAT, TP, XP, PP, SIG, PHI, PH,A, C, CR).
'Параметр XP включен в список значений.
'Назначение формальных параметров в точности такое же, как в процедуре FULEXP.
'При SIG= 1 ϕc - 0,988.

'Процедура CONDIV предназначена для расчета параметров истечения из сопла Лаваля с заданной площадью среза.
'Заголовок процедуры:
'PROCEDURE' CONDIV (GGAS, FKR, FCFKR, LAT, TF, XF, PF, SIG, PHI, PH, A, C, CR).
'Параметры XP и FCFKR включены в список значений.
'Параметр FCFKR = Fc /FCKR является входным.
'Назначение остальных формальных параметров то же, что и в процедуре FULEXP. . Выходной массив С соответствует температуре Тс . Точность определения температуры Тс в процессе последовательных приближений равна 0,001 К.
'Для определения приведенной скорости λ по известным значениям приведенной плотности тока q(λ) и показателя адиабаты К в диапазоне λ приводится процедура LMQ2 ( Q, К ).
'Заголовок процедуры: Real PROCEDURE LMQ2 (Q, К).
'Процедуры FULEXP, C0NDIV, CONNOZ позволяют вычислять параметры истечения из сопел в следующих единицах измерения: Фс - кгс; Рс ,  РН ,  РФ-  кгс/см .
'Для вычисления этих параметров в системе СИ используются аналогичные процедуры FULLSI, C0NDSI,  CONNSI (см. рекомендуемые приложения 2, 3, и 4).
'Примечание. Допускается применение таблиц T-PI-I-S - функций, а
'также кусочно-квадратичной аппроксимации используемых зависимостей. Однако для сопоставления расчетных характеристик двигателей в одинаковых условиях необходимо пользоваться настоящим РТМ.

'УСЛОВНЫЕ ОБОЗНАЧЕНИЯ И ИНДЕКСЫ
'G - массовый расход 
'W - скорость 
'р - давление 
'PI - отношение давлений 
'Т – температура
'i - энтальпия 
'Ср - истинная теплоемкость при постоянном давлении
'К - показатель адиабаты 
'Rг - газовая постоянная
'α - коэффициент избытка воздуха
'qT. - отношение расхода топлива к расходу воздуха
'L0 - стехиометрический коэффициент (количество килограммов воздуха,необходимое для полного сгорания 1 кг топлива) 
'Нц - теплотворная способность топлива
'ηГ - коэффициент полноты сгорания топлива 
'F - площадь
'T = T/1000 масштабированное значение температуры
'Индексом S обозначены параметры при изоэнтропическом истечении.

