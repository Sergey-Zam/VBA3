Option Explicit On
Imports Inventor
Imports Microsoft.Office.Interop

Public Class Form1
    'структура имя параметра-значение параметра
    Private Structure PartParameter
        Public name As String
        Public value As String
    End Structure

    'глобальные переменные
    Dim _loadingForm As New LoadingForm() 'загрузочная форма
    Dim _invApplication As Application = Nothing 'приложение Inventor
    Dim _openFileDialog As New OpenFileDialog 'диалог выбора файла
    Dim _conn As OleDb.OleDbConnection 'подключение к источнику данных
    Dim _listExcel, _listAssembly As New List(Of String)() 'списки для хранения данных из excel и assembly
    Dim _excelCellsRead As String = "F14:G225" 'какие ячейки считывать из excel

    'функция автоматически запускается перед открытием формы. 
    Public Sub New()
        'этот вызов является обязательным для конструктора.
        InitializeComponent()

        'ниже размещается любой инициализирующий код.
        'найти текущий сеанс Inventor (если Inventor не запущен - запустить)
        Try
            Try
                _invApplication = Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
            Catch ex As Exception
                Debug.Print(ex.ToString())
            End Try

            If _invApplication Is Nothing Then
                Dim inventorAppType As Type = Type.GetTypeFromProgID("Inventor.Application")
                _invApplication = Activator.CreateInstance(inventorAppType)
                _invApplication.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

        'добавить столбцы к обоим dgv
        dgvDataFromExcel.ColumnCount = 2
        dgvDataFromAssembly.ColumnCount = 2
        'и задать им ширину
        dgvDataFromExcel.Columns(0).Width = 300
        dgvDataFromExcel.Columns(1).Width = 100
        dgvDataFromAssembly.Columns(0).Width = 300
        dgvDataFromAssembly.Columns(1).Width = 100
    End Sub

    'функция вызывается по нажатию кнопки "Импорт данных из Excel (*.xlsx, *.xls)"
    Private Sub btnImportFromExcel_Click(sender As Object, e As EventArgs) Handles btnGetDataFromExcel.Click
        'выбрать файл excel
        Dim fullName As String = ""
        Try
            _openFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            _openFileDialog.Title = "Open Excel File"
            _openFileDialog.Filter = "Excel Files(2007)|*.xlsx|Excel Files(2003)|*.xls"

            If _openFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                Dim fi As New IO.FileInfo(_openFileDialog.FileName)
                fullName = fi.FullName
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            _conn.Close()
        End Try

        'если получен адрес (не пустой и не Nothing)
        If (Not String.IsNullOrEmpty(fullName)) Then
            _loadingForm.Show() ' долгий процесс - показать загрузочную форму

            Dim exl As New Excel.Application
            Dim exlSheet As Excel.Worksheet

            exl.Workbooks.Open(fullName) 'открыть документ
            tbExcelDirectory.Text = fullName 'заполнить tbExcelDirectory адресом этого документа

            exlSheet = exl.Workbooks(1).Worksheets(1) 'Переходим к первому листу
            Dim array As Object = exlSheet.Range(_excelCellsRead).Value 'Теперь вспомогательный array содержит таблицу из excel
            'закрыть документ excel - больше не нужен
            exl.Quit()
            exlSheet = Nothing
            exl = Nothing

            'все четные не значения array записать в массив evenElems (описания, столбец А), все нечетные значения записать в _listExcel (значения, столбец В)
            'при этом все значения не null и не пустые
            _listExcel.Clear() 'перед заполнением _listExcel надо очистить
            Dim evenElems As New List(Of String)()
            Dim k As Integer = 0
            For Each i As Object In array
                If (Not (i Is Nothing)) Then
                    If (Not (i.ToString = "")) Then
                        If (k Mod 2 = 0) Then 'четное в массив evenElems
                            evenElems.Add(i.ToString)
                        Else 'нечетное в _listExcel
                            _listExcel.Add(i.ToString)
                        End If
                    End If
                End If
                k += 1
            Next

            'записать полученные данные (вспомогательный array) в dgvDataFromExcel
            dgvDataFromExcel.Rows.Clear() 'перед этим dgv надо очистить
            k = 0
            For Each s As String In _listExcel
                dgvDataFromExcel.Rows.Add(evenElems(k), _listExcel(k))
                k += 1
            Next

            lblCountOfExcel.Text = _listExcel.Count

            _loadingForm.Hide() ' закрыть загрузочную форму
        End If
    End Sub

    'функция вызывается по нажатию кнопки "Получить данные из сборки"
    Private Sub btnGetDataFromAssembly_Click(sender As Object, e As EventArgs) Handles btnGetDataFromAssembly.Click
        'выбрать фаил сборки
        Dim fullName As String = ""
        Try
            _openFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            _openFileDialog.Title = "Open Assembly File"
            _openFileDialog.Filter = "Фаил сборки|*.iam"

            If _openFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                Dim fi As New IO.FileInfo(_openFileDialog.FileName)
                fullName = fi.FullName
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            _conn.Close()
        End Try

        'если получен адрес (не пустой и не Nothing)
        If (Not String.IsNullOrEmpty(fullName)) Then
            _loadingForm.Show() ' долгий процесс - показать загрузочную форму
            'открыть существующий документ сборки по указанному пути
            'переменная документа сборки, инициализировать
            Dim asmDoc As Document = _invApplication.Documents.Open(fullName)

            'заполнить tbAssemblyDirectory адресом этого документа
            tbAssemblyDirectory.Text = fullName

            'если ошибка, и не открыто ни 1 документа
            If _invApplication.Documents.Count = 0 Then
                MsgBox("Не открыт документ. Откройте документ сборки (.iam)")
                Return 'выход из функции обработчика кнопки
            End If

            'если тип открытого документа не сборка
            If _invApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MsgBox("Неправильный тип документа. Откройте документ сборки (.iam)")
                Return
            End If

            'перед проходом по всем документам: part, assembly, draw, необходимо очистить список _listAssembly - в нем могут быть старые значения
            _listAssembly.Clear()
            dgvDataFromAssembly.Rows.Clear() 'dgv так же надо очистить

            '---parts---
            'переменные для документов всех проверяемых деталей
            Dim part001Doc As Document = Nothing
            Dim part002Doc As Document = Nothing
            Dim part003Doc As Document = Nothing
            Dim part004Doc As Document = Nothing
            Dim part005Doc As Document = Nothing
            Dim part006Doc As Document = Nothing
            Dim part007Doc As Document = Nothing

            'необходимо инициализировать эти переменные
            'пройти по всем связанным part-документам сборки. (в обратном порядке - так документы будут извлечены из inventor с первого по последний)
            For i = asmDoc.AllReferencedDocuments.Count To 1 Step -1
                Dim currentDoc As Document = asmDoc.AllReferencedDocuments(i) 'получен связаный со сборкой документ детали

                'определить имя (обозначение) детали, и по нему инициализировать заданные ранее переменные деталей
                Select Case currentDoc.DisplayName
                    Case "05.01.001"
                        part001Doc = currentDoc
                    Case "05.01.002"
                        part002Doc = currentDoc
                    Case "05.01.003"
                        part003Doc = currentDoc
                    Case "05.01.004"
                        part004Doc = currentDoc
                    Case "05.01.005"
                        part005Doc = currentDoc
                    Case "05.01.006"
                        part006Doc = currentDoc
                    Case "05.01.007"
                        part007Doc = currentDoc
                End Select
            Next

            'Получить данные каждой детали
            getPart001(part001Doc)
            getPart002(part002Doc)
            getPart003(part003Doc)
            getPart004(part004Doc)
            getPart005(part005Doc)
            getPart006(part006Doc)
            getPart007(part007Doc)

            '---assembly 's---
            'Получить данные каждой сборки
            getAsm(asmDoc)

            '---drawings---
            'переменные для документов всех проверяемых чертежей
            Dim drawing007Doc As Document = Nothing
            'заполнить переменные чертежей
            Dim drawing007DocFullFileName As String = findDrawingFullFileNameForDocument(part007Doc) 'найти путь к чертежу детали
            'если путь к чертежу найден, инициализировать переменную чертежа и открыть чертеж
            If drawing007DocFullFileName IsNot "" Then
                drawing007Doc = _invApplication.Documents.Open(drawing007DocFullFileName) 'открыть чертеж
            End If

            'Получить данные каждого чертежа
            getDrawing007(drawing007Doc)

            lblCountOfAssembly.Text = _listAssembly.Count

            _loadingForm.Hide() ' закрыть загрузочную форму
        End If
    End Sub

    'функция по нажатию кнопки "сравнить"
    Private Sub btnCompare_Click(sender As Object, e As EventArgs) Handles btnCompare.Click
        Dim correct As Boolean = True
        Dim errors As Integer = 0
        If (_listExcel.Count = 0 Or _listAssembly.Count = 0) Then
            MsgBox("Сначала необходимо считать данные из эксель и сборки")
            Return
        ElseIf Not (_listExcel.Count = _listAssembly.Count) Then
            MsgBox("Ошибка: количество записей не совпадает")
            Return
        Else
            Dim style_wrong As New DataGridViewCellStyle
            style_wrong.BackColor = Drawing.Color.LightCoral
            Dim style_right As New DataGridViewCellStyle
            style_right.BackColor = Drawing.Color.LightGreen

            For i = 0 To (_listExcel.Count - 1)
                If Not (_listExcel(i) = _listAssembly(i)) Then
                    'если ошибка
                    correct = False
                    errors += 1
                    dgvDataFromAssembly.Rows(i).DefaultCellStyle = style_wrong
                    dgvDataFromExcel.Rows(i).DefaultCellStyle = style_wrong
                Else
                    'если правильно
                    dgvDataFromAssembly.Rows(i).DefaultCellStyle = style_right
                    dgvDataFromExcel.Rows(i).DefaultCellStyle = style_right

                    'dgvDataFromAssembly.Rows(i).Cells(0).Style = style_right
                End If
            Next
        End If

        If (correct = True) Then
            MsgBox("Не найдено ни одной ошибки")
        Else
            MsgBox("Найдено " & errors & " ошибок")
        End If
    End Sub

    'функция по нажатию кнопки "очистить обе таблицы"
    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        Dim result As Integer = MessageBox.Show("Вы действительно хотите очистить обе таблицы?", "Подтверждение действия", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            'отмена: ничего не делать
        ElseIf result = DialogResult.No Then
            'нет: ничего не делать
        ElseIf result = DialogResult.Yes Then
            'да: действие подтверждено
            'dgvDataFromExcel.DataSource = Nothing
            dgvDataFromExcel.Rows.Clear()
            dgvDataFromAssembly.Rows.Clear()
            _listExcel.Clear()
            _listAssembly.Clear()
            tbExcelDirectory.Clear()
            tbAssemblyDirectory.Clear()
            lblCountOfExcel.Text = ""
            lblCountOfAssembly.Text = ""
        End If
    End Sub

    'Вспомогательные функции
    'Вспомогательная функция найти чертеж к документу: сборке (assembly) или детали (part). если чертеж не найден, возвращает пустую строку: ""
    Private Function findDrawingFullFileNameForDocument(ByVal doc As Document) As String
        Dim fullFilename As String = doc.FullFileName

        'переменная drawingFilename будет хранить полное имя чертежа для сборки / детали
        Dim drawingFilename As String = ""

        ' Extract the path from the full filename.
        Dim path As String = Microsoft.VisualBasic.Left$(fullFilename, InStrRev(fullFilename, "\"))

        ' Extract the filename from the full filename.
        Dim filename As String = Microsoft.VisualBasic.Right$(fullFilename, Len(fullFilename) - InStrRev(fullFilename, "\"))

        ' Replace the extension with "dwg"
        filename = Microsoft.VisualBasic.Left$(filename, InStrRev(filename, ".")) & "dwg"
        ' Find if the drawing exists.
        drawingFilename = _invApplication.DesignProjectManager.ResolveFile(path, filename)

        ' Check the result.
        If drawingFilename = "" Then
            ' Try again with idw extension.
            filename = Microsoft.VisualBasic.Left$(filename, InStrRev(filename, ".")) & "idw"
            ' Find if the drawing exists.
            drawingFilename = _invApplication.DesignProjectManager.ResolveFile(path, filename)
        End If

        ' Display the result.
        If drawingFilename <> "" Then
            Return drawingFilename
        Else
            MsgBox("No drawing was found for """ & doc.FullFileName & """")
            Return drawingFilename
        End If
    End Function

    'вспомогательная функция: проверить видимость 2d эскизов и объектов вспомогательной геометрии (плоскости, оси, точки). true - они все невидимы, false - есть как минимум 1 видимый объект
    Private Function isOriginsInvisible(ByVal oDoc As Document) As Boolean
        Dim isInvisible As Boolean = True

        ' получть все 2d эскизы детали и проверить их видимость
        Dim oSketches As PlanarSketches = oDoc.ComponentDefinition.Sketches
        For Each oSketch In oSketches
            If oSketch.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkPlanes collection (все плоскости документа)
        For Each oWorkPlane In oDoc.ComponentDefinition.WorkPlanes
            If oWorkPlane.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkAxes collection (все оси документа)
        For Each oWorkAxe In oDoc.ComponentDefinition.WorkAxes
            If oWorkAxe.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkPoints collection (все точки документа)
        For Each oWorkPoint In oDoc.ComponentDefinition.WorkPoints
            If oWorkPoint.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkSurfaces collection (все поверхности(?) документа)
        'For Each oWorkSurface In oDoc.ComponentDefinition.WorkSurfaces
        '    If oWorkSurface.Visible = False Then
        '        MsgBox(oWorkSurface.Name & " Visible false: ok")
        '    Else
        '        MsgBox(oWorkSurface.Name & "Visible true: not ok")
        '    End If
        'Next        
        Return isInvisible
    End Function

    'вспомогательная функция: получить таблицу свойств детали (доступ к таблице параметров)
    Private Function getParametersFromPart(ByVal partDoc As Document) As List(Of PartParameter)
        Dim allParams As Parameters = partDoc.ComponentDefinition.Parameters
        Dim listOfParameters As New List(Of PartParameter)() 'список параметров документа

        If allParams.Count > 0 Then
            For Each param As Parameter In allParams
                Dim partParameter As PartParameter
                partParameter.name = param.Name
                partParameter.value = (param.ModelValue * 10).ToString
                listOfParameters.Add(partParameter)
            Next
        End If

        Return listOfParameters
    End Function

    'вспомогательная функция: вернуть из структуры типа PartParameter значение по имени
    Private Function findValueInPartParamListByName(ByVal name As String, ByVal list As List(Of PartParameter)) As String
        Dim value As String = ""

        For Each elem As PartParameter In list
            If elem.name = name Then
                'взять значение по модулю
                elem.value = Math.Abs(CDec(elem.value))
                value = elem.value.ToString
                Exit For
            End If
        Next

        Return value
    End Function

    'Функции получения данных из деталей, сборки, чертежей
    Private Sub getPart001(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d10", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø12", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d0", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер R10", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d3", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d26", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба (в отверстии)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d18", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде сверху)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d5", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде слева, ширина)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d2", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде слева, высота)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (центр отверстия Ø12 на в центре дуги R10)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (симметрия на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (симметрия на виде сверху)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (симметрия на виде слева)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getPart002(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d0", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø68", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба (в отверстии)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d20", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d20", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d20", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d20", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие (строится инструментом отверстие)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d10", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие Ø6", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d15", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие глубина 5", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие 4 экземпляра (круговым массивом)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (отверстия симметричны на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getPart003(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d2", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø48", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø30", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d3", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d26", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер R5", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d11", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d9", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d6", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d19", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d14", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø36", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба наружная", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d24", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба наружная", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d27", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде слева)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d7", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø20", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d8", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие Ø8", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d21", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер □18", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (симметрия на виде сверху, отверстие Ø36х120)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (отверстия Ø8 симметричны на виде слева)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (бобышки Ø20 симметричны на виде слева)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (□18 ориентирован относительно осей симметрии)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (плоскости, касательные к цилиндрам - 4 случая)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getPart004(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d0", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø12", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d3", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d3", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getPart005(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d0", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø25", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø12,5", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d2", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (соосность Ø25 и Ø12,5)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getPart006(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d14", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø64", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d15", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø60", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d2", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d3", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d5", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d7", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø36", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d14", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø64 (на виде слева)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d6", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø21", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба (в отверстии)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d20", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø50", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d13", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d19", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер R0,4", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d37", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (проточка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d18", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (проточка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d16", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер угловой (проточка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d11", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø38", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d12", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d43", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d40", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие (строится инструментом отверстие)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d30", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие Ø6", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d35", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие глубина 5", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Отверстие 4 экземпляра (круговым массивом)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d29", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди, положение отверстий)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (все цилиндрические и конические поверхности, кроме 4 отв., соосны)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getPart007(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        _listAssembly.Add(oPropSetDTP.Item("Part Number").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Обозначение", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Description").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Значение параметра Наименование", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение материала (с чертежа)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(oPropSetDTP.Item("Material").Value) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Присвоение представления", _listAssembly.Last) 'записать value в dgvAssembly

        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        _listAssembly.Add("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения
        dgvDataFromAssembly.Rows.Add("Проверка даты создания (изменения) файла", _listAssembly.Last) 'записать value в dgvAssembly

        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        For Each SrfBod In SrfBods
            _listAssembly.Add(SrfBod.IsSolid) 'доб. value в _listAssembly
        Next
        dgvDataFromAssembly.Rows.Add("Деталь твердотельная (не поверхности)", _listAssembly.Last) 'записать value в dgvAssembly

        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            _listAssembly.Add(True) 'true - да, из одного
        Else
            _listAssembly.Add(False) 'false - нет, не из одного
        End If
        dgvDataFromAssembly.Rows.Add("Деталь состоит из одного твердого тела", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый
        dgvDataFromAssembly.Rows.Add("Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d5", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø20", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d6", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø33", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d2", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d3", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d9", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø7,7", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d11", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (проточка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d10", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (проточка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер угловой (проточка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d12", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер R0,8", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d12", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер R0,8", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба наружная", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d18", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер □18", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d13", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add(findValueInPartParamListByName("d13", listOfParameters)) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (□18 ориентирован относительно осей симметрии)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия (все цилиндрические и конические поверхности соосны)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getAsm(ByVal asmDoc As Document)
        Dim occ As ComponentOccurrence

        'Деталь 05.01.003 Корпус закреплена (0 степеней свободы)
        Dim result As String = "EMPTY VALUE"
        For Each occ In asmDoc.ComponentDefinition.Occurrences 'occ - свойства part document (1..n) В assembly, их (документов) перебор
            'если деталь "05.01.003"
            If (occ.Name = "05.01.003" & ":1") Then
                result = occ.Grounded 'true - да, деталь закреплена, false - нет, деталь не закреплена
                Exit For
            End If
        Next
        _listAssembly.Add(result) 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Деталь 05.01.003 Корпус закреплена (0 степеней свободы)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.004 Ось (ось 004 совпадает с осью отверстия 001)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.004 Ось (ограничение перемещения 004 вдоль оси)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.005 Ролик (ось 005 совпадает с осью отверстия 001 или с осью 004)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.005 Ролик (ограничение перемещения 005 вдоль оси)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.007 Шток (ось резьбы 001 совпадает с осью резьбы 007)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.007 Шток (плоскость 001 совпадает с плоскостью 007)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.001 Вилка и 05.01.007 Шток (угол поворота 001 относительно 007 указан)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.003 Корпус и 05.01.007 Шток (ось цилиндра 003 совпадает с осью цилиндра 007, или две плоскости □18 на 003 и на 007 совпадают)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.003 Корпус и 05.01.007 Шток (ось цилиндра 003 совпадает с осью цилиндра 007 и по одной плоскости □18 совпадают или указан угол поворота, или две плоскости □18 на 003 и на 007 совпадают)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.003 Корпус и 05.01.007 Шток (007 упирается в 003 буртиком Ø33 (совпадение плоскостей))", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.002 Гайка и 05.01.003 Корпус (цилиндры 002 соосны резьбе 003)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.002 Гайка и 05.01.003 Корпус (плоскость (торец) гайки 002 находится в указанной позиции относительно резьбы на 003 (координата вычисляема))", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.002 Гайка и 05.01.003 Корпус (угловое положение 002 относительно 003)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.006 Стакан и 05.01.002 Гайка (или 003 Корпус, или 007 Шток) (соосность цилиндрических поверхностей)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.006 Стакан и 05.01.002 Гайка (совпадение плоскостей (торцев) 006 Стакан и 002 Гайка)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Зависимости в паре 05.01.006 Стакан и 05.01.002 Гайка (угловое положение 006 относительно 002)", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub

    Private Sub getDrawing007(ByVal drawingDoc As Document)
        Dim oSheet As Sheet = drawingDoc.ActiveSheet 'лист чертежа

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Выбор ориентации детали (протяженная, большая часть поверхностей цилиндры)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Выбор главного вида (в данном случае учитывается ориентация осей)", _listAssembly.Last) 'записать value в dgvAssembly

        'Выбор формата листа
        Dim result As String = "EMPTY VALUE"
        If oSheet.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize Then
            result = "А4"
        ElseIf oSheet.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize Then
            result = "А3"
        Else
            result = "Другой формат"
        End If
        _listAssembly.Add(result)
        dgvDataFromAssembly.Rows.Add("Выбор формата листа", _listAssembly.Last)

        'Выбор масштаба главного вида
        result = "EMPTY VALUE"
        Dim oPropSets As PropertySets = drawingDoc.PropertySets
        Dim oPropSetGOST As PropertySet = oPropSets.Item("Свойства ГОСТ")
        result = oPropSetGOST.Item("Масштаб").Value
        _listAssembly.Add(result)
        dgvDataFromAssembly.Rows.Add("Выбор масштаба главного вида", _listAssembly.Last)

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Как разместить главный вид на листе формата А4? Сделать разрыв", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("На главном виде отобразается плоскость □18. Перекрестие", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Главный вид □18. Перекрестие. Отдельный эскиз.", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Главный вид □18. Перекрестие. Отдельный эскиз. Вес линий.", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Деталь симметрична на главном виде. Наличие осевой во всю длину проекции +5 мм за пределы контура.", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (габаритный)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø20", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø33", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный (на виде спереди)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Резьба наружная", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Фаска", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер □18 не может быть указан -> Разрез", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Разрез находится за пределами листа -> Отключение выравнивания", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Разрез помещён на свободное пространство листа", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Геометрия на виде симметрична относительно двух осей (+окружность) -> 2 осевых линии", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Разрез. Масштаб (1:1) совпадает с масштабом главного вида -> удаляем (1:1) после А-А", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер □18", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размеры проточки на главном виде не удастся показать -> Выносной вид", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("По умолчанию наименование вида B (лат.) -> меняем на Б", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("По умолчанию масштаб вида Б устанавливается равным 2:1, для размещения всех необходимых размеров изменяем его на 4:1", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Чтобы показать линии перехода в местах скруглений разрываем связь стиля вида Б с главным видом", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("В параметрах отображения вида Б включаем линии перехода", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер R0,8 (2 места)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер Ø7,7", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный на виде Б (2,49 округлить до 2,5)", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Размер линейный на виде Б", _listAssembly.Last) 'записать value в dgvAssembly

        _listAssembly.Add("EMPTY VALUE") 'доб. value в _listAssembly
        dgvDataFromAssembly.Rows.Add("Заполнение основной надписи", _listAssembly.Last) 'записать value в dgvAssembly
    End Sub
End Class