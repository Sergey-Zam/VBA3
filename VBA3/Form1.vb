Option Explicit On
Imports Inventor
Imports Microsoft.Office.Interop
Imports System.IO

Public Class Form1
    'структура: исходные данные из excel
    Private Structure AspectData
        'поля, заполняемые через Excel:
        Public text As String 'текст аспекта
        Public valueFromExcel As String 'значение аспекта из Excel
        Public weight As Double 'вес (значимость) аспекта
        Public tolerance As Double 'допустимое отклонение аспекта
        Public interpretation As String 'интрпретация аспекта
        Public comment As String 'комментарий к аспекту

        'поля, заполняемые через Inventor
        Public valueFromInventor As String 'значение аспекта из Inventor
        Public delta As Double 'имеющееся отклонение аспекта
    End Structure

    'структура имя параметра-значение параметра
    Private Structure PartParameter
        Public name As String
        Public value As String
    End Structure

    'глобальные переменные
    Dim _invApplication As Application = Nothing 'приложение Inventor
    Dim _openFileDialog As New OpenFileDialog 'диалог выбора файла    
    Dim _conn As OleDb.OleDbConnection 'подключение к источнику данных
    Dim _listAspects As New List(Of AspectData)() 'список для хранения всех данных: из excel, из inventor
    Dim _excelCellsRead As String = "F14:K232" 'какие ячейки считывать из excel
    Dim _countOfExcelСolumns = 6 'количество столбцов, берущих данные из Excel
    Dim _counterForInventorAspects = 0 'счетчик, увеличивающийся при занесении записей из Inventor в _listAspects

    'функция автоматически запускается перед открытием формы. 
    Public Sub New()
        'этот вызов является обязательным для конструктора.
        InitializeComponent()

        'ниже размещается любой инициализирующий код.
        'увеличить форму на весь экран
        Me.WindowState = FormWindowState.Maximized
        'Me.FormBorderStyle = FormBorderStyle.None
    End Sub

    'функция запускается, как только форма загружена.
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        showLoadingMessage("Поиск/запуск Inventor") ' долгий процесс - показать сообщение загрузки

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

        'добавить столбцы к dgvAspects
        dgvAspects.ColumnCount = 8
        'и задать им заголовки и ibhbye
        Dim standartWindth = dgvAspects.Width / 8
        dgvAspects.Columns(0).HeaderText = "Аспект"
        dgvAspects.Columns(0).Width = standartWindth * 2
        dgvAspects.Columns(1).HeaderText = "Значение (из Excel)"
        dgvAspects.Columns(1).Width = standartWindth * 1.25
        dgvAspects.Columns(2).HeaderText = "Вес аспекта"
        dgvAspects.Columns(2).Width = standartWindth * 0.5
        dgvAspects.Columns(3).HeaderText = "Допустимое отклонение, точность (%)"
        dgvAspects.Columns(3).Width = standartWindth
        dgvAspects.Columns(4).HeaderText = "Интрпретация"
        dgvAspects.Columns(4).Width = standartWindth * 0.5
        dgvAspects.Columns(5).HeaderText = "Комментарий"
        dgvAspects.Columns(5).Width = standartWindth * 0.5
        dgvAspects.Columns(6).HeaderText = "Значение (из Inventor)"
        dgvAspects.Columns(6).Width = standartWindth * 1.25
        dgvAspects.Columns(7).HeaderText = "Имеющееся отклонение (%)"
        dgvAspects.Columns(7).Width = standartWindth


        lblLoading.Visible = False 'закрыть сообщение загрузки
    End Sub

    'функция вызывается по нажатию кнопки "Импорт данных из Excel (*.xlsx, *.xls)"
    Private Sub btnImportFromExcel_Click(sender As Object, e As EventArgs) Handles btnGetDataFromExcel.Click
        'выбрать файл excel
        Dim _openFileDialog As New OpenFileDialog
        Dim fullName As String = ""
        Try
            '_openFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            _openFileDialog.RestoreDirectory = True
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
            showLoadingMessage("Загрузка данных из Excel") ' долгий процесс - показать сообщение загрузки

            Dim exl As New Excel.Application
            Dim exlSheet As Excel.Worksheet

            exl.Workbooks.Open(fullName) 'открыть документ
            tbExcelDirectory.Text = fullName 'заполнить tbExcelDirectory адресом этого документа

            exlSheet = exl.Workbooks(1).Worksheets(1) 'Переходим к первому листу

            Dim a(,) As Object
            a = exlSheet.Range(_excelCellsRead).Value 'Теперь вспомогательный массив a содержит таблицу из excel

            'закрыть документ excel - больше не нужен
            exl.Quit()
            exlSheet = Nothing
            exl = Nothing

            Dim countOfA As Integer = a.Length 'количество всех элементов массива a
            Dim countOfRowsInA As Integer = countOfA / _countOfExcelСolumns 'количество строк в a

            'проход по всем строкам массива a, что бы переписать их в _listAspects. записаны будут только не пустые значения
            _listAspects.Clear() 'перед заполнением _listAspects надо очистить
            lblCountOfExcel.Text = "0"
            For i As Integer = 1 To countOfRowsInA
                If a(i, 1) IsNot Nothing Then 'если поле текста аспекта существует
                    If a(i, 1) IsNot "" Then 'если поле текста аспекта не пустое
                        Dim ed As AspectData = Nothing

                        ed.text = a(i, 1)
                        ed.valueFromExcel = a(i, 2)

                        Try
                            ed.weight = Convert.ToDouble(a(i, 3))
                        Catch ex As Exception
                            ed.weight = 0
                        End Try

                        Try
                            ed.tolerance = Convert.ToDouble(a(i, 4))
                        Catch ex As Exception
                            ed.tolerance = 0
                        End Try

                        ed.interpretation = a(i, 5)
                        ed.comment = a(i, 6)

                        'все значения из excel добавлены, но нельзя оставлять оставшиеся поля со значением Nothing
                        ed.valueFromInventor = ""
                        ed.delta = 0

                        _listAspects.Add(ed)
                    End If
                End If
            Next

            'теперь нужно пройти весь _listAspects, и каждый его элемент (содержащий 8 полей) переписать в dgv
            dgvAspects.Rows.Clear() 'перед заполнением dgv надо очистить
            For Each d As AspectData In _listAspects
                dgvAspects.Rows.Add(d.text, d.valueFromExcel, d.weight, d.tolerance, d.interpretation, d.comment, d.valueFromInventor, d.delta)
            Next

            lblCountOfExcel.Text = _listAspects.Count

            lblLoading.Visible = False 'закрыть сообщение загрузки
        End If
    End Sub

    'функция вызывается по нажатию кнопки "Получить данные из сборки"
    Private Sub btnGetDataFromAssembly_Click(sender As Object, e As EventArgs) Handles btnGetDataFromAssembly.Click
        If _listAspects.Count = 0 Then
            MsgBox("Сначала необходимо считать данные из Excel")
            Return 'выход из функции обработчика кнопки
        End If

        'выбрать фаил сборки
        Dim fullName As String = ""
        Try
            '_openFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            _openFileDialog.RestoreDirectory = True
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
            showLoadingMessage("Получение данных из Inventor") ' долгий процесс - показать сообщение загрузки

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

            'перед проходом по всем документам: part, assembly, draw, необходимо очистить поля, получаемые из Inventor в _listAspects - в них могут быть старые значения
            For Each aspekt As AspectData In _listAspects
                aspekt.valueFromInventor = ""
                aspekt.delta = 0
            Next
            'так же нужно сбросить счетчик в изначальное состояние
            _counterForInventorAspects = 0
            lblCountOfAssembly.Text = "0"

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
            If part001Doc IsNot Nothing Then
                getPart001(part001Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

            If part002Doc IsNot Nothing Then
                getPart002(part002Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

            If part003Doc IsNot Nothing Then
                getPart003(part003Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

            If part004Doc IsNot Nothing Then
                getPart004(part004Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

            If part005Doc IsNot Nothing Then
                getPart005(part005Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

            If part006Doc IsNot Nothing Then
                getPart006(part006Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

            If part007Doc IsNot Nothing Then
                getPart007(part007Doc)
            Else
                MsgBox("Ошибка: деталь не найдена")
            End If

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
            If part007Doc IsNot Nothing Then
                getDrawing007(drawing007Doc)
            Else
                MsgBox("Ошибка: деталь не найдена, ее чертеж не может быть получен")
            End If

            'Все данные из Inventor получены в _listAspects. dgv надо обновить
            dgvAspects.Rows.Clear()
            For Each d As AspectData In _listAspects
                dgvAspects.Rows.Add(d.text, d.valueFromExcel, d.weight, d.tolerance, d.interpretation, d.comment, d.valueFromInventor, d.delta)
            Next

            lblCountOfAssembly.Text = _counterForInventorAspects

            'ТЕСТОВАЯ функция эскизов, потом удалить
            'For Each pd As PartDocument In asmDoc.AllReferencedDocuments
            '    getInfoAboutSketches(pd)
            'Next

            lblLoading.Visible = False 'закрыть сообщение загрузки
        End If
    End Sub

    'функция по нажатию кнопки "сравнить"
    Private Sub btnCompare_Click(sender As Object, e As EventArgs) Handles btnCompare.Click
        Dim correct As Boolean = True
        Dim errors As Integer = 0
        Dim total_points As Double = 0
        If (_listAspects.Count = 0 Or lblCountOfAssembly.Text = 0) Then
            MsgBox("Сначала необходимо считать данные из эксель и сборки")
            Return 'выход из функции обработчика кнопки
        End If

        Dim style_wrong As New DataGridViewCellStyle
        style_wrong.BackColor = Drawing.Color.LightCoral
        Dim style_full_right As New DataGridViewCellStyle
        style_full_right.BackColor = Drawing.Color.LightGreen
        Dim style_right As New DataGridViewCellStyle
        style_right.BackColor = Drawing.Color.DarkSeaGreen

        For i = 0 To (_listAspects.Count - 1)
            If (_listAspects(i).valueFromExcel = _listAspects(i).valueFromInventor) Then
                'если значения точно совпадают, ответ верный
                dgvAspects.Rows(i).DefaultCellStyle = style_full_right
                total_points += _listAspects(i).weight
            Else
                Dim valueFromInventor, valueFromExcel As Double
                'значения не совпадают, необходимо проверить точность (если возможно)
                If Double.TryParse(_listAspects(i).valueFromInventor, valueFromInventor) And Double.TryParse(_listAspects(i).valueFromExcel, valueFromExcel) Then
                    'если допустимое отклонение больше, чем текущее отклонение, ответ верный
                    If (_listAspects(i).tolerance > _listAspects(i).delta) Then
                        'ответ верный, в пределах отклонения (но не точный)
                        dgvAspects.Rows(i).DefaultCellStyle = style_right
                        total_points += _listAspects(i).weight
                    Else
                        'ответ не верный
                        correct = False
                        errors += 1
                        dgvAspects.Rows(i).DefaultCellStyle = style_wrong
                    End If
                Else
                    'значение - не число, нет смысла проверять точность, ответ неверный
                    correct = False
                    errors += 1
                    dgvAspects.Rows(i).DefaultCellStyle = style_wrong
                End If
            End If
        Next

        If (correct = True) Then
            MsgBox("Не найдено ни одной ошибки" & vbCrLf & "Всего набрано баллов: " & total_points)
        Else
            MsgBox("Найдено " & errors & " ошибок" & vbCrLf & "Всего набрано баллов: " & total_points)
        End If
    End Sub

    'функция по нажатию кнопки "очистить таблицу"
    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        Dim result As Integer = MessageBox.Show("Вы действительно хотите очистить таблицу?", "Подтверждение действия", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            'отмена: ничего не делать
        ElseIf result = DialogResult.OK Then
            'да: действие подтверждено
            'dgvDataFromExcel.DataSource = Nothing
            dgvAspects.Rows.Clear()
            _listAspects.Clear()
            tbExcelDirectory.Clear()
            tbAssemblyDirectory.Clear()
            lblCountOfExcel.Text = "0"
            lblCountOfAssembly.Text = "0"
        End If
    End Sub

    'функция по нажатию кнопки "экспорт в эксель"
    Private Sub btnExportToExcel_Click(sender As Object, e As EventArgs) Handles btnExportToExcel.Click
        'проверка, имеются ли данные в dgv
        If (dgvAspects.Rows.Count = 0) Then
            MsgBox("Таблица пуста, экспорт невозможен")
            Return 'выход из функции обработчика кнопки
        End If

        'выбрать место сохрания файла excel
        Dim saveFileDialog As New SaveFileDialog 'диалог выбора места сохранения файла
        'saveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        saveFileDialog.RestoreDirectory = True
        saveFileDialog.Filter = "Excel Files(2007)|*.xlsx|Excel Files(2003)|*.xls"
        saveFileDialog.Title = "Save Excel File"
        saveFileDialog.ShowDialog()

        'если получена директория
        If (Not String.IsNullOrEmpty(saveFileDialog.FileName)) Then
            'Создание dataset для экспорта
            Dim dset As New DataSet
            dset.Tables.Add() 'добавить таблицу

            'добавление столбцов в эту таблицу
            For i As Integer = 0 To dgvAspects.ColumnCount - 1
                dset.Tables(0).Columns.Add(dgvAspects.Columns(i).HeaderText)
            Next

            'добавление строк в эту таблицу
            Dim dr1 As DataRow
            For i As Integer = 0 To dgvAspects.RowCount - 1
                dr1 = dset.Tables(0).NewRow
                For j As Integer = 0 To dgvAspects.Columns.Count - 1
                    dr1(j) = dgvAspects.Rows(i).Cells(j).Value
                Next
                dset.Tables(0).Rows.Add(dr1)
            Next

            Dim exl As New Excel.Application
            Dim exlBook As Excel.Workbook
            Dim exlSheet As Excel.Worksheet

            exlBook = exl.Workbooks.Add()
            exlSheet = exlBook.ActiveSheet()

            Dim dt As DataTable = dset.Tables(0)
            Dim dc As DataColumn
            Dim dr As DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                exl.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    exl.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next

            exlSheet.Columns.AutoFit()
            Dim strFileName As String = saveFileDialog.FileName
            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try

            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            exlBook.SaveAs(strFileName)
            exl.Workbooks.Open(strFileName)
            exl.Visible = True
        End If
    End Sub

    'функция по закрытию формы
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim result As Integer = MessageBox.Show("Вы действительно хотите выйти?", "Подтверждение действия", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            'отмена: не закрывать форму
            e.Cancel = True
        ElseIf result = DialogResult.OK Then
            'подтверждение: закрыть форму
            e.Cancel = False
        End If
    End Sub

    'Вспомогательные функции
    'вспомогательная функция: показать сообщение загрузки
    Private Sub showLoadingMessage(ByVal text)
        lblLoading.Text = text
        lblLoading.Top = (Me.ClientSize.Height / 2) - (lblLoading.Height / 2)
        lblLoading.Left = (Me.ClientSize.Width / 2) - (lblLoading.Width / 2)
        lblLoading.Visible = True
    End Sub

    'вспомогательная функция найти чертеж к документу: сборке (assembly) или детали (part). если чертеж не найден, возвращает пустую строку: ""
    Private Function findDrawingFullFileNameForDocument(ByVal doc As Document) As String
        Try
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
        Catch ex As Exception
            MsgBox("Ошибка: невозможно найти чертеж для документа" & vbCrLf & ex.ToString)
            Return ""
        End Try
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
        Dim listOfParameters As New List(Of PartParameter)() 'список параметров документа

        Dim allParams As Parameters = partDoc.ComponentDefinition.Parameters
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
                If CDec(elem.value) < 0 Then
                    elem.value = Math.Abs(CDec(elem.value))
                End If

                value = elem.value.ToString
                Exit For
            End If
        Next

        Return value
    End Function

    'вспомогательная функция: добавить в структуру типа AspectData значение с заполненными столбцами, получаемыми из Inventor
    Private Sub addInventorValuesInAspectDataList(ByVal value As String)
        Dim aspect As AspectData = Nothing
        aspect = _listAspects(_counterForInventorAspects) 'получить текущее значение счетчика записей

        aspect.valueFromInventor = value
        'подсчет дельты (формула?)
        Try
            Dim max_value As Double = 0
            If aspect.valueFromExcel > aspect.valueFromInventor Then
                aspect.delta = (Math.Abs(aspect.valueFromExcel - aspect.valueFromInventor) / aspect.valueFromExcel) * 100
            ElseIf aspect.valueFromExcel < aspect.valueFromInventor Then
                aspect.delta = (Math.Abs(aspect.valueFromExcel - aspect.valueFromInventor) / aspect.valueFromInventor) * 100
            Else
                aspect.delta = 0
            End If
        Catch ex As Exception
            aspect.delta = 0
        End Try

        'новое значение aspect, включающее данные из Inventor, нужно вставить вместо старого значения
        _listAspects(_counterForInventorAspects) = aspect

        'увеличить счетчик считанных свойств из Inventor
        _counterForInventorAspects += 1
    End Sub

    'вспомогательная функция: получение первых восьми одинаковых параметров для детали (объединение одинаковых параметров)
    Private Sub getFirstSameParametersOfPart(ByVal partDoc As Document)
        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        ' Get the Inventor Summary Information property set.
        'Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information")
        ' Get the Inventor Document Summary Information property set.
        'Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information")
        ' Get the Design Tracking Properties property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties")

        '"Значение параметра Обозначение"
        addInventorValuesInAspectDataList(oPropSetDTP.Item("Part Number").Value)

        '"Значение параметра Наименование"
        addInventorValuesInAspectDataList(oPropSetDTP.Item("Description").Value)

        '"Присвоение материала (с чертежа)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Присвоение представления"
        addInventorValuesInAspectDataList(oPropSetDTP.Item("Material").Value)

        '"Проверка даты создания (изменения) файла"
        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        addInventorValuesInAspectDataList("Создан: " & f.DateCreated.ToString & " Изменен: " & f.DateLastModified.ToString) 'дата создания и дата изменения

        '"Деталь твердотельная (не поверхности)"
        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        Dim b As Boolean
        For Each SrfBod In SrfBods
            b = SrfBod.IsSolid '? значение последнего surface body ?
        Next
        addInventorValuesInAspectDataList(b)

        '"Деталь состоит из одного твердого тела"
        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            addInventorValuesInAspectDataList(True) 'true - да, из одного
        Else
            addInventorValuesInAspectDataList(False) 'false - нет, не из одного
        End If

        '"Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы"
        addInventorValuesInAspectDataList(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый

        'НОВОЕ Все эскизы детали должны быть полностью определены
        Dim isOk As Boolean = True
        Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
        'пройти по всем эскизам детали
        For Each sketch As Sketch In partDef.Sketches
            'является ли эскиз полностью определенным? если нет, то записывем ошибку
            If sketch.ConstraintStatus <> ConstraintStatusEnum.kFullyConstrainedConstraintStatus Then
                isOk = False
                Exit For
            End If
        Next
        addInventorValuesInAspectDataList(isOk) 'записать true - все эскизы детали полностью определены; false - хотя бы один эскиз детали не полностью определен
    End Sub

    'вспомогательная функция: получить параметры резьбы
    Private Function getThreadsParams(ByVal partDoc As Document) As String
        Dim resultString As String = ""

        Dim fc As Face
        For Each fc In partDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces
            If fc.SurfaceType = Inventor.SurfaceTypeEnum.kCylinderSurface Or fc.SurfaceType = Inventor.SurfaceTypeEnum.kConeSurface Then
                If Not fc.ThreadInfos Is Nothing Then
                    If fc.ThreadInfos.Count > 0 Then
                        Dim thread As ThreadInfo
                        For Each thread In fc.ThreadInfos
                            resultString = "" ' !пока берется последняя резьба, старые рез-ы очищ.

                            Dim threadDesignation As String = thread.ThreadDesignation 'designation (пример: М10х1.5)
                            threadDesignation = Replace(threadDesignation, "M", "М") 'заменить английскую букву M на русскую
                            threadDesignation = Replace(threadDesignation, "x", "х") 'заменить английскую букву x на русскую
                            threadDesignation = Replace(threadDesignation, ".", ",") 'заменить точку на запятую
                            resultString &= threadDesignation
                            resultString &= "-"

                            If TypeOf thread Is StandardThreadInfo Then
                                resultString &= thread.Class 'class (пример: 6H)
                            End If
                        Next
                    End If
                End If
            End If
        Next

        Return resultString
    End Function

    'Функции получения данных из деталей, сборки, чертежей
    Private Sub getPart001(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '"Размер Ø12"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d10", listOfParameters))

        '"Размер R10"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d0", listOfParameters))

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d3", listOfParameters))

        '"Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d1", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d26", listOfParameters))

        '"Резьба (в отверстии)"
        addInventorValuesInAspectDataList(getThreadsParams(partDoc))

        '"Размер линейный (на виде сверху)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d18", listOfParameters))

        '"Размер линейный (на виде слева, ширина)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d5", listOfParameters))

        '"Размер линейный (на виде слева, высота)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d2", listOfParameters))

        '"Геометрия (центр отверстия Ø12 на в центре дуги R10)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (симметрия на виде спереди)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (симметрия на виде сверху)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (симметрия на виде слева)"
        addInventorValuesInAspectDataList("EMPTY VALUE")
    End Sub

    Private Sub getPart002(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '"Размер Ø68"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d0", listOfParameters))

        '"Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d1", listOfParameters))

        '"Резьба (в отверстии)"
        addInventorValuesInAspectDataList(getThreadsParams(partDoc))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d20", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d20", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d20", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d20", listOfParameters))

        '"Отверстие (строится инструментом отверстие)"
        Dim oHoles As HoleFeatures = partDoc.ComponentDefinition.Features.HoleFeatures
        Dim count As Integer = oHoles.Count 'если > 0, значит используется инструмент отверстие
        If count > 0 Then
            addInventorValuesInAspectDataList(True)
        Else
            addInventorValuesInAspectDataList(False)
        End If

        '"Отверстие Ø6"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d10", listOfParameters))

        '"Отверстие глубина 5"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d15", listOfParameters))

        '"Отверстие 4 экземпляра (круговым массивом)"
        Dim oCircularPatternFeatures As CircularPatternFeatures = partDoc.ComponentDefinition.Features.CircularPatternFeatures
        If oCircularPatternFeatures.Count = 1 Then
            'найден один круговой массив. необходимо получить количество элементов этого массива
            Dim countOfElems As Integer = oCircularPatternFeatures(1).Count.Value
            'если элементов 4 - верно, подходит условиям
            If countOfElems = 4 Then
                addInventorValuesInAspectDataList(True)
            Else
                addInventorValuesInAspectDataList(False)
            End If
        Else
            'круговых массивов нет, или больше одного (и то, и то неверно)
            addInventorValuesInAspectDataList(False)
        End If

        '"Геометрия (отверстия симметричны на виде спереди)"
        addInventorValuesInAspectDataList("EMPTY VALUE")
    End Sub

    Private Sub getPart003(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '"Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d4", listOfParameters))

        '"Размер Ø48"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d2", listOfParameters))

        '"Размер Ø30"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d1", listOfParameters))

        '"Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d3", listOfParameters))

        '"Размер R5"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d26", listOfParameters))

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d11", listOfParameters))

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d9", listOfParameters))

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d6", listOfParameters))

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d19", listOfParameters))

        '"Размер Ø36"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d14", listOfParameters))

        '"Резьба наружная"
        addInventorValuesInAspectDataList(getThreadsParams(partDoc))

        '"Резьба наружная"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d24", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d27", listOfParameters))

        '"Размер линейный (на виде слева)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер Ø20"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d7", listOfParameters))

        '"Отверстие Ø8"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d8", listOfParameters))

        '"Размер □18"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d21", listOfParameters))

        '"Геометрия (симметрия на виде сверху, отверстие Ø36х120)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (отверстия Ø8 симметричны на виде слева)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (бобышки Ø20 симметричны на виде слева)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (□18 ориентирован относительно осей симметрии)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия (плоскости, касательные к цилиндрам - 4 случая)"
        addInventorValuesInAspectDataList("EMPTY VALUE")
    End Sub

    Private Sub getPart004(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '"Размер Ø12"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d0", listOfParameters))

        '"Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d1", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d3", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d3", listOfParameters))
    End Sub

    Private Sub getPart005(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '"Размер Ø25"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d0", listOfParameters)) 'доб. value в _listAssembly

        '"Размер Ø12,5"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d1", listOfParameters)) 'доб. value в _listAssembly

        '"Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d2", listOfParameters)) 'доб. value в _listAssembly

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d4", listOfParameters)) 'доб. value в _listAssembly

        '"Геометрия (соосность Ø25 и Ø12,5)"
        addInventorValuesInAspectDataList("EMPTY VALUE") 'доб. value в _listAssembly      
    End Sub

    Private Sub getPart006(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '«Размер Ø64"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d14", listOfParameters))

        '«Размер Ø60"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d15", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d2", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d3", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d4", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d5", listOfParameters))

        '«Размер Ø36"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d7", listOfParameters))

        '«Размер Ø64 (на виде слева)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d14", listOfParameters))

        '«Размер Ø21"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d6", listOfParameters))

        '«Резьба (в отверстии)"
        addInventorValuesInAspectDataList(getThreadsParams(partDoc))

        '«Размер Ø50"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d20", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d13", listOfParameters))

        '«Размер R0,4"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d19", listOfParameters))

        '«Размер линейный (проточка)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d37", listOfParameters))

        '«Размер линейный (проточка)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d18", listOfParameters))

        '«Размер угловой (проточка)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d16", listOfParameters))

        '«Размер Ø38"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d11", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d12", listOfParameters))

        '«Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d43", listOfParameters))

        '"Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d40", listOfParameters))

        '"Отверстие (строится инструментом отверстие)"
        Dim oHoles As HoleFeatures = partDoc.ComponentDefinition.Features.HoleFeatures
        Dim count As Integer = oHoles.Count 'если > 0, значит используется инструмент отверстие
        If count > 0 Then
            addInventorValuesInAspectDataList(True)
        Else
            addInventorValuesInAspectDataList(False)
        End If

        '"Отверстие Ø6"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d30", listOfParameters))

        '"Отверстие глубина 5"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d35", listOfParameters))

        '"Отверстие 4 экземпляра (круговым массивом)"
        Dim oCircularPatternFeatures As CircularPatternFeatures = partDoc.ComponentDefinition.Features.CircularPatternFeatures
        If oCircularPatternFeatures.Count = 1 Then
            'найден один круговой массив. необходимо получить количество элементов этого массива
            Dim countOfElems As Integer = oCircularPatternFeatures(1).Count.Value
            'если элементов 4 - верно, подходит условиям
            If countOfElems = 4 Then
                addInventorValuesInAspectDataList(True)
            Else
                addInventorValuesInAspectDataList(False)
            End If
        Else
            'круговых массивов нет, или больше одного (и то, и то неверно)
            addInventorValuesInAspectDataList(False)
        End If

        '«Размер линейный (на виде спереди, положение отверстий)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d29", listOfParameters))

        '"Геометрия (все цилиндрические и конические поверхности, кроме 4 отв., соосны)"
        addInventorValuesInAspectDataList("EMPTY VALUE")
    End Sub

    Private Sub getPart007(ByVal partDoc As Document)
        Dim listOfParameters As New List(Of PartParameter)()  'получить список параметров документа
        listOfParameters = getParametersFromPart(partDoc)

        getFirstSameParametersOfPart(partDoc)

        '«Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d1", listOfParameters))

        '«Размер Ø20"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d5", listOfParameters))

        '«Размер Ø33"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d6", listOfParameters))

        '«Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d2", listOfParameters))

        '«Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d3", listOfParameters))

        '«Размер линейный"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d4", listOfParameters))

        '«Размер Ø7,7"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d9", listOfParameters))

        '«Размер линейный (проточка)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d11", listOfParameters))

        '«Размер линейный (проточка)"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d10", listOfParameters))

        '«Размер угловой (проточка)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '«Размер R0,8"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d12", listOfParameters))

        '«Размер R0,8"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d12", listOfParameters))

        '«Резьба наружная"
        addInventorValuesInAspectDataList(getThreadsParams(partDoc))

        '«Размер □18"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d18", listOfParameters))

        '«Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d13", listOfParameters))

        '«Фаска"
        addInventorValuesInAspectDataList(findValueInPartParamListByName("d13", listOfParameters))

        '«Геометрия (□18 ориентирован относительно осей симметрии)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '«Геометрия (все цилиндрические и конические поверхности соосны)"
        addInventorValuesInAspectDataList("EMPTY VALUE")
    End Sub

    Private Sub getAsm(ByVal asmDoc As Document)
        Dim occ As ComponentOccurrence

        '"Деталь 05.01.003 Корпус закреплена (0 степеней свободы)"
        Dim result As String = "EMPTY VALUE"
        For Each occ In asmDoc.ComponentDefinition.Occurrences 'occ - свойства part document (1..n) В assembly, их (документов) перебор
            'если деталь "05.01.003"
            If (occ.Name = "05.01.003" & ":1") Then
                result = occ.Grounded 'true - да, деталь закреплена, false - нет, деталь не закреплена
                Exit For
            End If
        Next
        addInventorValuesInAspectDataList(result)

        '"Зависимости в паре 05.01.001 Вилка и 05.01.004 Ось (ось 004 совпадает с осью отверстия 001)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.001 Вилка и 05.01.004 Ось (ограничение перемещения 004 вдоль оси)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.001 Вилка и 05.01.005 Ролик (ось 005 совпадает с осью отверстия 001 или с осью 004)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.001 Вилка и 05.01.005 Ролик (ограничение перемещения 005 вдоль оси)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.001 Вилка и 05.01.007 Шток (ось резьбы 001 совпадает с осью резьбы 007)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.001 Вилка и 05.01.007 Шток (плоскость 001 совпадает с плоскостью 007)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.001 Вилка и 05.01.007 Шток (угол поворота 001 относительно 007 указан)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.003 Корпус и 05.01.007 Шток (ось цилиндра 003 совпадает с осью цилиндра 007, или две плоскости □18 на 003 и на 007 совпадают)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.003 Корпус и 05.01.007 Шток (ось цилиндра 003 совпадает с осью цилиндра 007 и по одной плоскости □18 совпадают или указан угол поворота, или две плоскости □18 на 003 и на 007 совпадают)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.003 Корпус и 05.01.007 Шток (007 упирается в 003 буртиком Ø33 (совпадение плоскостей))"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.002 Гайка и 05.01.003 Корпус (цилиндры 002 соосны резьбе 003)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.002 Гайка и 05.01.003 Корпус (плоскость (торец) гайки 002 находится в указанной позиции относительно резьбы на 003 (координата вычисляема))"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.002 Гайка и 05.01.003 Корпус (угловое положение 002 относительно 003)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.006 Стакан и 05.01.002 Гайка (или 003 Корпус, или 007 Шток) (соосность цилиндрических поверхностей)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.006 Стакан и 05.01.002 Гайка (совпадение плоскостей (торцев) 006 Стакан и 002 Гайка)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Зависимости в паре 05.01.006 Стакан и 05.01.002 Гайка (угловое положение 006 относительно 002)"
        addInventorValuesInAspectDataList("EMPTY VALUE")
    End Sub

    Private Sub getDrawing007(ByVal drawingDoc As Document)

        Dim oSheet As Sheet = drawingDoc.Sheets.Item(1) 'лист чертежа
        Dim oView As DrawingView = oSheet.DrawingViews.Item(1) 'вид листа     

        'start (get numeric parameters)
        'Dim str As String = ""
        'For Each drawDim In oSheet.DrawingDimensions
        '    str &= "ModelValue: " & drawDim.ModelValue.ToString & vbCrLf
        '    'MsgBox(drawDim.Text.Origin.X)
        'Next
        'MsgBox(str)
        'end

        '"Выбор ориентации детали (протяженная, большая часть поверхностей цилиндры)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Выбор главного вида (в данном случае учитывается ориентация осей)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Выбор формата листа"
        Dim result As String = "EMPTY VALUE"
        If oSheet.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize Then
            result = "А4"
        ElseIf oSheet.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize Then
            result = "А3"
        Else
            result = "Другой формат"
        End If
        addInventorValuesInAspectDataList(result)

        '"Выбор масштаба главного вида"
        result = "EMPTY VALUE"
        Dim oPropSets As PropertySets = drawingDoc.PropertySets
        Dim oPropSetGOST As PropertySet = oPropSets.Item("Свойства ГОСТ")
        result = oPropSetGOST.Item("Масштаб").Value
        addInventorValuesInAspectDataList(result)

        '"Как разместить главный вид на листе формата А4? Сделать разрыв"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"На главном виде отобразается плоскость □18. Перекрестие"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Главный вид □18. Перекрестие. Отдельный эскиз."
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Главный вид □18. Перекрестие. Отдельный эскиз. Вес линий."
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Деталь симметрична на главном виде. Наличие осевой во всю длину проекции +5 мм за пределы контура."
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер линейный (габаритный)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер Ø20"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер Ø33"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер линейный (на виде спереди)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Резьба наружная"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Фаска"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер □18 не может быть указан -> Разрез"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Разрез находится за пределами листа -> Отключение выравнивания"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Разрез помещён на свободное пространство листа"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Геометрия на виде симметрична относительно двух осей (+окружность) -> 2 осевых линии"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Разрез. Масштаб (1:1) совпадает с масштабом главного вида -> удаляем (1:1) после А-А"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер □18"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размеры проточки на главном виде не удастся показать -> Выносной вид"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"По умолчанию наименование вида B (лат.) -> меняем на Б"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"По умолчанию масштаб вида Б устанавливается равным 2:1, для размещения всех необходимых размеров изменяем его на 4:1"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Чтобы показать линии перехода в местах скруглений разрываем связь стиля вида Б с главным видом"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"В параметрах отображения вида Б включаем линии перехода"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер R0,8 (2 места)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер Ø7,7"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер линейный на виде Б (2,49 округлить до 2,5)"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Размер линейный на виде Б"
        addInventorValuesInAspectDataList("EMPTY VALUE")

        '"Заполнение основной надписи"
        Dim author As String = Nothing
        Dim designation As String = Nothing
        Dim header As String = Nothing

        Dim oTitleBlock As TitleBlock = oSheet.TitleBlock
        For Each tb As TextBox In oTitleBlock.Definition.Sketch.TextBoxes
            If tb.Text = "<АВТОР>" Then
                author = oTitleBlock.GetResultText(tb)
            End If
            If tb.Text = "<ОБОЗНАЧЕНИЕ>" Then
                designation = oTitleBlock.GetResultText(tb)
            End If
            If tb.Text = "<ЗАГОЛОВОК>" Then
                header = oTitleBlock.GetResultText(tb)
            End If
        Next
        'если одна из строк пустая - ошибка, основная надпись не заполнена
        If (String.IsNullOrEmpty(author) Or String.IsNullOrEmpty(designation) Or String.IsNullOrEmpty(header)) Then
            addInventorValuesInAspectDataList(False)
        Else
            addInventorValuesInAspectDataList(True)
        End If

    End Sub

    'ДОПОЛНИТЕЛЬНАЯ ТЕСТОВАЯ функция получения информации об эскизах
    Private Sub getInfoAboutSketches(ByVal partDoc As Document)
        Dim finalString As String = ""
        finalString &= vbCrLf & "Имя детали: " & partDoc.DisplayName & vbCrLf
        finalString &= "Всего содержит эскизов: " & partDoc.ComponentDefinition.Sketches.Count & vbCrLf

        For Each oSketch As Sketch In partDoc.ComponentDefinition.Sketches
            finalString &= vbCrLf & "-Имя эскиза: " & oSketch.Name & vbCrLf

            finalString &= "Всего DimensionConstraints в текущем эскизе: " & oSketch.DimensionConstraints.Count & vbCrLf
            For Each oDimensionConstraint As DimensionConstraint In oSketch.DimensionConstraints
                Select Case oDimensionConstraint.Type
                    Case ObjectTypeEnum.kArcLengthDimConstraintObject
                        finalString &= "ArcLengthDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kDiameterDimConstraintObject
                        finalString &= "DiameterDimConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kEllipseRadiusDimConstraintObject
                        finalString &= "EllipseRadiusDimConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kOffsetDimConstraintObject
                        finalString &= "OffsetDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kRadiusDimConstraintObject
                        finalString &= "RadiusDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kTangentDistanceDimConstraintObject
                        finalString &= "kTangentDistanceDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kThreePointAngleDimConstraintObject
                        finalString &= "ThreePointAngleDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kTwoLineAngleDimConstraintObject
                        finalString &= "TwoLineAngleDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kTwoPointDistanceDimConstraintObject
                        finalString &= "TwoPointDistanceDimConstraint" & vbCrLf
                    Case Else
                        finalString &= "Неизвестно" & vbCrLf
                End Select
            Next

            finalString &= "Всего GeometricConstraints в текущем эскизе: " & oSketch.GeometricConstraints.Count & vbCrLf
            For Each oGeometricConstraint As GeometricConstraint In oSketch.GeometricConstraints
                Select Case oGeometricConstraint.Type
                    Case ObjectTypeEnum.kCoincidentConstraintObject
                        finalString &= "CoincidentConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kCollinearConstraintObject
                        finalString &= "CollinearConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kConcentricConstraintObject
                        finalString &= "ConcentricConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kEqualLengthConstraintObject
                        finalString &= "EqualLengthConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kEqualRadiusConstraintObject
                        finalString &= "EqualRadiusConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kGroundConstraintObject
                        finalString &= "GroundConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kHorizontalAlignConstraintObject
                        finalString &= "HorizontalAlignConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kHorizontalConstraintObject
                        finalString &= "HorizontalConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kMidpointConstraintObject
                        finalString &= "MidpointConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kOffsetConstraintObject
                        finalString &= "OffsetConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kParallelConstraintObject
                        finalString &= "ParallelConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kPatternConstraintObject
                        finalString &= "PatternConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kPerpendicularConstraintObject
                        finalString &= "PerpendicularConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kSmoothConstraintObject
                        finalString &= "SmoothConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kSplineFitPointConstraintObject
                        finalString &= "SplineFitPointConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kSymmetryConstraintObject
                        finalString &= "SymmetryConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kTangentSketchConstraintObject
                        finalString &= "TangentSketchConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kVerticalAlignConstraintObject
                        finalString &= "VerticalAlignConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kVerticalConstraintObject
                        finalString &= "VerticalConstraintObject" & vbCrLf
                    Case Else
                        finalString &= "Неизвестно" & vbCrLf
                End Select
            Next
        Next

        My.Computer.FileSystem.WriteAllText("C:\Users\Сергей\Desktop\sketches_info.txt", finalString, True)
    End Sub


    'НЕАКТУЛЬНАЯ вспомогательная функция: заменить в структуре типа AspectData значение, получаемое из Inventor (2 последних столбца)
    'неактуальная потому, что имена аспектов не уникальные
    'Private Sub changeValueFromInventorInAspectDataList(ByVal text As String, ByVal value As String)
    '    For Each aspect As AspectData In _listAspects
    '        If aspect.text = text Then
    '            aspect.valueFromInventor = value

    '            'подсчет дельты
    '            Try
    '                aspect.delta = Math.Abs((aspect.valueFromExcel - aspect.valueFromInventor) / aspect.valueFromExcel)
    '            Catch ex As Exception
    '                aspect.delta = 0
    '            End Try

    '            'увеличить счетчик считанных свойств из Inventor
    '            Dim c As Integer = CInt(lblCountOfAssembly.Text)
    '            c += 1
    '            lblCountOfAssembly.Text = c

    '            Exit For
    '        End If
    '    Next
    'End Sub

End Class