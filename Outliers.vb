Imports System.Linq
Imports System.Collections.Generic

Namespace Stat

    ''' <summary>
    ''' Отсев грубых погрешностей в малых (3..25 элементов), средних (25..500) и больших (более 500) выборках.
    ''' </summary>
    Public Module Outliers

#Region "КОНСТАНТЫ, ТАБЛИЧНЫЕ ЗНАЧЕНИЯ"

        Public Const MIN_DATALEN As Integer = 3
        Public Const SMALL_DATALEN As Integer = 25
        Public Const BIG_DATALEN As Integer = 500

        ''' <summary>
        ''' Массив коэффициентов критических значений статистики при разных значениях числа измерений для надёжности 0,90.
        ''' </summary>
        ''' <value>Индекс числа в массиве соответствует количеству измерений. Например, для 4-х измерений значение статистики 1,64.</value>
        ''' <remarks>Приложение IX из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Private ReadOnly Property TauCritical90 As Double() = {
            0, 0, 0, 1.41, 1.64, 1.79, 1.89, 1.97, 2.04, 2.1, 2.15, 2.19, 2.23, 2.26, 2.3, 2.33, 2.35, 2.38, 2.4,
            2.43, 2.45, 2.47, 2.49, 2.5, 2.52, 2.54, 2.55, 2.57, 2.58, 2.6, 2.61, 2.62, 2.63, 2.65, 2.66, 2.67,
            2.68, 2.69, 2.7, 2.71, 2.72, 2.73, 2.74, 2.74, 2.75, 2.76, 2.77, 2.78, 2.78, 2.79, 2.8, 2.81, 2.81
        }

        ''' <summary>
        ''' Массив коэффициентов критических значений статистики при разных значениях числа измерений для надёжности 0,95.
        ''' </summary>
        ''' <value>Индекс числа в массиве соответствует количеству измерений. Например, для 4-х измерений значение статистики 1,69.</value>
        ''' <remarks>Приложение IX из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Private ReadOnly Property TauCritical95 As Double() = {
            0, 0, 0, 1.41, 1.69, 1.87, 2.0, 2.09, 2.17, 2.24, 2.29, 2.34, 2.39, 2.43, 2.46, 2.49, 2.52, 2.55, 2.58,
            2.6, 2.62, 2.64, 2.66, 2.68, 2.7, 2.72, 2.73, 2.75, 2.76, 2.78, 2.79, 2.8, 2.82, 2.83, 2.84, 2.85,
            2.86, 2.87, 2.88, 2.89, 2.9, 2.91, 2.92, 2.93, 2.94, 2.95, 2.96, 2.96, 2.97, 2.98, 2.99, 2.99, 3.0
        }

        ''' <summary>
        ''' Массив коэффициентов критических значений статистики при разных значениях числа измерений для надёжности 0,99. 
        ''' </summary>
        ''' <value>Индекс числа в массиве соответствует количеству измерений. Например, для 4-х измерений, значение статистики 1,72.</value>
        ''' <remarks>Приложение IX из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Private ReadOnly Property TauCritical99 As Double() = {
            0, 0, 0, 1.41, 1.72, 1.96, 2.13, 2.26, 2.37, 2.46, 2.54, 2.61, 2.66, 2.71, 2.76, 2.8, 2.84, 2.87, 2.9,
            2.93, 2.96, 2.98, 3.01, 3.03, 3.05, 3.07, 3.09, 3.11, 3.12, 3.14, 3.16, 3.17, 3.18, 3.2, 3.21, 3.22,
            3.24, 3.25, 3.26, 3.27, 3.28, 3.29, 3.3, 3.31, 3.32, 3.33, 3.34, 3.35, 3.35, 3.36, 3.37, 338, 3.39
        }

        ''' <summary>
        ''' Количества измерений коэффициентов Стьюдента - столбец N-2 в таблице.
        ''' </summary>
        Private ReadOnly Property StudentN As Integer() = {
            1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30,
            32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 55, 60, 65, 70, 80, 90, 100, 120, 150, 200, 250, 300, 400, 500
        }

        ''' <summary>
        ''' Словарь количества измерений N-1 и соответствующих коэффициентов Стьюдента при разных значениях числа измерений для надёжности 0,90.
        ''' </summary>
        ''' <value>Количество измерений - это значение ключа в словаре, а значение коэффициента - значение в словаре.</value>
        ''' <remarks>Приложение VII из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Public ReadOnly Property Student90 As Dictionary(Of Integer, Double)
            Get
                If (_Student90.Count = 0) Then
                    Dim vals As Double() = {
                        6.31, 2.92, 2.35, 2.13, 2.02, 1.94, 1.89, 1.86, 1.83, 1.81, 1.8, 1.78, 1.77, 1.76, 1.75, 1.75, 1.74, 1.73,
                        1.73, 1.72, 1.72, 1.72, 1.71, 1.71, 1.71, 1.71, 1.7, 1.7, 1.7, 1.7, 1.69, 1.69, 1.69, 1.69, 1.68, 1.68,
                        1.68, 1.68, 1.68, 1.68, 1.67, 1.67, 1.67, 1.67, 1.66, 1.66, 1.66, 1.66, 1.66, 1.65, 1.65, 1.65, 1.65, 1.65
                    }
                    For i As Integer = 0 To vals.Length - 1
                        _Student90.Add(StudentN(i), vals(i))
                    Next
                End If
                Return _Student90
            End Get
        End Property
        Private _Student90 As New Dictionary(Of Integer, Double)

        ''' <summary>
        ''' Словарь количества измерений N-1 и соответствующих коэффициентов Стьюдента при разных значениях числа измерений для надёжности 0,95.
        ''' </summary>
        ''' <value>Количество измерений - это значение ключа в словаре, а значение коэффициента - значение в словаре.</value>
        ''' <remarks>Приложение VII из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Public ReadOnly Property Student95 As Dictionary(Of Integer, Double)
            Get
                If (_Student95.Count = 0) Then
                    Dim vals As Double() = {
                        12.7, 4.3, 3.18, 2.78, 2.57, 2.45, 2.36, 2.31, 2.26, 2.23, 2.2, 2.18, 2.16, 2.14, 2.13, 2.12, 2.11, 2.1,
                        2.09, 2.09, 2.08, 2.07, 2.07, 2.06, 2.06, 2.06, 2.05, 2.05, 2.05, 2.04, 2.04, 2.03, 2.03, 2.02, 2.02,
                        2.02, 2.02, 2.01, 2.01, 2.01, 2, 2, 2, 1.99, 1.99, 1.99, 1.98, 1.98, 1.98, 1.97, 1.97, 1.97, 1.97, 1.96
                    }
                    For i As Integer = 0 To vals.Length - 1
                        _Student95.Add(StudentN(i), vals(i))
                    Next
                End If
                Return _Student95
            End Get
        End Property
        Private _Student95 As New Dictionary(Of Integer, Double)

        ''' <summary>
        ''' Массив коэффициентов Стьюдента при разных значениях числа измерений для надёжности 0,98.
        ''' </summary>
        ''' <value>Количество измерений - это значение ключа в словаре, а значение коэффициента - значение в словаре.</value>
        ''' <remarks>Приложение VII из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Public ReadOnly Property Student98 As Dictionary(Of Integer, Double)
            Get
                If (_Student98.Count = 0) Then
                    Dim vals As Double() = {
                        31.8, 6.96, 4.54, 3.75, 3.36, 3.14, 3, 2.9, 2.82, 2.76, 2.72, 2.68, 2.65, 2.62, 2.6, 2.58, 2.57, 2.55,
                        2.54, 2.53, 2.52, 2.51, 2.5, 2.49, 2.49, 2.48, 2.47, 2.47, 2.46, 2.46, 2.45, 2.44, 2.43, 2.43, 2.42, 2.42,
                        2.41, 2.41, 2.41, 2.4, 2.4, 2.39, 2.39, 2.38, 2.37, 2.37, 2.36, 2.36, 2.35, 2.35, 2.34, 2.34, 2.34, 2.33
                    }
                    For i As Integer = 0 To vals.Length - 1
                        _Student98.Add(StudentN(i), vals(i))
                    Next
                End If
                Return _Student98
            End Get
        End Property
        Private _Student98 As New Dictionary(Of Integer, Double)

        ''' <summary>
        ''' Массив коэффициентов Стьюдента при разных значениях числа измерений для надёжности 0,99.
        ''' </summary>
        ''' <value>Количество измерений - это значение ключа в словаре, а значение коэффициента - значение в словаре.</value>
        ''' <remarks>Приложение VII из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Public ReadOnly Property Student99 As Dictionary(Of Integer, Double)
            Get
                If (_Student99.Count = 0) Then
                    Dim vals As Double() = {
                        63.7, 9.92, 5.84, 4.6, 4.03, 3.71, 3.5, 3.36, 3.25, 3.17, 3.11, 3.05, 3.01, 2.98, 2.95, 2.92, 2.9, 2.88,
                        2.86, 2.85, 2.83, 2.82, 2.81, 2.8, 2.79, 2.78, 2.77, 2.76, 2.76, 2.75, 2.74, 2.73, 2.72, 2.71, 2.7, 2.7,
                        2.69, 2.69, 2.68, 2.68, 2.67, 2.66, 2.65, 2.65, 2.64, 2.63, 2.63, 2.62, 2.61, 2.6, 2.6, 2.59, 2.59, 2.59
                    }
                    For i As Integer = 0 To vals.Length - 1
                        _Student99.Add(StudentN(i), vals(i))
                    Next
                End If
                Return _Student99
            End Get
        End Property
        Private _Student99 As New Dictionary(Of Integer, Double)

        ''' <summary>
        ''' Массив коэффициентов Стьюдента при разных значениях числа измерений для надёжности 0,999.
        ''' </summary>
        ''' <value>Индекс числа в массиве соответствует количеству измерений. Например, для 5-ти измерений, значение статистики 6,87.</value>
        ''' <remarks>Приложение VII из книги О.Н.Кассандровой "Обработка результатов наблюдений".</remarks>
        Public ReadOnly Property Student999 As Dictionary(Of Integer, Double)
            Get
                If (_Student999.Count = 0) Then
                    Dim vals As Double() = {
                        636.6, 31.6, 12.9, 8.61, 6.87, 5.96, 5.41, 5.04, 4.78, 4.59, 4.44, 4.32, 4.22, 4.14, 4.07, 4.02, 3.97, 3.92,
                        3.88, 3.85, 3.82, 3.79, 3.77, 3.75, 3.73, 3.71, 3.69, 3.67, 3.66, 3.65, 3.62, 3.6, 3.58, 3.57, 3.55, 3.54,
                        3.53, 3.52, 3.51, 3.5, 3.48, 3.49, 3.45, 3.44, 3.42, 3.4, 3.39, 3.37, 3.36, 3.34, 3.33, 3.32, 3.32, 3.31
                    }
                    For i As Integer = 0 To vals.Length - 1
                        _Student999.Add(StudentN(i), vals(i))
                    Next
                End If
                Return _Student999
            End Get
        End Property
        Private _Student999 As New Dictionary(Of Integer, Double)

        ''' <summary>
        ''' Надёжность.
        ''' </summary>
        Public Enum Reliability
            _09
            _095
            _098
            _099
            _0999
        End Enum

#End Region '/КОНСТАНТЫ, ТАБЛИЧНЫЕ ЗНАЧЕНИЯ

#Region "МЕТОДЫ"

        ''' <summary>
        ''' Исключает из выборки (от 3-х элементов и выше) данные с резко отличающимися значениями (грубые погрешности). 
        ''' </summary>
        ''' <param name="data">Массив данных, которые нужно проверить на предмет резко отличающихся значений. Минимальный набор - <see cref="MIN_DATALEN"/> элемента.
        ''' Считается, что значения элементов подчинены нормальному закону распределения.</param>
        ''' <param name="reliability">Надёжность.</param>
        ''' <remarks>В зависимости от длины выборки используются разные алгоритмы отсеивания грубых погрешностей:
        ''' (10..25] - метод максимального относительного отклонения;
        ''' [3...10] - метод максимального относительного отклонения с уточняющим к-том;
        ''' (25...x) - распределение Стьюдента.
        ''' </remarks>
        Public Function RemoveOutliers(data As IEnumerable(Of Double), Optional reliability As Reliability = Reliability._09) As IEnumerable(Of Double)
            If (data.Count < MIN_DATALEN) Then
                Throw New ArgumentException($"Выборка должна содержать хотя бы {MIN_DATALEN} элемента.")

            ElseIf (data.Count <= SMALL_DATALEN) Then 'малая выборка
                Return GetShortDataWoOutliers(data, reliability)

            Else 'большая выборка
                If (data.Count <= BIG_DATALEN) Then
                    Return GetLongDataWoOutliers(data, reliability)
                Else
                    Dim r As New List(Of Double)
                    Dim stp As Integer = CInt(data.Count / System.Math.Ceiling(data.Count / BIG_DATALEN))
                    For i As Integer = 0 To data.Count - 1 Step stp
                        Dim d As IEnumerable(Of Double) = From dt As Double In data Skip (i) Take stp Select dt
                        Dim r1 As IEnumerable(Of Double) = GetLongDataWoOutliers(d, reliability)
                        r.AddRange(r1)
                    Next
                    Return r
                End If
            End If
        End Function

#Region "ЗАКРЫТЫЕ"

        ''' <summary>
        ''' Исключает из малой выборки (от 3-х до 25-ти элементов) грубые погрешности.
        ''' </summary>
        Private Function GetShortDataWoOutliers(data As IEnumerable(Of Double), reliability As Reliability) As IEnumerable(Of Double)
            If (data.Count >= MIN_DATALEN) AndAlso (data.Count <= SMALL_DATALEN) Then
                Dim resData As New List(Of Double)(data)
                Do
                    If (resData.Count >= MIN_DATALEN) Then
                        'Определяем оценки среднего значения и среднеквадратичного отклонения, минимальное и максимальное значения:
                        Dim meanX As Double = resData.Average()
                        Dim minX As Double = resData.Min()
                        Dim maxX As Double = resData.Max()
                        Dim xStdDeviation As Double = CalculateStdDeviation(resData)

                        'Для найденных экстремальных значений определяем тау-статистики:
                        Dim tau1 As Double = (meanX - minX) / xStdDeviation
                        Dim tauN As Double = (maxX - meanX) / xStdDeviation

                        'Для выборок менее 10-ти элементов вводим уточняющий множитель:
                        If (data.Count <= 10) Then
                            Dim rectificationCoef As Double = 1 / System.Math.Sqrt((resData.Count - 1) / resData.Count)
                            tau1 *= rectificationCoef
                            tauN *= rectificationCoef
                        End If

                        'Определяем критическое значение тау-статистики:
                        Dim tauCritical As Double = GetTauCriticalTabulated(resData.Count, reliability)

                        'Проверяем вектор на наличие резко отличающихся значений:
                        If (tau1 > tauCritical) OrElse (tauN > tauCritical) Then
                            'Удаляем экстремальный элемент:
                            If (tau1 >= tauN) Then
                                Dim minXIndex As Integer = resData.IndexOf(minX)
                                resData.RemoveAt(minXIndex)
                            Else
                                Dim maxXIndex As Integer = resData.IndexOf(maxX)
                                resData.RemoveAt(maxXIndex)
                            End If
                        Else
                            Return resData
                        End If
                    Else
                        Return resData
                    End If
                Loop
            Else
                Throw New ArgumentException($"Выборка должна содержать от {MIN_DATALEN} до {SMALL_DATALEN} элементов.")
            End If
        End Function

        ''' <summary>
        ''' Исключает из большой выборки (от <see cref="SMALL_DATALEN"/> до <see cref="BIG_DATALEN"/> элементов) грубые погрешности.
        ''' </summary>
        Private Function GetLongDataWoOutliers(data As IEnumerable(Of Double), reliability As Reliability) As IEnumerable(Of Double)
            If (data.Count >= SMALL_DATALEN) AndAlso (data.Count <= BIG_DATALEN) Then
                Dim resList As New List(Of Double)(data)
                Do
                    '1) Определение наибольшего отклонения:
                    Dim xMean As Double = resList.Average()
                    Dim dXs As List(Of Double) = (From x As Double In resList Select System.Math.Abs(x - xMean)).ToList()
                    Dim dxMax As Double = dXs.Max()
                    Dim maxDxIndex As Integer = dXs.IndexOf(dxMax)
                    Dim xi As Double = resList(maxDxIndex)

                    '2) Определение тау-статистики:
                    Dim xStdDeviation As Double = CalculateStdDeviation(resList)
                    Dim tau As Double = System.Math.Abs(xi - xMean) / xStdDeviation

                    '3) По таблице Стьюдента ищём процентные точки распределения:
                    Dim cnt As Integer = resList.Count()
                    Dim t001 As Double = GetStudentCoefTabulated(cnt - 2, Reliability._0999)
                    Dim t5 As Double = GetStudentCoefTabulated(cnt - 2, Reliability._095)

                    '4) Вычисляем критические значения распределения Стьюдента:
                    Dim tau001 As Double = t001 * System.Math.Sqrt(cnt - 1) / System.Math.Sqrt(cnt - 2 + t001 * t001)
                    Dim tau5 As Double = t5 * System.Math.Sqrt(cnt - 1) / System.Math.Sqrt(cnt - 2 + t5 * t5)

                    '5) Отсеиваем грубую погрешность:
                    If (tau >= tau001) Then
                        resList.RemoveAt(maxDxIndex)
                    Else
                        Return resList
                    End If
                Loop
            Else
                Throw New ArgumentException($"Выборка должна содержать от {SMALL_DATALEN} до {BIG_DATALEN} элементов.")
            End If
        End Function

        ''' <summary>
        ''' Вычисляет среднеквадратическое отклонение.
        ''' </summary>
        ''' <param name="data">Массив данных, по которым вычисляется СКО.</param>
        ''' <param name="isShifted">Нужна ли смещённая оценка СКО. По умолчанию вычисляется несмещённая.</param>
        Private Function CalculateStdDeviation(data As IEnumerable(Of Double), Optional isShifted As Boolean = False) As Double

            Dim meanX As Double = data.Average()

            Dim squaresSum As Double = 0
            For Each currentX As Double In data
                Dim s As Double = currentX - meanX
                squaresSum += (s * s)
            Next

            Dim n As Integer = data.Count - 1
            If isShifted Then
                n += 1
            End If

            Dim res As Double = System.Math.Sqrt(squaresSum / n)
            Return res

        End Function

        ''' <summary>
        ''' Возвращает табличное критическое значение статистики.
        ''' </summary>
        ''' <param name="dataLength">Длина выборки, для которой проводится отсеивание грубых погрешностей. Допустимые значения - от <see cref="MIN_DATALEN"/> до 52.</param>
        ''' <param name="reliability">Надёжность. Допустимые значения 0.9, 0.95 или 0.99. По умолчанию - 0.9.</param>
        Private Function GetTauCriticalTabulated(dataLength As Integer, Optional reliability As Reliability = Reliability._09) As Double
            If (dataLength >= MIN_DATALEN) AndAlso (dataLength <= TauCritical90.Length) Then
                Select Case reliability
                    Case Reliability._099
                        Return TauCritical99(dataLength)

                    Case Reliability._095
                        Return TauCritical95(dataLength)

                    Case Else
                        Return TauCritical90(dataLength)
                End Select
            Else
                Throw New ArgumentException($"Длина выборки должна лежать в диапазоне от {MIN_DATALEN} до {TauCritical90.Length} элементов.")
            End If
        End Function

        ''' <summary>
        ''' Возвращает расчётное критическое значение тау-статистики.
        ''' </summary>
        ''' <param name="dataLength">Длина выборки, для которой проводится отсеивание грубых погрешностей. Допустимые значения - от <see cref="MIN_DATALEN"/> до 52.</param>
        <Obsolete("Желательно использовать табличную функцию GetTauCriticalTabulated() вместо аналитической.")>
        Private Function GetTauCriticalAnalytic(dataLength As Integer) As Double
            If (dataLength >= MIN_DATALEN) AndAlso (dataLength <= TauCritical90.Length) Then
                If (dataLength < 40) Then
                    Return (2.4 + dataLength / 57 - 4.0 / dataLength)
                Else
                    Return 3
                End If
            Else
                Throw New ArgumentException($"Длина выборки должна лежать в диапазоне от {MIN_DATALEN} до {TauCritical90.Length} элементов.")
            End If
        End Function

        ''' <summary>
        ''' Возвращает табличное значение коэффициента Стьюдента для заданного количества измерений.
        ''' </summary>
        ''' <param name="dataLength">Длина выборки N-2, для которой проводится отсеивание грубых погрешностей. Допустимые значения - от 2 до <see cref="BIG_DATALEN"/>.</param>
        ''' <param name="reliability">Надёжность. Допустимые значения 0.9, 0.95, 0,98, 0,99 или 0.999. По умолчанию - 0.9.</param>
        Private Function GetStudentCoefTabulated(dataLength As Integer, Optional reliability As Reliability = Reliability._09) As Double
            If (dataLength >= 2) AndAlso (dataLength <= BIG_DATALEN) Then

                Dim coefStud As Double = 0.0
                Dim kvpMax As KeyValuePair(Of Integer, Double)
                Dim kvpMin As KeyValuePair(Of Integer, Double)

                Select Case reliability
                    Case Reliability._0999 'значение есть в словаре
                        If Student999.ContainsKey(dataLength) Then
                            Student999.TryGetValue(dataLength, coefStud)
                            Return coefStud
                        Else 'значения нет в словаре, интерполируем
                            'Ищем две пары точек, между которыми лежит интересующее значение:
                            kvpMax = (From d As KeyValuePair(Of Integer, Double) In Student999 Where d.Key >= dataLength Select d).First()
                            kvpMin = (From d As KeyValuePair(Of Integer, Double) In Student999 Where d.Key <= dataLength Select d).Last()
                        End If

                    Case Reliability._099
                        If Student99.ContainsKey(dataLength) Then
                            Student99.TryGetValue(dataLength, coefStud)
                            Return coefStud
                        Else
                            kvpMax = (From d As KeyValuePair(Of Integer, Double) In Student99 Where d.Key >= dataLength Select d).First()
                            kvpMin = (From d As KeyValuePair(Of Integer, Double) In Student99 Where d.Key <= dataLength Select d).Last()
                        End If

                    Case Reliability._098
                        If Student98.ContainsKey(dataLength) Then
                            Student98.TryGetValue(dataLength, coefStud)
                            Return coefStud
                        Else
                            kvpMax = (From d As KeyValuePair(Of Integer, Double) In Student98 Where d.Key >= dataLength Select d).First()
                            kvpMin = (From d As KeyValuePair(Of Integer, Double) In Student98 Where d.Key <= dataLength Select d).Last()
                        End If

                    Case Reliability._095
                        If Student95.ContainsKey(dataLength) Then
                            Student95.TryGetValue(dataLength, coefStud)
                            Return coefStud
                        Else
                            kvpMax = (From d As KeyValuePair(Of Integer, Double) In Student95 Where d.Key >= dataLength Select d).First()
                            kvpMin = (From d As KeyValuePair(Of Integer, Double) In Student95 Where d.Key <= dataLength Select d).Last()
                        End If

                    Case Else
                        If Student90.ContainsKey(dataLength) Then
                            Student90.TryGetValue(dataLength, coefStud)
                            Return coefStud
                        Else
                            kvpMax = (From d As KeyValuePair(Of Integer, Double) In Student90 Where d.Key >= dataLength Select d).First()
                            kvpMin = (From d As KeyValuePair(Of Integer, Double) In Student90 Where d.Key <= dataLength Select d).Last()
                        End If

                End Select

                Dim x1 As Integer = kvpMin.Key
                Dim x2 As Integer = kvpMax.Key
                Dim y1 As Double = kvpMin.Value
                Dim y2 As Double = kvpMax.Value

                'По ур-нию прямой, проходящей через две точки, ищем к-т Стьюдента: 
                coefStud = (y2 - y1) * (dataLength - x1) / (x2 - x1) + y1
                Return coefStud

            Else
                Throw New ArgumentException($"В таблице имеются значения коэффициента Стьюдента для выборок от {MIN_DATALEN - 1} до {BIG_DATALEN} элементов; другое число элементов не поддерживается.")
            End If
        End Function

#End Region '/ЗАКРЫТЫЕ

#End Region '/МЕТОДЫ 

    End Module '/Outliers

End Namespace