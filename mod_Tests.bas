Attribute VB_Name = "mod_Tests"
Option Explicit

' =========================================================
' ТЕСТОВАЯ МОДУЛЬ ДЛЯ IsMatch
' =========================================================

Sub RunAllTests()
    Debug.Print "--- Запуск тестов IsMatch ---"

    Test_ExactMatch
    Test_SubstringMatch_Rejected
    Test_RelaxedMatch
    Test_BoundaryConditions
    Test_MultipleOccurrences
    Test_CaseInsensitivity

    Debug.Print "--- Тесты завершены ---"
End Sub

Private Sub Assert(condition As Boolean, testName As String)
    If condition Then
        Debug.Print "[PASS] " & testName
    Else
        Debug.Print "[FAIL] " & testName
        ' Stop ' Раскомментировать для отладки в VBA
    End If
End Sub

Sub Test_ExactMatch()
    Dim ord As New clsOrder
    ord.SearchTerm = "12345"
    Assert ord.IsMatch("12345"), "Exact match"
    Assert ord.IsMatch("Order 12345 Section"), "Exact match in string"
End Sub

Sub Test_SubstringMatch_Rejected()
    Dim ord As New clsOrder
    ord.SearchTerm = "123"
    Assert Not ord.IsMatch("12345"), "Substring match (suffix) rejected"
    Assert Not ord.IsMatch("0123"), "Substring match (prefix) rejected"
    Assert Not ord.IsMatch("123-A"), "Substring match (hyphen suffix) rejected"
End Sub

Sub Test_RelaxedMatch()
    Dim ord As New clsOrder
    ord.SearchTerm = "-123" ' Расслабленный режим
    Assert ord.IsMatch("12345"), "Relaxed match (prefix) accepted"
    Assert ord.IsMatch("0123"), "Relaxed match (suffix) accepted"
    Assert ord.IsMatch("ABC123XYZ"), "Relaxed match (middle) accepted"
End Sub

Sub Test_BoundaryConditions()
    Dim ord As New clsOrder
    ord.SearchTerm = "ABC"
    Assert ord.IsMatch("ABC"), "Start of string"
    Assert ord.IsMatch("X ABC"), "Middle of string"
    Assert ord.IsMatch("X ABC"), "End of string"

    Assert Not ord.IsMatch("X-ABC"), "Hyphen prefix rejected in strict mode"
End Sub

Sub Test_MultipleOccurrences()
    Dim ord As New clsOrder
    ord.SearchTerm = "123"
    ' Первая "123" - часть "1234", должна быть проигнорирована.
    ' Вторая "123" - отдельная, должна дать совпадение.
    Assert ord.IsMatch("1234 123"), "Multiple occurrences - valid second"

    ' Наоборот: первая валидная, вторая нет.
    Assert ord.IsMatch("123 1234"), "Multiple occurrences - valid first"
End Sub

Sub Test_CaseInsensitivity()
    Dim ord As New clsOrder
    ord.SearchTerm = "abc"
    Assert ord.IsMatch("ABC"), "Case insensitive match"

    ord.SearchTerm = "ABC"
    Assert ord.IsMatch("abc"), "Case insensitive match (reverse)"
End Sub
