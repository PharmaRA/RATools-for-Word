Attribute VB_Name = "Core_Text"
Option Explicit

' =============================================
' 公共层：文本工具
' =============================================

' 由 Unicode 码点数组构造字符串。
' 用于个别必须规避源码编码差异的场合；常规中文字面量仍直接以 GBK 源码书写。
' 注意：窗体代码（如 frmAbout）因导入隔离保留私有拷贝，修改时需同步。
Public Function FromCodePoints(ByVal values As Variant) As String
    Dim i As Long

    For i = LBound(values) To UBound(values)
        FromCodePoints = FromCodePoints & ChrW$(CLng(values(i)))
    Next i
End Function