'/// @file GetRangeCoordinates.bas
'/// @author Austin Vandegriffe
'/// @date 2020-05-20
'/// @brief A cool VBA function that returns upper and lower range coordinates.
'/// @pre N/A
'/// @style K&R, and "one true brace style" (OTBS), and '_' variable naming
'/////////////////////////////////////////////////////////////////////
'/// @references
'/// ## N/A

Function GetRangeCoordinates(rng As String) As Variant

    Dim t_rng As Range
    Set t_rng = Range(rng)
    Dim ret() As Variant
    
    ReDim Preserve ret(0) As Variant
    
    ret(0) = Array(t_rng.Row, t_rng.Column)
    
    If t_rng.Rows.Count + t_rng.Columns.Count > 2 Then
        ReDim Preserve ret(1) As Variant
        ret(1) = Array(ret(0)(0) + t_rng.Rows.Count - 1, ret(0)(1) + t_rng.Columns.Count - 1)
    End If
    
    GetRangeCoordinates = ret
End Function