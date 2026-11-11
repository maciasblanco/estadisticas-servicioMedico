Attribute VB_Name = "modSeguridad"
Option Explicit

' -------------------------------------------------------------------
'  MÓDULO: modSeguridad
'  SHA-256, validación de login y bloqueo por intentos
' -------------------------------------------------------------------

Public Function GenerarHash(ByVal texto As String) As String
    Dim bytes() As Byte
    bytes = StrConv(SALT_FIJO & "|" & texto, vbFromUnicode)
    GenerarHash = SHA256_Hex(bytes)
End Function

Private Function SHA256_Hex(ByRef data() As Byte) As String
    Dim K(0 To 63) As Long
    Dim h(0 To 7) As Long
    Dim w(0 To 63) As Long
    Dim Kvals As Variant
    Kvals = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, _
        &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, _
        &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, _
        &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, _
        &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, _
        &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, _
        &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, _
        &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, _
        &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    
    Dim i As Long
    Dim j As Long
    Dim bloque As Long
    Dim t As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim s0 As Long
    Dim s1 As Long
    Dim ch As Long
    Dim maj As Long
    Dim temp1 As Double
    Dim temp2 As Double
    Dim origLen As Long
    Dim bitLen As Double
    Dim lenBytes(0 To 7) As Byte
    
    For i = 0 To 63
        K(i) = Kvals(i)
    Next
    
    h(0) = &H6A09E667
    h(1) = &HBB67AE85
    h(2) = &H3C6EF372
    h(3) = &HA54FF53A
    h(4) = &H510E527F
    h(5) = &H9B05688C
    h(6) = &H1F83D9AB
    h(7) = &H5BE0CD19
    
    origLen = UBound(data) + 1
    ReDim Preserve data(0 To (((origLen + 9) \ 64) + 1) * 64 - 1)
    data(origLen) = &H80
    bitLen = CDbl(origLen) * 8
    
    For j = 7 To 0 Step -1
        lenBytes(7 - j) = (bitLen / (256 ^ j)) And &HFF
    Next
    For j = 0 To 7
        data(UBound(data) - 7 + j) = lenBytes(j)
    Next
    
    For bloque = 0 To UBound(data) Step 64
        For t = 0 To 15
            w(t) = (CLng(data(bloque + t * 4)) And &HFF) * &H1000000 _
                 + (CLng(data(bloque + t * 4 + 1)) And &HFF) * &H10000 _
                 + (CLng(data(bloque + t * 4 + 2)) And &HFF) * &H100 _
                 + (CLng(data(bloque + t * 4 + 3)) And &HFF)
        Next
        For t = 16 To 63
            s0 = RShift(w(t - 15), 7) Xor RShift(w(t - 15), 18) Xor (w(t - 15) \ 3)
            s1 = RShift(w(t - 2), 17) Xor RShift(w(t - 2), 19) Xor (w(t - 2) \ 10)
            w(t) = w(t - 16) + s0 + w(t - 7) + s1
            w(t) = w(t) And &HFFFFFFFF
        Next
        
        a = h(0)
        b = h(1)
        c = h(2)
        d = h(3)
        e = h(4)
        f = h(5)
        g = h(6)
        h = h(7)
        
        For t = 0 To 63
            s1 = RShift(e, 6) Xor RShift(e, 11) Xor RShift(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            temp1 = (CDbl(h) + s1 + ch + K(t) + w(t))
            temp1 = temp1 - Int(temp1 / 4294967296#) * 4294967296#
            s0 = RShift(a, 2) Xor RShift(a, 13) Xor RShift(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            temp2 = (CDbl(s0) + maj)
            temp2 = temp2 - Int(temp2 / 4294967296#) * 4294967296#
            
            h = g
            g = f
            f = e
            e = (d + temp1)
            e = e - Int(e / 4294967296#) * 4294967296#
            d = c
            c = b
            b = a
            a = (temp1 + temp2)
            a = a - Int(a / 4294967296#) * 4294967296#
        Next
        
        h(0) = (h(0) + a) And &HFFFFFFFF
        h(1) = (h(1) + b) And &HFFFFFFFF
        h(2) = (h(2) + c) And &HFFFFFFFF
        h(3) = (h(3) + d) And &HFFFFFFFF
        h(4) = (h(4) + e) And &HFFFFFFFF
        h(5) = (h(5) + f) And &HFFFFFFFF
        h(6) = (h(6) + g) And &HFFFFFFFF
        h(7) = (h(7) + h) And &HFFFFFFFF
    Next
    
    Dim s As String
    For i = 0 To 7
        s = s & Right("00000000" & Hex(h(i)), 8)
    Next
    SHA256_Hex = LCase(s)
End Function

Private Function RShift(ByVal v As Long, ByVal n As Integer) As Long
    Dim d As Double
    d = v
    If d < 0 Then d = d + 4294967296#
    RShift = Int(d / (2 ^ n)) And &HFFFFFFFF
    If RShift < 0 Then RShift = RShift + 4294967296#
End Function

Public Function ValidarLogin(ByVal usuario As String, ByVal clave As String, _
                              ByRef mensaje As String) As Boolean
    ValidarLogin = False
    
    If Not ExisteHoja(HOJA_USUARIOS) Then
        mensaje = "No existe la hoja de usuarios."
        Exit Function
    End If
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim usuarioLimpio As String
    Dim intentos As Integer
    Dim hashGuardado As String
    
    Set ws = ThisWorkbook.Sheets(HOJA_USUARIOS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    usuarioLimpio = UCase(Trim(usuario))
    
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, UCOL_USUARIO).Value))) = usuarioLimpio Then
            If UCase(Trim(CStr(ws.Cells(i, UCOL_ACTIVO).Value))) <> "SI" Then
                mensaje = "Usuario deshabilitado. Contacte al administrador."
                Exit Function
            End If
            
            intentos = Val(CStr(ws.Cells(i, UCOL_INTENTOS).Value))
            If intentos >= MAX_INTENTOS Then
                mensaje = "Usuario bloqueado por " & MAX_INTENTOS & " intentos fallidos."
                Exit Function
            End If
            
            hashGuardado = CStr(ws.Cells(i, UCOL_HASH).Value)
            
            If hashGuardado = "INICIAL" Then
                If clave = "Admin2026!" Then
                    ws.Cells(i, UCOL_HASH).Value = GenerarHash(clave)
                    ws.Cells(i, UCOL_INTENTOS).Value = 0
                    ws.Cells(i, UCOL_ULTIMO_ACCESO).Value = Format(Now, FORMATO_DATETIME)
                    InicializarSesion usuarioLimpio, CStr(ws.Cells(i, UCOL_NOMBRE).Value), CStr(ws.Cells(i, UCOL_ROL).Value)
                    mensaje = "OK"
                    ValidarLogin = True
                    Exit Function
                End If
            ElseIf LCase(GenerarHash(clave)) = LCase(hashGuardado) Then
                ws.Cells(i, UCOL_INTENTOS).Value = 0
                ws.Cells(i, UCOL_ULTIMO_ACCESO).Value = Format(Now, FORMATO_DATETIME)
                InicializarSesion usuarioLimpio, CStr(ws.Cells(i, UCOL_NOMBRE).Value), CStr(ws.Cells(i, UCOL_ROL).Value)
                mensaje = "OK"
                ValidarLogin = True
                Exit Function
            End If
            
            ws.Cells(i, UCOL_INTENTOS).Value = intentos + 1
            mensaje = "Clave incorrecta. Intentos: " & (intentos + 1) & " de " & MAX_INTENTOS
            Exit Function
        End If
    Next i
    mensaje = "Usuario no encontrado."
End Function

Public Function CambiarClave(ByVal usuario As String, ByVal claveNueva As String) As Boolean
    If Len(claveNueva) < MIN_LONG_CLAVE Then Exit Function
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets(HOJA_USUARIOS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, UCOL_USUARIO).Value))) = UCase(Trim(usuario)) Then
            ws.Cells(i, UCOL_HASH).Value = GenerarHash(claveNueva)
            ws.Cells(i, UCOL_INTENTOS).Value = 0
            CambiarClave = True
            Exit Function
        End If
    Next
End Function
