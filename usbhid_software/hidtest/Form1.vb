Imports System.Runtime.InteropServices

Public Class Form1
    <DllImport("hidapi.dll", EntryPoint:="hid_enumerate", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_enumerate(ByVal a As UShort, ByVal b As UShort) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_open", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_open(ByVal a As UShort, ByVal b As UShort, ByVal serial As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_get_manufacturer_string", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_get_manufacturer_string(ByVal handle As Integer, ByVal wstr() As Int16, ByVal MAX_STR As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_get_product_string", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_get_product_string(ByVal handle As Integer, ByVal wstr() As Int16, ByVal MAX_STR As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_get_serial_number_string", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_get_serial_number_string(ByVal handle As Integer, ByVal wstr() As Int16, ByVal MAX_STR As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_get_indexed_string", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_get_indexed_string(ByVal handle As Integer, ByVal index As Integer, ByVal wstr() As Int16, ByVal MAX_STR As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_set_nonblocking", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_set_nonblocking(ByVal handle As Integer, ByVal nonblock As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_send_feature_report", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_send_feature_report(ByVal handle As Integer, ByVal data() As Byte, ByVal length As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_get_feature_report", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_get_feature_report(ByVal handle As Integer, ByVal data() As Byte, ByVal length As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_read", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_read(ByVal handle As Integer, ByVal data() As Byte, ByVal length As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_write", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_write(ByVal handle As Integer, ByVal data() As Byte, ByVal length As Integer) As Integer
    End Function
    <DllImport("hidapi.dll", EntryPoint:="hid_close", CallingConvention:=CallingConvention.Cdecl)> Private Shared Function hid_close(ByVal handle As Integer) As Integer
    End Function

    'res = hid_get_manufacturer_string(handle, wstr, MAX_STR);
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim handle As Integer
        Dim dev_ptr As Integer
        Dim MyArrayOfBytes(0 To 255) As Int16
        Dim res As Int16
        Dim manufact_str As String
        Dim product_str As String
        Dim serialnum_str As String
        Dim tmp As String

        dev_ptr = hid_enumerate(&HF1B, &H119)
        If (dev_ptr = 0) Then
            'no device is connected
            Return
        End If

        handle = hid_open(&HF1B, &H119, &H0)
        res = hid_get_manufacturer_string(handle, MyArrayOfBytes, 255)
        manufact_str = wchar_to_string(MyArrayOfBytes)

        res = hid_get_product_string(handle, MyArrayOfBytes, 255)
        product_str = wchar_to_string(MyArrayOfBytes)

        res = hid_get_serial_number_string(handle, MyArrayOfBytes, 255)
        serialnum_str = wchar_to_string(MyArrayOfBytes)

        '__ I am not using this function
        res = hid_get_indexed_string(handle, 1, MyArrayOfBytes, 255)
        tmp = wchar_to_string(MyArrayOfBytes)
        '__
        Dim buf(0 To 20) As Byte
        buf(0) = &H0
        buf(1) = &H1
        buf(2) = &H0
        buf(3) = &H0
        buf(4) = &H0
        buf(5) = &H0
        buf(6) = &H0
        buf(7) = &H0
        buf(8) = &H0

        res = hid_send_feature_report(handle, buf, 9)
        'hid_set_nonblocking(handle, 1)
        res = hid_read(handle, buf, 8)
        hid_close(handle)

    End Sub

    Private Function wchar_to_string(ByVal charArray() As Int16) As String
        Dim x As Char
        Dim i As Integer
        Dim arr(charArray.Length) As Char

        For i = 0 To charArray.Length - 1
            If (charArray(i) > 0) Then
                arr(i) = ChrW(charArray(i))
            End If
        Next
        Dim value As String = New String(arr)
        Return value
    End Function
End Class
