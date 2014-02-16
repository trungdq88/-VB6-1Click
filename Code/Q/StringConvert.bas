Attribute VB_Name = "StringConvert"
Option Explicit

'Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long

'ham dung de huy control duoc tao ra neu chua co ban quyen
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const UNI_REGISTERED = "OCX REGISTERED !!!!"
Private Const sANSI = "a1,a2,a3,a4,a5,a8,a81,a82,a83,a84,a85,a6,a61,a62,a63,a64,a65,e1,e2,e3,e4,e5,e6,e61,e62,e63,e64,e65,i1,i2,i3,i4,i5,o1,o2,o3,o4,o5,o6,o61,o62,o63,o64,o65,o7,o71,o72,o73,o74,o75,u1,u2,u3,u4,u5,u7,u71,u72,u73,u74,u75,y1,y2,y3,y4,y5,d9,A1,A2,A3,A4,A5,A8,A81,A82,A83,A84,A85,A6,A61,A62,A63,A64,A65,E1,E2,E3,E4,E5,E6,E61,E62,E63,E64,E65,I1,I2,I3,I4,I5,O1,O2,O3,O4,O5,O6,O61,O62,O63,O64,O65,O7,O71,O72,O73,O74,O75,U1,U2,U3,U4,U5,U7,U71,U72,U73,U74,U75,Y1,Y2,Y3,Y4,Y5,D9"
Private Const sUNICODE = "00E1,00E0,1EA3,00E3,1EA1,0103,1EAF,1EB1,1EB3,1EB5,1EB7,00E2,1EA5,1EA7,1EA9,1EAB,1EAD,00E9,00E8,1EBB,1EBD,1EB9,00EA,1EBF,1EC1,1EC3,1EC5,1EC7,00ED,00EC,1EC9,0129,1ECB,00F3,00F2,1ECF,00F5,1ECD,00F4,1ED1,1ED3,1ED5,1ED7,1ED9,01A1,1EDB,1EDD,1EDF,1EE1,1EE3,00FA,00F9,1EE7,0169,1EE5,01B0,1EE9,1EEB,1EED,1EEF,1EF1,00FD,1EF3,1EF7,1EF9,1EF5,0111,00C1,00C0,1EA2,00C3,1EA0,0102,1EAE,1EB0,1EB2,1EB4,1EB6,00C2,1EA4,1EA6,1EA8,1EAA,1EAC,00C9,00C8,1EBA,1EBC,1EB8,00CA,1EBE,1EC0,1EC2,1EC4,1EC6,00CD,00CC,1EC8,0128,1ECA,00D3,00D2,1ECE,00D5,1ECC,00D4,1ED0,1ED2,1ED4,1ED6,1ED8,01A0,1EDA,1EDC,1EDE,1EE0,1EE2,00DA,00D9,1EE6,0168,1EE4,01AF,1EE8,1EEA,1EEC,1EEE,1EF0,00DD,1EF2,1EF6,1EF8,1EF4,0110"

Private ArrFromCode() As String
Private ArrToCode() As String

Public bEnable As Boolean   'luu thuoc tinh enable cua UniXPFrame
Public bListViewUnicode As Boolean 'cho biet co su dung font chu unicode trong dieu khien khong

'Public msgb As New clsUniMsgbox

Public Function zToUnicode(ByRef sString As String) As String
    Dim sTam As String
    Dim i As Long, ArrChuoiXuLy() As String
    Dim k As Long, j As Long
    Dim sKyTu1 As String, sKyTu2 As String, sKyTu3 As String, sKyTu4 As String
    Dim sChuoiBenPhai As String, sKhoangTrang As String
    Dim iVitri As Integer
        
        If Trim$(sString) = "" Then zToUnicode = sString:    Exit Function
        
        ArrFromCode = Split(sANSI, ",")
        ArrToCode = Split(sUNICODE, ",")
        ArrChuoiXuLy = Split(sString, " ")
        For i = 0 To UBound(ArrChuoiXuLy)
            j = HaveNumber(ArrChuoiXuLy(i))
            If (j > 1) And (Not IsNumeric(ArrChuoiXuLy(i))) Then
                If j > 2 Then sTam = sTam & Left$(ArrChuoiXuLy(i), j - 2)
                sKyTu1 = Mid$(ArrChuoiXuLy(i), j - 1, 1)
                sKyTu2 = Mid$(ArrChuoiXuLy(i), j, 1)
                For k = j To Len(ArrChuoiXuLy(i))
                    If IsNumeric(Mid$(ArrChuoiXuLy(i), k + 1, 1)) And sChuoiBenPhai = "" Then
                        sKyTu3 = sKyTu3 & Mid$(ArrChuoiXuLy(i), k + 1, 1)
                    Else
                        sChuoiBenPhai = sChuoiBenPhai & Mid$(ArrChuoiXuLy(i), k + 1, 1)
                    End If
                Next
                If Trim$(sChuoiBenPhai) <> "" Then If HaveNumber(sChuoiBenPhai) > 0 Then sChuoiBenPhai = Trim$(zToUnicode(sChuoiBenPhai))
                sKyTu4 = sKyTu1 & sKyTu2 & sKyTu3
                sTam = sTam & ChangeString(sKyTu4)
            Else
                sTam = sTam & ArrChuoiXuLy(i) & " "
                GoTo TT
            End If
            sTam = sTam & sChuoiBenPhai & " "
TT:
            sKyTu1 = "":    sKyTu2 = "":    sKyTu3 = "":    sChuoiBenPhai = ""
        Next
        zToUnicode = IIf(Right$(sTam, 1) = " ", Left$(sTam, Len(sTam) - 1), sTam)
End Function

'HAM CHUYEN DOI CHUOI UNICODE SANG CHUOI ANSI
Public Function UNICODE_To_ANSI(ByRef sString As String) As String
    Dim sTam As String
    Dim i As Long, ArrChuoiXuLy() As String
    Dim k As Long, j As Long
    Dim bThay As Boolean
        
        If Trim$(sString) = "" Then UNICODE_To_ANSI = sString:   Exit Function
        
        ArrFromCode = Split(sUNICODE, ",")
        ArrToCode = Split(sANSI, ",")
        ArrChuoiXuLy = Split(sString, " ")
        
        For i = 0 To UBound(ArrChuoiXuLy)   'cho vong lap chay den het cac tu trong 1 chuoi (cac tu cach nhau 1 khoang trang)
            For j = 1 To Len(ArrChuoiXuLy(i))  'cho vong lap chay tu ky tu trong 1 tu
                'chi kiem tra cac ky tu nam sau z ma thoi neu la cac ky tu tu A -> Z, a -> z thi khong kiem tra
                If AscW(Mid$(ArrChuoiXuLy(i), j, 1)) = Asc("?") Or AscW(Mid$(ArrChuoiXuLy(i), j, 1)) > Asc("z") Then
                    For k = 0 To UBound(ArrFromCode)
                        If Mid$(ArrChuoiXuLy(i), j, 1) = ChrW$(CLng("&H" & ArrFromCode(k))) Then      'neu tim thay ky tu can thay the trong chuoi can chuyen doi
                            bThay = True 'tha^'y ky tu can chuyen doi
                            Exit For
                        End If
                    Next
                End If
                
                If bThay = True Then
                    sTam = sTam & ArrToCode(k)  'thay ky tu trong chuoi can chuyen doi thanh ky tu trong bang ma sau khi chuyen
                    bThay = False
                Else
                    sTam = sTam & Mid$(ArrChuoiXuLy(i), j, 1)
                End If
            Next
            sTam = sTam & " "   'sau khi kiem tra xong 1 chu thi them vao sau no 1 khoang trang
        Next
        
'cat bo 1 ky tu khoang trang du ra phia sau chuoi sau khi xu ly xong
        UNICODE_To_ANSI = IIf(Right$(sTam, 1) = " ", Left$(sTam, Len(sTam) - 1), sTam) ' Replace$(sTam, "  ", " ")
End Function

Private Function HaveNumber(sString As String) As Long
    Dim i As Long
    Dim sKytu As String
    
        For i = 1 To Len(sString)
            sKytu = Mid(sString, i, 1)
            If IsNumeric(sKytu) Then HaveNumber = i
            If HaveNumber > 0 Then Exit Function
        Next
End Function

Private Function ChangeString(sString As String) As String
    Dim k As Long, bThayDoi As Boolean

        For k = 0 To UBound(ArrToCode)
            If sString = ArrFromCode(k) Then
                ChangeString = ChangeString & ChrW$(CLng("&H" & ArrToCode(k)))
                bThayDoi = True
                Exit For
            End If
        Next
        If bThayDoi = False Then ChangeString = ChangeString & sString
End Function


