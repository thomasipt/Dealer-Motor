Attribute VB_Name = "Module1"
Public CN As String
Public OYEN, STS_Nama, STS_Mesin, STS_Rangka, STS_Biaya, NamaHutang, NoHutang, TBACKUP, KRK, Status, NoUrut, NoPinjaman, No_Nas, Nama_Nas, NO_RANGKA, JML_UNIT, FAKTUR, TANGGAL_FAKTUR, G_DEBET, G_CREDIT, F_DEBET, F_CREDIT, N_ALAMAT, N_CCAB, KodeBR, NamaBR, K, BR, KB, Tanggal, CodeCab, Operator, CodeBag, gdebet, GCredit, MutasiD, MutasiC, Skin, NToko, NAlamat, NMOtto, NTelepon As String

Sub Main()
CN = "DSN=DEALER;DRIVER={Microsoft Access Driver};Server=CENTRAL;UID= ;PWD= ;Database = DEALER.mdb;"
LOGIN.Show
End Sub

Public Sub ClearTextBoxes(frmClearMe As Form)
Dim txt As Control
For Each txt In frmClearMe
  If TypeOf txt Is TextBox Then txt.Text = ""
Next
End Sub

Public Function Digit(panjang, Nilai As Double) As String
Dim Y, NilaiP As Double
Dim Kar, NilaiS As String

If panjang <= 0 Then panjang = 1

NilaiS = Trim(Str(Nilai))
NilaiP = Len(NilaiS)
If NilaiP >= panjang Then panjang = NilaiP

Kar = " "
For Y = 1 To panjang
    Kar = Trim(Kar) + "0"
Next
If (panjang - NilaiP) <= 0 Then
    Digit = NilaiS
Else
    Digit = Mid(Kar, 1, (panjang - NilaiP)) + NilaiS
End If
End Function

Public Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Public Function SumHari(Dari, Ke As Date) As Integer
If Ke - Dari <= 1 Then
    SumHari = 1
Else
    SumHari = Ke - Dari
End If
End Function

Public Function Sisip(Kar As String, Posisi As Integer, Kar2 As String) As String
Dim PJ As Integer
Dim Akhir As String
Dim depan As String
PJ = Len(Kar)
If Len(Kar) < Len(Kar2) Then
    Sisip = Kar2
Else
    If Posisi = 1 Then Sisip = Kar2 + Mid(Kar, 2, PJ - 1)
    If Posisi > 1 And Posisi < PJ Then
        depan = Mid(Kar, 1, Posisi - 1)
        Akhir = Mid(Kar, Posisi + 1, PJ - Posisi)
        Sisip = depan + Kar2 + Akhir
    End If
    If Posisi = PJ Then Sisip = Mid(Kar, 1, Posisi - 1) + Kar2
End If
End Function
Public Function Satuan(ByVal Nilai As Currency) As String
Select Case Nilai
    Case 1: Satuan = "SATU "
    Case 2: Satuan = "DUA "
    Case 3: Satuan = "TIGA "
    Case 4: Satuan = "EMPAT "
    Case 5: Satuan = "LIMA "
    Case 6: Satuan = "ENAM "
    Case 7: Satuan = "TUJUH "
    Case 8: Satuan = "DELAPAN "
    Case 9: Satuan = "SEMBILAN "
End Select
End Function
Public Function Ribuan(ByVal Bilangan As Double) As String
Dim A, B As Currency
Dim C As String

C = ""
A = Bilangan \ 1000
B = Bilangan Mod 1000
If A > 1 Then C = C + Satuan(A) + "RIBU "
If A = 1 Then C = C + "SERIBU "

A = B \ 100
B = B Mod 100
If A > 1 Then C = C + Satuan(A) + "RATUS "
If A = 1 Then C = C + "SERATUS "

A = B \ 10
B = B Mod 10
If A > 1 Then C = C + Satuan(A) + "PULUH "
If A = 1 Then
    If B = 0 Then Ribuan = C + "SEPULUH" '"SERATUS "
    If B = 1 Then Ribuan = C + "SEBELAS "
    If B > 1 Then Ribuan = C + Satuan(B) + "BELAS "
Else
    Ribuan = C + Satuan(B)
End If
End Function

Public Function Terbilang(Bilangan As Double) As String
Dim A, D As Double
Dim B, e, F As Double
Dim C As String
If Bilangan > 2000000000 Then
    C = "#"
    
    D = Mid(Bilangan, 1, 7)
    A = D \ 1000000
    B = D Mod 1000000
    If A > 0 Then C = Ribuan(A) + "MILYAR "
    
    e = Mid(Bilangan, 2, 10)
    A = e \ 1000000
    B = e Mod 1000000
    If A > 0 Then C = C + Ribuan(A) + "JUTA "
    
    F = Mid(Bilangan, 5, 10)
    A = F \ 1000
    B = F Mod 1000
    If A > 0 Then C = C + Ribuan(A) + "RIBU "
    If A = 1 Then C = C + "SERIBU "
Terbilang = C + Ribuan(B) + "RUPIAH#"
Else
    C = "#"
    A = Bilangan \ 1000000000
    B = Bilangan Mod 1000000000
    If A > 0 Then C = Ribuan(A) + "MILYAR "
    
    A = B \ 1000000
    B = B Mod 1000000
    If A > 0 Then C = C + Ribuan(A) + "JUTA "
    
    A = B \ 1000
    B = B Mod 1000
    If A > 1 Then C = C + Ribuan(A) + "RIBU "
    If A = 1 Then C = C + "SERIBU "
Terbilang = C + Ribuan(B) + "RUPIAH#"

End If
End Function

Public Function BulanStr(ByVal Bulan As Currency) As String
Select Case Bulan
    Case 1: BulanStr = "Januari"
    Case 2: BulanStr = "Februari"
    Case 3: BulanStr = "Maret"
    Case 4: BulanStr = "April"
    Case 5: BulanStr = "M e i"
    Case 6: BulanStr = "J u n i"
    Case 7: BulanStr = "J u l i"
    Case 8: BulanStr = "Agustus"
    Case 9: BulanStr = "September"
    Case 10: BulanStr = "Oktober"
    Case 11: BulanStr = "Nopember"
    Case 12: BulanStr = "Desember"
End Select
End Function

Public Function BlkKoma(Bilangan As Double) As String
Dim A, D As Double
Dim B, e, F As Double
Dim C As String
If Bilangan > 2000000000 Then
    C = ""
    
    D = Mid(Bilangan, 1, 7)
    A = D \ 1000000
    B = D Mod 1000000
    If A > 0 Then C = Ribuan(A) + "MILYAR "
    
    e = Mid(Bilangan, 2, 10)
    A = e \ 1000000
    B = e Mod 1000000
    If A > 0 Then C = C + Ribuan(A) + "JUTA "
    
    F = Mid(Bilangan, 5, 10)
    A = F \ 1000
    B = F Mod 1000
    If A > 0 Then C = C + Ribuan(A) + "RIBU "
    If A = 1 Then C = C + "SERIBU "
BlkKoma = C + Ribuan(B)
Else
    C = ""
    A = Bilangan \ 1000000000
    B = Bilangan Mod 1000000000
    If A > 0 Then C = Ribuan(A) + "MILYAR "
    
    A = B \ 1000000
    B = B Mod 1000000
    If A > 0 Then C = C + Ribuan(A) + "JUTA "
    
    A = B \ 1000
    B = B Mod 1000
    If A > 1 Then C = C + Ribuan(A) + "RIBU "
    If A = 1 Then C = C + "SERIBU "
BlkKoma = C + Ribuan(B)
End If
End Function
