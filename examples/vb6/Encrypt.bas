Attribute VB_Name = "Encrypt"
Public ENCRYPTION_NONE As Long
Public ENCRYPTION_WEAK_XOR As Long
Public ENCRYPTION_OFFICE_97_2000_COMPATIBLE As Long


Sub Initialize()
        ENCRYPTION_NONE = 0
        ENCRYPTION_WEAK_XOR = 1
        ENCRYPTION_OFFICE_97_2000_COMPATIBLE = 2
End Sub
