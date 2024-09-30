Function ConvertUTCtoKST(utcTime As Date) As Date
    ' UTC+9 (한국 표준시)로 변환
    ConvertUTCtoKST = utcTime + (9 / 24)
End Function
