Public Moc as Object
Public Klucz as object
Public Czest as object
Public IP as object
Public imageCNBOP as object
Public imageEX as object
Public CNBOP as object
Public KDBnumer as object
Public OznaczEX as object
Public MinTemp as object
Public MaxTemp as object
Public cosinusFi as object
Public Napiecie as object
Public NumerCE as object
Public OznaczenieCNBOP1 as object
Public OznaczenieCNBOP2 as object
Public OznaczenieCNBOP3 as object
Public OznaczenieCNBOP4 as object
Public KratkiCNBOP as object
Public KluczText as string
Public IndexSN as object
Public SerialNumber as object
Public Index as object
Public ROK as object
Public currentYear as Integer
Public ObrazEx as object
Public ObrazCNBOP as object
Public Oznacz2Klasa as object
Public LicznikWierszy as Integer
Public ZnakCzest as Object
Public Battery as Object
Public OznaczenieSFRC as object
Public Oznacz3Klasa as object
Sub Main
    DIM Form as Object
    
    Form = ThisComponent.Drawpage.forms.getByName("Formularz")
    Klucz = Form.getByName("KluczProduktu")
    Czest = Form.getByName("Czestotliwosc")
    IP = Form.getByName("IP")
    CNBOP = Form.getByName("CnbopNumer")
    KDBnumer = Form.getByName("KDB")
    OznaczEX = Form.getByName("OznaczenieEX")
    MinTemp = Form.getByName("MinTemp")
    MaxTemp = Form.getByName("MaxTemp")
    cosinusFi = Form.getByName("CosFi")
    Napiecie = Form.getByName("Napiecie")
    Moc = Form.getByName("Moc")
    NumerCE = Form.getByName("NumerCE")
    OznaczenieCNBOP1 = Form.getByName("OznaczenieCNBOP1")
    OznaczenieCNBOP2 = Form.getByName("OznaczenieCNBOP2")
    OznaczenieCNBOP3 = Form.getByName("OznaczenieCNBOP3")
    OznaczenieCNBOP4 = Form.getByName("OznaczenieCNBOP4")
    SerialNumber = Form.getByName("SerialNumber")
    IndexSN = Form.getByName("IndexSN")
    Index = Form.getByName("Index")
    ROK = Form.getByName("ROK")
    ObrazEx = Form.getByName("ObrazEx")
    ObrazCNBOP = Form.getByName("ObrazCNBOP")
    Oznacz2Klasa = Form.getByName("Oznacz2Klasa")
    ZnakCzest = Form.getByName("ZnakCzest")
    Battery = Form.getByName("Battery")
	OznaczenieSFRC = Form.getByName("OznaczenieSFRC")
    Oznacz3Klasa = Form.getByName("TRZECIAKLASA")

    ObrazEx.EnableVisible = False
    Oznacz2Klasa.EnableVisible = False
    Czest.Text=""
    IP.Text=""
    KDBnumer.Text=""
    OznaczEX.Text=""
    cosinusFi.Text=""
    Napiecie.Text=""
    NumerCE.Text=""
    ObrazCNBOP.EnableVisible = False
    CNBOP.Text=""
    MinTemp.Text=""
    MaxTemp.Text=""
    OznaczenieCNBOP1.Text=""
    OznaczenieCNBOP2.Text=""
    OznaczenieCNBOP3.Text=""
    OznaczenieCNBOP4.Text=""
    OznaczenieCNBOP1.EnableVisible=False
    OznaczenieCNBOP2.EnableVisible=False
    OznaczenieCNBOP3.EnableVisible=False
    OznaczenieCNBOP4.EnableVisible=False
    Moc.Text=""   
    Battery.Text=""  
    OznaczenieSFRC.Text=""  
    Oznacz3Klasa.EnableVisible = False
    

     If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Czest.Text="0/50/60 Hz"
        End If
      
         

    If Instr(Klucz.Text, "EXL210LED")>0 Then
    ZnakCzest.Text = "f:"
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 3G Ex ec op is IIC T4 Gc" & chr(13) & "II 2D Ex tb op is IIIC T70°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            OznaczEX.Text="II 3G Ex ec IIC T4 Gc" & chr(13) & "II 3D Ex tc IIIC T70°C IP66/67 Dc"
            KDBnumer.Text=""
            NumerCE.Text=""
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3549/2019"
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3549/2019"
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "0600-E2")>0 Then
            Moc.Text="17,4W LED"
        End If
        If Instr(Klucz.Text, "0600-E4")>0 Then
            Moc.Text="33,4W LED"
        End If
        If Instr(Klucz.Text, "1200-E4")>0 Then
            Moc.Text="40,4W LED"
        End If
        If Instr(Klucz.Text, "1200-E8")>0 Then
            Moc.Text="65,7W LED"
            MaxTemp.Text="40"
        End If
        If Instr(Klucz.Text, "1500-E6")>0 Then
            Moc.Text="50,4W LED"
        End If
         If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If
    Else
    End If
   
    
    
    If Instr(Klucz.Text, "EXL310LED")>0 Then
    ZnakCzest.Text = "f:"
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 3G Ex ec op is IIC T4 Gc" & chr(13) & "II 2D Ex tb op is IIIC T70°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="50"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            OznaczEX.Text="II 3G Ex ec IIC T4 Gc" & chr(13) & "II 3D Ex tc IIIC T70°C IP66/67 Dc"
            KDBnumer.Text=""
            NumerCE.Text=""
            ObrazCNBOP.EnableVisible = True
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****E**"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "0600-E2")>0 Then
            Moc.Text="17,4W LED"
        End If
        If Instr(Klucz.Text, "0600-E4")>0 Then
            Moc.Text="33,4W LED"
        End If
        If Instr(Klucz.Text, "1200-E4")>0 Then
            Moc.Text="40,4W LED"
        End If
        If Instr(Klucz.Text, "1200-E8")>0 Then
            Moc.Text="65,7W LED"
        End If
    Else
    End If
   
    
    
    If Instr(Klucz.Text, "EXL380LED")>0 Then
    ZnakCzest.Text = "f:"
    Czest.Text="50/60 Hz"
    IP.Text="65"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 3G Ex ec op is IIC T4 Gc" & chr(13) & "II 2D Ex tb op is IIIC T80°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="40"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "045-E4")>0 Then
            Moc.Text="53 W LED"
        End If
        If Instr(Klucz.Text, "090-E8")>0 Then
            Moc.Text="98 W LED"
        End If
        If Instr(Klucz.Text, "130-E12")>0 Then
            Moc.Text="146 W LED"
        End If
    Else
    End If
 
    
    
    If Instr(Klucz.Text, "EXL390LED")>0 Then
    'Dim pradEXL390LED As String'
    'pradEXL390LED = InputBox ("Jaki prąd sterowania 250mA czy 350mA:","Drogi uzytkowniku", 250)'
    ZnakCzest.Text = "f:"
    Czest.Text="0/50/60 Hz"
    IP.Text="65"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 3G Ex ec op is IIC T4 Gc" & chr(13) & "II 2D Ex tb op is IIIC T70°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="50"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            OznaczEX.Text="II 3G Ex ec IIC T4 Gc" & chr(13) & "II 3D Ex tc IIIC T65°C IP65 Dc"
            KDBnumer.Text=""
            NumerCE.Text=""
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            MaxTemp.Text="50"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        'If Instr(Klucz.Text, "0600-E4")>0 And pradEXL390LED = 250 Then
         '   Moc.Text="35,0W LED"
        'End If
       ' If Instr(Klucz.Text, "0600-E4")>0 And pradEXL390LED = 350 Then
        '    Moc.Text="47,0W LED"
        'End If
        'If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 250 Then
       '     Moc.Text="27,1W LED"
       ' End If
        'If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 350 Then
        '    Moc.Text="36,6W LED"
       ' End If
        'If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 250 Then
        '    Moc.Text="52,0W LED"
       ' End If
        'If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 350 Then
       '     Moc.Text="70,5W LED"
       ' End If
    Else
    End If
    


  
    If Instr(Klucz.Text, "EXF200LED")>0 Then
    ZnakCzest.Text = ""
    Czest.Text="220-250VDC"
    IP.Text="66/67"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 2G Ex eb mb op is IIC T5 Gb" & chr(13) & "II 2D Ex tb op is IIIC T55°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="110-254VAC,"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "A3")>0 Then
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
             OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiCd 4Ah 4,8V" 
        End If
         If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "0600-F1")>0 Then
            Moc.Text="21,2W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "0600-F2")>0 Then
            Moc.Text="38,6W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-F2")>0 Then
            Moc.Text="38,6W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-F4")>0 Then
            Moc.Text="78,1W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-G4")>0 Then
            Moc.Text="47,0W LED"
            MaxTemp.Text="55"
        End If
        If Instr(Klucz.Text, "0600-G2")>0 Then
            Moc.Text="25,3W LED"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "1200-G2")>0 Then
            Moc.Text="25,3W LED"
            MaxTemp.Text="60"
        End If
    Else
    End If
   
   
   
    If Instr(Klucz.Text, "EXF250LED")>0 Then
    ZnakCzest.Text = ""
    Czest.Text="220-250VDC"
    IP.Text="66/67"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 2G Ex eb mb op is IIC T5 Gb" & chr(13) & "II 2D Ex tb op is IIIC T55°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="110-254VAC,"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "A3")>0 Then
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3529/2019"
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiCd 4Ah 4,8V" 
        End If
           If Instr(Klucz.Text, "ZB")>0 Then
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3529/2019"
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "0600-F1")>0 Then
            Moc.Text="21,2W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "0600-F2")>0 Then
            Moc.Text="38,6W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-F2")>0 Then
            Moc.Text="38,6W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-F4")>0 Then
            Moc.Text="78,1W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-G4")>0 Then
            Moc.Text="47,0W LED"
            MaxTemp.Text="55"
        End If
        If Instr(Klucz.Text, "0600-G2")>0 Then
            Moc.Text="25,3W LED"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "1200-G2")>0 Then
            Moc.Text="25,3W LED"
            MaxTemp.Text="60"
        End If
    Else
    End If



    If Instr(Klucz.Text, "EXF300LED")>0 Then
    ZnakCzest.Text = ""
    Czest.Text="220-250VDC"
    IP.Text="67"
    KDBnumer.Text="KDB15ATEX0049X"
    OznaczEX.Text="II 2G Ex eb mb op is IIC T5 Gb" & chr(13) & "II 2D Ex tb op is IIIC T55°C Db"
    MinTemp.Text="-40"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="110-254VAC,"
    NumerCE.Text="1453"
    ObrazEx.EnableVisible = True
        If Instr(Klucz.Text, "A3")>0 Then
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3529/2019"
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiCd 4Ah 4,8V" 
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "0600-F1")>0 Then
            Moc.Text="21,2W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "0600-FX2")>0 Then
            Moc.Text="38,6W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-F2")>0 Then
            Moc.Text="38,6W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-FX4")>0 Then
            Moc.Text="78,1W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "1200-GX4")>0 Then
            Moc.Text="47,0W LED"
            MaxTemp.Text="55"
        End If
        If Instr(Klucz.Text, "0600-GX2")>0 Then
            Moc.Text="25,3W LED"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "1200-G2")>0 Then
            Moc.Text="25,3W LED"
            MaxTemp.Text="60"
        End If
    Else
    End If
  
  
  
    If Instr(Klucz.Text, "INV320LED")>0 Then
     ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = True
    Czest.Text="0/50/60 Hz"
    IP.Text="65"
    MinTemp.Text="-40"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="2910/2017"
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
        If Instr(Klucz.Text, "A3S")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="2910/2017"
            MinTemp.Text="-20"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="*BC*EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: LiFePO4" 
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "B2-1")>0 Then
            Moc.Text="36,6W LED"
        End If
        If Instr(Klucz.Text, "B2-2")>0 Then
            Moc.Text="42,1W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-2")>0 Then
            Moc.Text="42,1W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
        End If
         If Instr(Klucz.Text, "B4-2")>0 Then
            Moc.Text="83,2W LED"
        End If
         If Instr(Klucz.Text, "J4M2-1")>0 Then
            Moc.Text="45,4W LED"
        End If
         If Instr(Klucz.Text, "J4M2-2")>0 Then
            Moc.Text="55,9W LED"
            MaxTemp.Text="40"
        End If
         If Instr(Klucz.Text, "J4M2-3")>0 Then
            Moc.Text="72,7W LED"
        End If
          If Instr(Klucz.Text, "SF")>0 Then
            OznaczenieSFRC.Text="SF"
        End If
         If Instr(Klucz.Text, "RC")>0 Then
            OznaczenieSFRC.Text="RC"
        End If
         If Instr(Klucz.Text, "B2-1")>0 And Instr(Klucz.Text, "RC")>0 Then
            Moc.Text="36,6W LED"
            MaxTemp.Text="40"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INV360LED")>0 Or Instr(Klucz.Text, "INP360LED")>0 Then
    ZnakCzest.Text = "f:"
    Czest.Text="0/50/60 Hz"
    Oznacz2Klasa.EnableVisible = False
    IP.Text="67"
    MinTemp.Text="-30"
    MaxTemp.Text="30"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "R2-1")>0 Then
            Moc.Text="8,7W LED"
        End If
        If Instr(Klucz.Text, "L2-1")>0 Then
            Moc.Text="11,3W LED"
        End If
        If Instr(Klucz.Text, "L2-3")>0 Then
            Moc.Text="14,2W LED"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INP320LED")>0 Then
     ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="65"
    MinTemp.Text="-30"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "B2-1")>0 Then
            Moc.Text="36,6W LED"
        End If
        If Instr(Klucz.Text, "B2-2")>0 Then
            Moc.Text="42,1W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-2")>0 Then
            Moc.Text="42,1W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
        End If
         If Instr(Klucz.Text, "B4-2")>0 Then
            Moc.Text="83,2W LED"
        End If
         If Instr(Klucz.Text, "J4M2-1")>0 Then
            Moc.Text="45,4W LED"
        End If
         If Instr(Klucz.Text, "J4M2-2")>0 Then
            Moc.Text="55,9W LED"
            MaxTemp.Text="40"
        End If
          If Instr(Klucz.Text, "J2M1-1")>0 Then
            Moc.Text="19+8W LED"
        End If
         If Instr(Klucz.Text, "J2M1-3")>0 Then
            Moc.Text="27+8W LED"
        End If
          If Instr(Klucz.Text, "J4M1-1")>0 Then
            Moc.Text="37+8W LED"
        End If
         If Instr(Klucz.Text, "J4M1-2")>0 Then
            Moc.Text="42+8W LED"
        End If
         If Instr(Klucz.Text, "J4M2-3")>0 Then
            Moc.Text="72,7W LED"
        End If
          If Instr(Klucz.Text, "SF")>0 Then
            OznaczenieSFRC.Text="SF"
        End If
         If Instr(Klucz.Text, "RC")>0 Then
            OznaczenieSFRC.Text="RC"
        End If
          If Instr(Klucz.Text, "B2-1")>0 And Instr(Klucz.Text, "RC")>0 Then
            Moc.Text="36,6W LED"
            MaxTemp.Text="40"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS250LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
        If Instr(Klucz.Text, "J1-0")>0 Then
            Moc.Text="6,0W LED"
        End If
        If Instr(Klucz.Text, "J1-2")>0 Then
            Moc.Text="10,5W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
        End If
        If Instr(Klucz.Text, "ZBC")>0 Then
            MinTemp.Text="-20"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If

    Else
    End If



    If Instr(Klucz.Text, "INS270LED")>0 Then
     ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
        End If
         If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS310LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="65"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS350LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="68"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="26,7W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="50"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="37,5W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="60"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="49,7W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="60"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="59,5W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
        End If
    Else
    End If



    If Instr(Klucz.Text, "HPL425LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="65"
    MinTemp.Text="-25"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "050")>0 Then
            Moc.Text="49,0W LED"
            MaxTemp.Text="50"
        End If
        If Instr(Klucz.Text, "060")>0 Then
            Moc.Text="62,0W LED"
            MaxTemp.Text="40"
        End If
        If Instr(Klucz.Text, "080")>0 Then
            Moc.Text="78,8W LED"
            MaxTemp.Text="35"
        End If
    Else
    End If



    If Instr(Klucz.Text, "HPL430LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = True
    Czest.Text="50/60 Hz"
    IP.Text="66"
    MinTemp.Text="-30"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "050")>0 Then
            Moc.Text="53,6W LED"
        End If
        If Instr(Klucz.Text, "100")>0 Then
            Moc.Text="101,7W LED"
        End If
        If Instr(Klucz.Text, "150")>0 Then
            Moc.Text="153,1W LED"
        End If
    Else
    End If



    If Instr(Klucz.Text, "HPL440LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="50/60 Hz"
    IP.Text="65"
    MinTemp.Text="-40"
    MaxTemp.Text="40"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "090-H4")>0 Then
            Moc.Text="70W LED"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "130-H8")>0 Then
            Moc.Text="128W LED"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "170-H8")>0 Then
            Moc.Text="173W LED"
            MaxTemp.Text="55"
        End If
         If Instr(Klucz.Text, "250-H8")>0 Then
            Moc.Text="235W LED"
            MaxTemp.Text="40"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS230LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
      If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3660/2019"
            MinTemp.Text="0"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            ObrazCNBOP.EnableVisible = True
            CNBOP.Text="3660/2019"
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "P2-1")>0 Then
            Moc.Text="19,5W LED"
        End If
        If Instr(Klucz.Text, "P4-1")>0 Then
            Moc.Text="35,9W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="40"
        End If
            If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If
        If Instr(Klucz.Text, "II")>0 Then
           Oznacz2Klasa.EnableVisible = True
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS300LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="44"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "0900-A3")>0 Then
            Moc.Text="38,3W LED"
        End If
        If Instr(Klucz.Text, "1500-A1B2")>0 Then
            Moc.Text="57,0W LED"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS340LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Oznacz3Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
      If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            MinTemp.Text="0"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="40"
        End If
            If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If
        If Instr(Klucz.Text, "II")>0 Then
           Oznacz2Klasa.EnableVisible = True
        End If
          If Instr(Klucz.Text, "D2")>0 Then
            Moc.Text="20,0W LED"
            Czest.Text="0 Hz"
            Napiecie.Text="24VDC"
             Oznacz3Klasa.EnableVisible = True
              Oznacz2Klasa.EnableVisible = False
        End If 
         If Instr(Klucz.Text, "D4")>0 Then
            Moc.Text="40,0W LED"
            Czest.Text="0 Hz"
            Napiecie.Text="24VDC"
             Oznacz3Klasa.EnableVisible = True
              Oznacz2Klasa.EnableVisible = False
        End If 
         If Instr(Klucz.Text, "D2C1")>0 Then
            Moc.Text="19+9W LED"
             Napiecie.Text="24VDC"
             Czest.Text="0 Hz"
             Oznacz3Klasa.EnableVisible = True
             Oznacz2Klasa.EnableVisible = False
        End If
          If Instr(Klucz.Text, "C2")>0 Then
            Moc.Text="18,0W LED"
            Czest.Text="0 Hz"
            Napiecie.Text="24VDC"
             Oznacz3Klasa.EnableVisible = True
              Oznacz2Klasa.EnableVisible = False
        End If
        
    Else
    End If



    If Instr(Klucz.Text, "INS352LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="68"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="60"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="26,7W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="50"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="37,5W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="60"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="49,7W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="60"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="59,5W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
        End If
           If Instr(Klucz.Text, "D2")>0 Then
            Moc.Text="20,0W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="60"
            Oznacz3Klasa.EnableVisible = True
        End If
           If Instr(Klucz.Text, "D4")>0 Then
            Moc.Text="39,3W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="60"
            Oznacz3Klasa.EnableVisible = True
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS370LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    MinTemp.Text="-25"
    IP.Text="65"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "B8-1")>0 Then
            Moc.Text="141,6W LED"
            MaxTemp.Text="40"
        End If
        If Instr(Klucz.Text, "A2B4-0")>0 Then
            Moc.Text="50,0W LED"
            MaxTemp.Text="45"
        End If
         If Instr(Klucz.Text, "A2B4-1")>0 Then
            Moc.Text="89,9W LED"
            MaxTemp.Text="45"
        End If
         If Instr(Klucz.Text, "A4B8-0")>0 Then
            Moc.Text="99,0W LED"
            MaxTemp.Text="45"
        End If
           If Instr(Klucz.Text, "A4B8-1")>0 Then
            Moc.Text="178W LED"
            MaxTemp.Text="35"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS370LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    MinTemp.Text="-25"
    IP.Text="65"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "B8-1")>0 Then
            Moc.Text="141,6W LED"
            MaxTemp.Text="40"
        End If
        If Instr(Klucz.Text, "A2B4-0")>0 Then
            Moc.Text="50,0W LED"
            MaxTemp.Text="45"
        End If
         If Instr(Klucz.Text, "A2B4-1")>0 Then
            Moc.Text="89,9W LED"
            MaxTemp.Text="45"
        End If
         If Instr(Klucz.Text, "A4B8-0")>0 Then
            Moc.Text="99,0W LED"
            MaxTemp.Text="45"
        End If
           If Instr(Klucz.Text, "A4B8-1")>0 Then
            Moc.Text="178W LED"
            MaxTemp.Text="35"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INS395LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    MinTemp.Text="-25"
    IP.Text="54"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="45"
        End If
        If Instr(Klucz.Text, "B8-1")>0 Then
            Moc.Text="141,6W LED"
            MaxTemp.Text="45"
        End If
         If Instr(Klucz.Text, "A2B4-1")>0 Then
            Moc.Text="89,9W LED"
            MaxTemp.Text="45"
        End If
           If Instr(Klucz.Text, "A4B8-1")>0 Then
            Moc.Text="178W LED"
            MaxTemp.Text="40"
        End If
    Else
    End If



    If Instr(Klucz.Text, "INX230LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    IP.Text="66/67"
    cosinusFi.Text="0,95"
        If Instr(Klucz.Text, "W1-0")>0 Then
            Moc.Text="11,1W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="70"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-0")>0 Then
            Moc.Text="22,3W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="70"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W1-1")>0 Then
            Moc.Text="20,4W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="65"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-1")>0 Then
            Moc.Text="41,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="65"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W1-2")>0 Then
            Moc.Text="35,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="50"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-2")>0 Then
            Moc.Text="71,1W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="50"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "Y2-3")>0 Then
            Moc.Text="27,2W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "X2-1")>0 Then
            Moc.Text="37,0W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "Y4-2")>0 Then
            Moc.Text="42,0W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "Y4-3")>0 Then
            Moc.Text="53,5W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
        If Instr(Klucz.Text, "A3")>0 Then
            Napiecie.Text="230V"
            Czest.Text="50/60 Hz"
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True 
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
        If Instr(Klucz.Text, "A3S")>0 Then
            Napiecie.Text="230V"
            Czest.Text="50/60 Hz"
            MinTemp.Text="-20"
            MaxTemp.Text="50"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="*BC*EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True 
            Battery.Text="Bat: LiFePO4" 
        End If
    Else
    End If



    If Instr(Klucz.Text, "INX340LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    IP.Text="67"
    cosinusFi.Text="0,95"
        If Instr(Klucz.Text, "W1-0")>0 Then
            Moc.Text="11,1W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="75"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-0")>0 Then
            Moc.Text="22,3W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="75"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W1-1")>0 Then
            Moc.Text="20,1W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="70"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-1")>0 Then
            Moc.Text="40,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="70"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W1-2")>0 Then
            Moc.Text="35,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="60"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-2")>0 Then
            Moc.Text="71,2W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="60"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
            If Instr(Klucz.Text, "W1-3")>0 Then
            Moc.Text="39,9W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="50"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "W2-3")>0 Then
            Moc.Text="79,8W LED"
            MinTemp.Text="-25"
            MaxTemp.Text="50"
            Napiecie.Text="120-277V"
            Czest.Text="50/60 Hz"
        End If
           If Instr(Klucz.Text, "Y2-3")>0 Then
            Moc.Text="27,2W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "X2-1")>0 Then
            Moc.Text="37,0W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "Y4-2")>0 Then
            Moc.Text="42,0W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "Y4-3")>0 Then
            Moc.Text="53,5W LED"
            MinTemp.Text="-40"
            MaxTemp.Text="40"
            Napiecie.Text="230V"
            Czest.Text="0/50/60 Hz"
        End If
         If Instr(Klucz.Text, "A3")>0 Then
            Napiecie.Text="230V"
            Czest.Text="50/60 Hz"
            MinTemp.Text="0"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True 
            Battery.Text="Bat: NiMH 4Ah 4,8V" 
        End If
        If Instr(Klucz.Text, "A3S")>0 Then
            Napiecie.Text="230V"
            Czest.Text="50/60 Hz"
            MinTemp.Text="-20"
            MaxTemp.Text="50"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="*BC*EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True 
            Battery.Text="Bat: LiFePO4" 
        End If
    Else
    End If



    If Instr(Klucz.Text, "INX385LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    IP.Text="67"
    MinTemp.Text="-30"
    Napiecie.Text="230V"
    Czest.Text="0/50/60 Hz"
    cosinusFi.Text="0,95"
        If Instr(Klucz.Text, "X2-0")>0 Then
            Moc.Text="23,8W LED"
             MaxTemp.Text="65"
        End If
        If Instr(Klucz.Text, "X2-1")>0 Then
            Moc.Text="36,1W LED"
             MaxTemp.Text="55"
        End If
    Else
    End If
    
     If Instr(Klucz.Text, "INS340")>0  And Instr(Klucz.Text, "-LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
      If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            MinTemp.Text="0"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
         If Instr(Klucz.Text, "A3S")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="-20"
            MaxTemp.Text="45"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="*BC*EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: LiFePO4" 
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="27,2W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="40"
        End If
            If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If
        If Instr(Klucz.Text, "II")>0 Then
           Oznacz2Klasa.EnableVisible = True
        End If
    Else
    End If
    
     If Instr(Klucz.Text, "INS230")>0 And Instr(Klucz.Text, "-LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
      If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "P2-1")>0 Then
            Moc.Text="19,5W LED"
        End If
        If Instr(Klucz.Text, "P4-1")>0 Then
            Moc.Text="35,9W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="40"
        End If
            If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If
        If Instr(Klucz.Text, "II")>0 Then
           Oznacz2Klasa.EnableVisible = True
        End If
    Else
    End If
    
    If Instr(Klucz.Text, "INS341LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
      If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
	If Instr(Klucz.Text, "B1-1")>0 Then
            Moc.Text="17,2W LED"
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,8W LED"
        End If
        If Instr(Klucz.Text, "P2-1")>0 Then
            Moc.Text="19,5W LED"
        End If
        If Instr(Klucz.Text, "P4-1")>0 Then
            Moc.Text="35,9W LED"
        End If
        If Instr(Klucz.Text, "J4-1")>0 Then
            Moc.Text="36,6W LED"
        End If
         If Instr(Klucz.Text, "J4-3")>0 Then
            Moc.Text="53,5W LED"
        End If
         If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="72,1W LED"
            MaxTemp.Text="40"
        End If
            If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If
        If Instr(Klucz.Text, "II")>0 Then
           Oznacz2Klasa.EnableVisible = True
        End If
        If Instr(Klucz.Text, "D2")>0 Then
            Moc.Text="20,0W LED"
            Czest.Text="0 Hz"
            Napiecie.Text="24VDC"
        End If 
         If Instr(Klucz.Text, "D4")>0 Then
            Moc.Text="40,0W LED"
            Czest.Text="0 Hz"
            Napiecie.Text="24VDC"
        End If   
    Else
    End If
    
        If Instr(Klucz.Text, "STL430LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="50/60 Hz"
    IP.Text="66"
    MinTemp.Text="-30"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "100")>0 Then
            Moc.Text="101,7W LED"
        End If
        If Instr(Klucz.Text, "154")>0 Then
            Moc.Text="153,1W LED"
        End If
        If Instr(Klucz.Text, "50")>0 Then
            Moc.Text="53,6W LED"
        End If
    Else
    End If

    If Instr(Klucz.Text, "INS240LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="68"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "N2-2")>0 Then
            Moc.Text="12,3W LED"
        End If
        If Instr(Klucz.Text, "N3-2")>0 Then
            Moc.Text="17,9W LED"
        End If
        If Instr(Klucz.Text, "N4-2")>0 Then
            Moc.Text="23,6W LED"
        End If
	If Instr(Klucz.Text, "A1B1-2")>0 Then
            Moc.Text="31,8W LED"
        End If
	If Instr(Klucz.Text, "C2-0")>0 Then
            Moc.Text="18,0W LED"
		Napiecie.Text="24VDC"
		Czest.Text="0 Hz"
	Oznacz3Klasa.EnableVisible = True
        End If
	If Instr(Klucz.Text, "C3-0")>0 Then
            Moc.Text="26,9W LED"
		Napiecie.Text="24VDC"
		Czest.Text="0 Hz"
	Oznacz3Klasa.EnableVisible = True
        End If
    Else
    End If
    
    
    If Instr(Klucz.Text, "INS242LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="69K"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "J1-3")>0 Then
            Moc.Text="12,6W LED"
        End If
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="17,2W LED"
        End If
        If Instr(Klucz.Text, "J2-3")>0 Then
            Moc.Text="25,2W LED"
		MaxTemp.Text="40"
        End If
	If Instr(Klucz.Text, "B2-1")>0 Then
            Moc.Text="34,4W LED"
        End If
	If Instr(Klucz.Text, "B2-3")>0 Then
            Moc.Text="50,3W LED"
	MaxTemp.Text="40"
        End If
	If Instr(Klucz.Text, "B4-1")>0 Then
            Moc.Text="68,8W LED"
	MaxTemp.Text="40"
        End If
	If Instr(Klucz.Text, "D1-0")>0 Then
            Moc.Text="10,6W LED"
		Napiecie.Text="24VDC"
		Czest.Text="0 Hz"
	Oznacz3Klasa.EnableVisible = True
        End If
		If Instr(Klucz.Text, "D2-0")>0 Then
            Moc.Text="20,2W LED"
		Napiecie.Text="24VDC"
		Czest.Text="0 Hz"
	Oznacz3Klasa.EnableVisible = True
        End If
		If Instr(Klucz.Text, "D4-0")>0 Then
            Moc.Text="39,4W LED"
		Napiecie.Text="24VDC"
		Czest.Text="0 Hz"
	Oznacz3Klasa.EnableVisible = True
        End If
    Else
    End If
    
     If Instr(Klucz.Text, "INL340LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-40"
    MaxTemp.Text="50"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "J2-1")>0 Then
            Moc.Text="18,0W LED"
        End If
    Else
    End If
    
    If Instr(Klucz.Text, "Protec")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-25"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
    KDBnumer.Text="OBAC19ATEX0356X"
    OznaczEX.Text="II 3G Ex ec IIC T4 Gc" & chr(13) & "II 3D Ex tc IIIC T70°C IP66/67 Dc"
    ObrazEx.EnableVisible = True
      If Instr(Klucz.Text, "A3")>0 Then
            Czest.Text="50/60 Hz"
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            MinTemp.Text="0"
            OznaczenieCNBOP1.Text="X"
            OznaczenieCNBOP2.Text="1"
            OznaczenieCNBOP3.Text="****EF"
            OznaczenieCNBOP4.Text="180"
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=True
            Battery.Text="Bat: NiMH 4Ah 4,8V"  
        End If
            If Instr(Klucz.Text, "ZB")>0 Then
            ObrazCNBOP.EnableVisible = False
            CNBOP.Text=""
            OznaczenieCNBOP1.Text="Z"
            OznaczenieCNBOP2.Text="0/1"
            OznaczenieCNBOP3.Text="****E*"
            OznaczenieCNBOP4.Text=""
            OznaczenieCNBOP1.EnableVisible=True
            OznaczenieCNBOP2.EnableVisible=True
            OznaczenieCNBOP3.EnableVisible=True
            OznaczenieCNBOP4.EnableVisible=False
            Battery.Text="" 
        End If
	If Instr(Klucz.Text, "S-1")>0 Then
            Moc.Text="18,2W LED"
        End If
        If Instr(Klucz.Text, "S-2")>0 Then
            Moc.Text="26,1W LED"
        End If
        If Instr(Klucz.Text, "M-1")>0 Then
            Moc.Text="35,4W LED"
        End If
        If Instr(Klucz.Text, "M-2")>0 Then
            Moc.Text="351,3W LED"
        End If
        If Instr(Klucz.Text, "L-1")>0 Then
            Moc.Text="44,0W LED"
        End If
         If Instr(Klucz.Text, "L-2")>0 Then
            Moc.Text="59,8W LED"
        End If   
    Else
    End If
    
     If Instr(Klucz.Text, "INS400LED")>0 Then
    ZnakCzest.Text = "f:"
    Oznacz2Klasa.EnableVisible = False
    Czest.Text="0/50/60 Hz"
    IP.Text="66/67"
    MinTemp.Text="-40"
    MaxTemp.Text="45"
    cosinusFi.Text="0,95"
    Napiecie.Text="230V"
        If Instr(Klucz.Text, "B4-0")>0 Then
            Moc.Text="48,9W LED"
        End If
        If Instr(Klucz.Text, "B4-2")>0 Then
            Moc.Text="67,8W LED"
        End If
    Else
    End If
    
    
    


    If Instr(Klucz.Text, "ZBS")>0 Then
            MinTemp.Text="-5"
        End If
         If Instr(Klucz.Text, "ZBC")>0 Or Instr(Klucz.Text, "ZBD")>0 Or Instr(Klucz.Text, "ZBM")>0  Or Instr(Klucz.Text, "ZBR")>0  Then
            MinTemp.Text="-20"
        End If
         If Instr(Klucz.Text, "ZBT")>0 Then
            MinTemp.Text="-0"
        End If
         If Instr(Klucz.Text, "ZBH")>0 Then
            MinTemp.Text="5"
        End If

 
    If Instr(Index.Text, "EXL210LED-06E206")>0 Then
            Moc.Text="19W LED"
            MinTemp.Text="-20"
            MaxTemp.Text="55"
            KDBnumer.Text=""
            OznaczEX.Text="II 3G Ex nA IIC T4 Gc" & chr(13) & "II 3D Ex tc IIIC T70°C IP66/67 Db"
        End If
        
    If Instr(Index.Text, "EXL210LED-12E411")>0 Then
            Moc.Text="38W LED"
            MinTemp.Text="-20"
            MaxTemp.Text="55"
            KDBnumer.Text=""
            OznaczEX.Text="II 3G Ex nA IIC T4 Gc" & chr(13) & "II 3D Ex tc IIIC T65°C IP66/67 Db"
        End If
        
        
         If Instr(Index.Text, "11E")>0 Then
          Oznacz3Klasa.EnableVisible=True
        End If
        
       
        IndexSN.Text = Index.Text


    

    'Wyswietlanie indeksu przed numerem seryjnym

    
    'wyswietlanie aktualnego roku
    currentYear = Year(Now())
    ROK.Text = currentYear
  
End Sub



Sub Printing

Main()
Dim sText As String
    sText = InputBox ("Wpisz numer zamówienia:","Dear User")
  If Instr(Klucz.Text, "A3")>0 Or Instr(Klucz.Text, "ZB")>0 Then
    Dim piktogram As Integer
        piktogram = MsgBox ("Czy zamówienie jest z piktogramem:",4,"Piktogram") 
         If piktogram=6 Then
             OznaczenieCNBOP3.Text = OznaczenieCNBOP3.Text + "G"
        End If
         If piktogram=7 Then
             OznaczenieCNBOP3.Text = OznaczenieCNBOP3.Text + "*"
        End If
    End If
If Instr(Klucz.Text, "EXL390LED")>0 Then
 Dim pradEXL390LED As String
    pradEXL390LED = InputBox ("Jaki prąd sterowania 250 czy 350:","Drogi uzytkowniku", 250)
        If Instr(Klucz.Text, "600-E4")>0 And pradEXL390LED = 250 Then
            Moc.Text="35,0W LED"
        End If
        If Instr(Klucz.Text, "600-E4")>0 And pradEXL390LED = 350 Then
            Moc.Text="47,0W LED"
        End If
        If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 250 Then
            Moc.Text="27,1W LED"
        End If
        If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 350 Then
            Moc.Text="36,6W LED"
        End If
        If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 250 Then
            Moc.Text="52,0W LED"
        End If
        If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 350 Then
            Moc.Text="70,5W LED"
        End If
End If
Dim Doc As Object
Dim PathDoc As String
Dim dataUtworzenia as date
Dim Props(0) As New com.sun.star.beans.PropertyValue
Props(0).Name  = "Hidden"  'the document will open hidden"
Props(0).Value = True
PathDoc = ConvertToURL("C:\Users\Montaż\Dropbox\Dokumentacja\atm\GENERATOR-METEK-TEST\DATABASE-LOGING.ods")
Doc = StarDesktop.loadComponentFromURL(PathDoc, "_default", 0, _
									Props()) 																
MySheet = Doc.Sheets.getByName("Metka-Srebrna")
'okno z licznikiem wiersza CelAH	
CelAH = MySheet.getCellRangeByName("AH1")

dataUtworzenia = now()

CelA = MySheet.getCellRangeByName("A" + CelAH.Value)
CelA.String = dataUtworzenia							
CelB = MySheet.getCellRangeByName("B" + CelAH.Value)
CelB.String = sText
CelC = MySheet.getCellRangeByName("C" + CelAH.Value)
CelC.String = Index.Text
CelD = MySheet.getCellRangeByName("D" + CelAH.Value)
CelD.String = Klucz.Text
CelE = MySheet.getCellRangeByName("E" + CelAH.Value)
CelE.String = ObrazCNBOP.EnableVisible
CelF = MySheet.getCellRangeByName("F" + CelAH.Value)
CelF.String = ObrazEX.EnableVisible
CelG = MySheet.getCellRangeByName("G" + CelAH.Value)
CelG.String = Czest.Text
CelH = MySheet.getCellRangeByName("H" + CelAH.Value)
CelH.String = IP.Text
CelI = MySheet.getCellRangeByName("I" + CelAH.Value)
CelI.String = CNBOP.Text
CelJ = MySheet.getCellRangeByName("J" + CelAH.Value)
CelJ.String = KDBnumer.Text
CelK = MySheet.getCellRangeByName("K" + CelAH.Value)
CelK.String = OznaczEX.Text
CelL = MySheet.getCellRangeByName("L" + CelAH.Value)
CelL.String = MinTemp.Text
CelM = MySheet.getCellRangeByName("M" + CelAH.Value)
CelM.String = MaxTemp.Text
CelN = MySheet.getCellRangeByName("N" + CelAH.Value)
CelN.String = cosinusFi.Text
CelO = MySheet.getCellRangeByName("O" + CelAH.Value)
CelO.String = Napiecie.Text
CelP = MySheet.getCellRangeByName("P" + CelAH.Value)
CelP.String = Moc.Text
CelQ = MySheet.getCellRangeByName("Q" + CelAH.Value)
CelQ.String = NumerCE.Text
CelR = MySheet.getCellRangeByName("R" + CelAH.Value)
CelR.String = OznaczenieCNBOP1.Text
CelS = MySheet.getCellRangeByName("S" + CelAH.Value)
CelS.String = OznaczenieCNBOP2.Text
CelT = MySheet.getCellRangeByName("T" + CelAH.Value)
CelT.String = OznaczenieCNBOP3.Text
CelU = MySheet.getCellRangeByName("U" + CelAH.Value)
CelU.String =OznaczenieCNBOP4.Text
CelW = MySheet.getCellRangeByName("W" + CelAH.Value)
CelW.String = Oznacz2Klasa.EnableVisible


'arkusz z numerami seryjnymi
'spradzanie czy dany indeks istnieje
Dim Flag as boolean
Dim WierszZIndeksem as Integer
Dim PrzechowanieNumeruSeryjnego as Integer
Dim CzyMniejszyNumer as Integer

SheetSerialNumbers = Doc.Sheets.getByName("Numery-Seryjne")
CelLicznikWierszy = SheetSerialNumbers.getCellRangeByName("M2")
Desc = SheetSerialNumbers.createSearchDescriptor()
Desc.SearchString = Index.Text
oFound = SheetSerialNumbers.findFirst(Desc)

if not IsNull(oFound) Then
WierszZIndeksem = oFound.CellAddress.Row + 1
CelSerialNumber = SheetSerialNumbers.getCellRangeByName("B" +  WierszZIndeksem)
PrzechowanieNumeruSeryjnego = CelSerialNumber.Value
end If

if IsNull(oFound) Then
CelDodajIndeks = SheetSerialNumbers.getCellRangeByName("A" + (CelLicznikWierszy.Value +1))
CelDodajIndeks.String = Index.Text
CelDodajDefaultNumeru = SheetSerialNumbers.getCellRangeByName("B" + (CelLicznikWierszy.Value - 1))
CelLicznikWierszy.Value = CelLicznikWierszy.Value + 1
CelSerialNumber = SheetSerialNumbers.getCellRangeByName("B" + CelLicznikWierszy.Value)
end If

'wyznaczenie poczatkowego numeru seryjnego
dim poczatkowyNumerSeryjny as Integer
dim KoncowyNumerSeryjny as Integer
poczatkowyNumerSeryjny = PrzechowanieNumeruSeryjnego + 1

'drukowanie metek, numery seryjne
Dim Counter As Integer
Counter = 0
Dim Iloscwydruku As Integer
    Iloscwydruku = InputBox ("Dla ilu opraw wydrukowac:","Dear User")
SerialNumber.Text= Format(PrzechowanieNumeruSeryjnego, "00000")

Dim Ilosckopii As Integer
    Ilosckopii = InputBox ("Ile kopii tego samego numeru seryjnego wydrukowac:","Dear User", 1)

CzyMniejszyNumer = InputBox ("Potwierdz początkowy numer seryjny:","Dear User", poczatkowyNumerSeryjny)
poczatkowyNumerSeryjny = CzyMniejszyNumer
CzyMniejszyNumer = CzyMniejszyNumer - 1

'walka z drukowaniem
Dim aPrintOps(0) AS NEW com.sun.star.beans.PropertyValue
aPrintOps(0).Name = "PaperFormat" 
aPrintOps(0).Value = com.sun.star.view.PaperFormat.USER

While Counter < Iloscwydruku
CzyMniejszyNumer = CzyMniejszyNumer +1
SerialNumber.Text=  Format(CzyMniejszyNumer, "00000")

PathDocMetka = ConvertToURL("C:\Users\Montaż\Dropbox\Dokumentacja\atm\GENERATOR-METEK-TEST\Generator-metek-100x50mm.odt")
printingDoc = StarDesktop.loadComponentFromURL(PathDocMetka, "_default", 0, _
									Props()) 	
printingDoc.Printer = aPrintOps()
for i = 1 to Ilosckopii
	printingDoc.print(aPrintOps())
	Wait 1000
	next i
Counter = Counter + 1
Wend
KoncowyNumerSeryjny = SerialNumber.Text

if CzyMniejszyNumer > CelSerialNumber.Value Then
CelSerialNumber.Value = CzyMniejszyNumer
end If

'dodanie numerow od do seryjnych
CelX = MySheet.getCellRangeByName("X" + CelAH.Value)
CelX.String = poczatkowyNumerSeryjny
CelY = MySheet.getCellRangeByName("Y" + CelAH.Value)
CelY.String = KoncowyNumerSeryjny


'okno z licznikiem wiersza CelAH
CelAH.Value = CelAH.Value + 1
Doc.Store	
  
End Sub



Sub Printing100x70
Main()
Dim sText As String
    sText = InputBox ("Wpisz numer zamówienia:","Dear User")
  If Instr(Klucz.Text, "A3")>0 Or Instr(Klucz.Text, "ZB")>0 Then
    Dim piktogram As Integer
        piktogram = MsgBox ("Czy zamówienie jest z piktogramem:",4,"Piktogram") 
         If piktogram=6 Then
             OznaczenieCNBOP3.Text = OznaczenieCNBOP3.Text + "G"
        End If
         If piktogram=7 Then
             OznaczenieCNBOP3.Text = OznaczenieCNBOP3.Text + "*"
        End If
    End If
If Instr(Klucz.Text, "EXL390LED")>0 Then
 Dim pradEXL390LED As String
    pradEXL390LED = InputBox ("Jaki prąd sterowania 250mA czy 350mA:","Drogi uzytkowniku", 250)
        If Instr(Klucz.Text, "0600-E4")>0 And pradEXL390LED = 250 Then
            Moc.Text="35,0W LED"
        End If
        If Instr(Klucz.Text, "0600-E4")>0 And pradEXL390LED = 350 Then
            Moc.Text="47,0W LED"
        End If
        If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 250 Then
            Moc.Text="27,1W LED"
        End If
        If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 350 Then
            Moc.Text="36,6W LED"
        End If
        If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 250 Then
            Moc.Text="52,0W LED"
        End If
        If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 350 Then
            Moc.Text="70,5W LED"
        End If
End If
Dim Doc As Object
Dim PathDoc As String
Dim dataUtworzenia as date
Dim Props(0) As New com.sun.star.beans.PropertyValue
Props(0).Name  = "Hidden"  'the document will open hidden"
Props(0).Value = True
PathDoc = ConvertToURL("C:\Users\Montaż\Dropbox\Dokumentacja\atm\GENERATOR-METEK-TEST\DATABASE-LOGING.ods")
Doc = StarDesktop.loadComponentFromURL(PathDoc, "_default", 0, _
									Props()) 																
MySheet = Doc.Sheets.getByName("Metka-Biala")
'okno z licznikiem wiersza CelAH	
CelAH = MySheet.getCellRangeByName("AH1")

dataUtworzenia = now()

CelA = MySheet.getCellRangeByName("A" + CelAH.Value)
CelA.String = dataUtworzenia							
CelB = MySheet.getCellRangeByName("B" + CelAH.Value)
CelB.String = sText
CelC = MySheet.getCellRangeByName("C" + CelAH.Value)
CelC.String = Index.Text
CelD = MySheet.getCellRangeByName("D" + CelAH.Value)
CelD.String = Klucz.Text
CelE = MySheet.getCellRangeByName("E" + CelAH.Value)
CelE.String = ObrazCNBOP.EnableVisible
CelF = MySheet.getCellRangeByName("F" + CelAH.Value)
CelF.String = ObrazEX.EnableVisible
CelG = MySheet.getCellRangeByName("G" + CelAH.Value)
CelG.String = Czest.Text
CelH = MySheet.getCellRangeByName("H" + CelAH.Value)
CelH.String = IP.Text
CelI = MySheet.getCellRangeByName("I" + CelAH.Value)
CelI.String = CNBOP.Text
CelJ = MySheet.getCellRangeByName("J" + CelAH.Value)
CelJ.String = KDBnumer.Text
CelK = MySheet.getCellRangeByName("K" + CelAH.Value)
CelK.String = OznaczEX.Text
CelL = MySheet.getCellRangeByName("L" + CelAH.Value)
CelL.String = MinTemp.Text
CelM = MySheet.getCellRangeByName("M" + CelAH.Value)
CelM.String = MaxTemp.Text
CelN = MySheet.getCellRangeByName("N" + CelAH.Value)
CelN.String = cosinusFi.Text
CelO = MySheet.getCellRangeByName("O" + CelAH.Value)
CelO.String = Napiecie.Text
CelP = MySheet.getCellRangeByName("P" + CelAH.Value)
CelP.String = Moc.Text
CelQ = MySheet.getCellRangeByName("Q" + CelAH.Value)
CelQ.String = NumerCE.Text
CelR = MySheet.getCellRangeByName("R" + CelAH.Value)
CelR.String = OznaczenieCNBOP1.Text
CelS = MySheet.getCellRangeByName("S" + CelAH.Value)
CelS.String = OznaczenieCNBOP2.Text
CelT = MySheet.getCellRangeByName("T" + CelAH.Value)
CelT.String = OznaczenieCNBOP3.Text
CelU = MySheet.getCellRangeByName("U" + CelAH.Value)
CelU.String =OznaczenieCNBOP4.Text
CelW = MySheet.getCellRangeByName("W" + CelAH.Value)
CelW.String = Oznacz2Klasa.EnableVisible


'arkusz z numerami seryjnymi
'spradzanie czy dany indeks istnieje
Dim Flag as boolean
Dim WierszZIndeksem as Integer
Dim PrzechowanieNumeruSeryjnego as Integer
Dim CzyMniejszyNumer as Integer

SheetSerialNumbers = Doc.Sheets.getByName("Numery-Seryjne")
CelLicznikWierszy = SheetSerialNumbers.getCellRangeByName("M2")
Desc = SheetSerialNumbers.createSearchDescriptor()
Desc.SearchString = Index.Text
oFound = SheetSerialNumbers.findFirst(Desc)

if not IsNull(oFound) Then
WierszZIndeksem = oFound.CellAddress.Row + 1
CelSerialNumber = SheetSerialNumbers.getCellRangeByName("C" +  WierszZIndeksem)
PrzechowanieNumeruSeryjnego = CelSerialNumber.Value
end If

if IsNull(oFound) Then
CelDodajIndeks = SheetSerialNumbers.getCellRangeByName("A" + (CelLicznikWierszy.Value +1))
CelDodajIndeks.String = Index.Text
CelDodajDefaultNumeru = SheetSerialNumbers.getCellRangeByName("C" + (CelLicznikWierszy.Value - 1))
CelLicznikWierszy.Value = CelLicznikWierszy.Value + 1
CelSerialNumber = SheetSerialNumbers.getCellRangeByName("C" + CelLicznikWierszy.Value)
end If

'wyznaczenie poczatkowego numeru seryjnego
dim poczatkowyNumerSeryjny as Integer
dim KoncowyNumerSeryjny as Integer
poczatkowyNumerSeryjny = PrzechowanieNumeruSeryjnego + 1

'drukowanie metek, numery seryjne
Dim Counter As Integer
Counter = 0
Dim Iloscwydruku As Integer
    Iloscwydruku = InputBox ("Dla ilu opraw wydrukowacc:","Dear User")
SerialNumber.Text= Format(PrzechowanieNumeruSeryjnego, "00000")

CzyMniejszyNumer = InputBox ("Potwierdz początkowy numer seryjny:","Dear User", poczatkowyNumerSeryjny)
poczatkowyNumerSeryjny = CzyMniejszyNumer
CzyMniejszyNumer = CzyMniejszyNumer - 1

'walka z drukowaniem
Dim aPrintOps(0) AS NEW com.sun.star.beans.PropertyValue
aPrintOps(0).Name = "PaperFormat" 
aPrintOps(0).Value = com.sun.star.view.PaperFormat.USER

While Counter < Iloscwydruku
CzyMniejszyNumer = CzyMniejszyNumer +1
SerialNumber.Text=  Format(CzyMniejszyNumer, "00000")

PathDocMetka = ConvertToURL("C:\Users\Montaż\Dropbox\Dokumentacja\atm\GENERATOR-METEK-TEST\Generator-metek-100x70mm.odt")
printingDoc = StarDesktop.loadComponentFromURL(PathDocMetka, "_default", 0, _
									Props()) 	
printingDoc.Printer = aPrintOps()
printingDoc.print(aPrintOps())
Wait 1000
Counter = Counter + 1
Wend
KoncowyNumerSeryjny = SerialNumber.Text


'dodanie numerow od do seryjnych
if CzyMniejszyNumer > CelSerialNumber.Value Then
CelSerialNumber.Value = CzyMniejszyNumer
end If

CelX = MySheet.getCellRangeByName("X" + CelAH.Value)
CelX.String = poczatkowyNumerSeryjny
CelY = MySheet.getCellRangeByName("Y" + CelAH.Value)
CelY.String = KoncowyNumerSeryjny


'okno z licznikiem wiersza CelAH
CelAH.Value = CelAH.Value + 1
Doc.Store	
  
  
  
End Sub


Sub Printing75x35

Main()
Dim sText As String
    sText = InputBox ("Wpisz numer zamówienia:","Dear User")
  If Instr(Klucz.Text, "A3")>0 Or Instr(Klucz.Text, "ZB")>0 Then
    Dim piktogram As Integer
        piktogram = MsgBox ("Czy zamówienie jest z piktogramem:",4,"Piktogram") 
         If piktogram=6 Then
             OznaczenieCNBOP3.Text = OznaczenieCNBOP3.Text + "G"
        End If
         If piktogram=7 Then
             OznaczenieCNBOP3.Text = OznaczenieCNBOP3.Text + "*"
        End If
    End If
If Instr(Klucz.Text, "EXL390LED")>0 Then
 Dim pradEXL390LED As String
    pradEXL390LED = InputBox ("Jaki prąd sterowania 250 czy 350:","Drogi uzytkowniku", 250)
        If Instr(Klucz.Text, "600-E4")>0 And pradEXL390LED = 250 Then
            Moc.Text="35,0W LED"
        End If
        If Instr(Klucz.Text, "600-E4")>0 And pradEXL390LED = 350 Then
            Moc.Text="47,0W LED"
        End If
        If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 250 Then
            Moc.Text="27,1W LED"
        End If
        If Instr(Klucz.Text, "1200-E3")>0 And pradEXL390LED = 350 Then
            Moc.Text="36,6W LED"
        End If
        If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 250 Then
            Moc.Text="52,0W LED"
        End If
        If Instr(Klucz.Text, "1200-E6")>0 And pradEXL390LED = 350 Then
            Moc.Text="70,5W LED"
        End If
End If
Dim Doc As Object
Dim PathDoc As String
Dim dataUtworzenia as date
Dim Props(0) As New com.sun.star.beans.PropertyValue
Props(0).Name  = "Hidden"  'the document will open hidden"
Props(0).Value = True
PathDoc = ConvertToURL("C:\Users\Montaż\Dropbox\Dokumentacja\atm\GENERATOR-METEK-TEST\DATABASE-LOGING.ods")
Doc = StarDesktop.loadComponentFromURL(PathDoc, "_default", 0, _
									Props()) 																
MySheet = Doc.Sheets.getByName("Metka-Srebrna")
'okno z licznikiem wiersza CelAH	
CelAH = MySheet.getCellRangeByName("AH1")

dataUtworzenia = now()

CelA = MySheet.getCellRangeByName("A" + CelAH.Value)
CelA.String = dataUtworzenia							
CelB = MySheet.getCellRangeByName("B" + CelAH.Value)
CelB.String = sText
CelC = MySheet.getCellRangeByName("C" + CelAH.Value)
CelC.String = Index.Text
CelD = MySheet.getCellRangeByName("D" + CelAH.Value)
CelD.String = Klucz.Text
CelE = MySheet.getCellRangeByName("E" + CelAH.Value)
CelE.String = ObrazCNBOP.EnableVisible
CelF = MySheet.getCellRangeByName("F" + CelAH.Value)
CelF.String = ObrazEX.EnableVisible
CelG = MySheet.getCellRangeByName("G" + CelAH.Value)
CelG.String = Czest.Text
CelH = MySheet.getCellRangeByName("H" + CelAH.Value)
CelH.String = IP.Text
CelI = MySheet.getCellRangeByName("I" + CelAH.Value)
CelI.String = CNBOP.Text
CelJ = MySheet.getCellRangeByName("J" + CelAH.Value)
CelJ.String = KDBnumer.Text
CelK = MySheet.getCellRangeByName("K" + CelAH.Value)
CelK.String = OznaczEX.Text
CelL = MySheet.getCellRangeByName("L" + CelAH.Value)
CelL.String = MinTemp.Text
CelM = MySheet.getCellRangeByName("M" + CelAH.Value)
CelM.String = MaxTemp.Text
CelN = MySheet.getCellRangeByName("N" + CelAH.Value)
CelN.String = cosinusFi.Text
CelO = MySheet.getCellRangeByName("O" + CelAH.Value)
CelO.String = Napiecie.Text
CelP = MySheet.getCellRangeByName("P" + CelAH.Value)
CelP.String = Moc.Text
CelQ = MySheet.getCellRangeByName("Q" + CelAH.Value)
CelQ.String = NumerCE.Text
CelR = MySheet.getCellRangeByName("R" + CelAH.Value)
CelR.String = OznaczenieCNBOP1.Text
CelS = MySheet.getCellRangeByName("S" + CelAH.Value)
CelS.String = OznaczenieCNBOP2.Text
CelT = MySheet.getCellRangeByName("T" + CelAH.Value)
CelT.String = OznaczenieCNBOP3.Text
CelU = MySheet.getCellRangeByName("U" + CelAH.Value)
CelU.String =OznaczenieCNBOP4.Text
CelW = MySheet.getCellRangeByName("W" + CelAH.Value)
CelW.String = Oznacz2Klasa.EnableVisible


'arkusz z numerami seryjnymi
'spradzanie czy dany indeks istnieje
Dim Flag as boolean
Dim WierszZIndeksem as Integer
Dim PrzechowanieNumeruSeryjnego as Integer
Dim CzyMniejszyNumer as Integer

SheetSerialNumbers = Doc.Sheets.getByName("Numery-Seryjne")
CelLicznikWierszy = SheetSerialNumbers.getCellRangeByName("M2")
Desc = SheetSerialNumbers.createSearchDescriptor()
Desc.SearchString = Index.Text
oFound = SheetSerialNumbers.findFirst(Desc)

if not IsNull(oFound) Then
WierszZIndeksem = oFound.CellAddress.Row + 1
CelSerialNumber = SheetSerialNumbers.getCellRangeByName("D" +  WierszZIndeksem)
PrzechowanieNumeruSeryjnego = CelSerialNumber.Value
end If

if IsNull(oFound) Then
CelDodajIndeks = SheetSerialNumbers.getCellRangeByName("A" + (CelLicznikWierszy.Value +1))
CelDodajIndeks.String = Index.Text
CelDodajDefaultNumeru = SheetSerialNumbers.getCellRangeByName("D" + (CelLicznikWierszy.Value - 1))
CelLicznikWierszy.Value = CelLicznikWierszy.Value + 1
CelSerialNumber = SheetSerialNumbers.getCellRangeByName("D" + CelLicznikWierszy.Value)
end If

'wyznaczenie poczatkowego numeru seryjnego
dim poczatkowyNumerSeryjny as Integer
dim KoncowyNumerSeryjny as Integer
poczatkowyNumerSeryjny = PrzechowanieNumeruSeryjnego + 1

'drukowanie metek, numery seryjne
Dim Counter As Integer
Counter = 0
Dim Iloscwydruku As Integer
    Iloscwydruku = InputBox ("Dla ilu opraw wydrukowac:","Dear User")
SerialNumber.Text= Format(PrzechowanieNumeruSeryjnego, "00000")

Dim Ilosckopii As Integer
    Ilosckopii = InputBox ("Ile kopii tego samego numeru seryjnego wydrukowac:","Dear User", 1)

CzyMniejszyNumer = InputBox ("Potwierdz początkowy numer seryjny:","Dear User", poczatkowyNumerSeryjny)
poczatkowyNumerSeryjny = CzyMniejszyNumer
CzyMniejszyNumer = CzyMniejszyNumer - 1

'walka z drukowaniem
Dim aPrintOps(0) AS NEW com.sun.star.beans.PropertyValue
aPrintOps(0).Name = "PaperFormat" 
aPrintOps(0).Value = com.sun.star.view.PaperFormat.USER

While Counter < Iloscwydruku
CzyMniejszyNumer = CzyMniejszyNumer +1
SerialNumber.Text=  Format(CzyMniejszyNumer, "00000")

PathDocMetka = ConvertToURL("C:\Users\Montaż\Dropbox\Dokumentacja\atm\GENERATOR-METEK-TEST\Generator-metek-75x35mm.odt")
printingDoc = StarDesktop.loadComponentFromURL(PathDocMetka, "_default", 0, _
									Props()) 	
printingDoc.Printer = aPrintOps()
for i = 1 to Ilosckopii
	printingDoc.print(aPrintOps())
	Wait 1000
	next i
Counter = Counter + 1
Wend
KoncowyNumerSeryjny = SerialNumber.Text

if CzyMniejszyNumer > CelSerialNumber.Value Then
CelSerialNumber.Value = CzyMniejszyNumer
end If

'dodanie numerow od do seryjnych
CelX = MySheet.getCellRangeByName("X" + CelAH.Value)
CelX.String = poczatkowyNumerSeryjny
CelY = MySheet.getCellRangeByName("Y" + CelAH.Value)
CelY.String = KoncowyNumerSeryjny


'okno z licznikiem wiersza CelAH
CelAH.Value = CelAH.Value + 1
Doc.Store	
  
End Sub

