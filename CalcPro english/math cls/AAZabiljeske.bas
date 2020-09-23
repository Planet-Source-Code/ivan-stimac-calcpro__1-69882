Attribute VB_Name = "AAZabiljeske"
'
'kod unosa imaginarne jedinice velikim slovom I umjesto i dolazi
'   do ludiranja
'
'
'napraviti not operator ili funkciju (funkcija je jednostavnija
'   varijanta
'
'
'kod provjere da li je operand -$ ili -_sys moguca pogreska ukoliko
'   imamo razmak pa -$
'
'
'betaInc ili betaCF nevalja (za neke parametre, mislim vece)
'
'
'u getSubRes dodan kod:
'        ElseIf InStrB(1, strSub, "\") <> 0 Then
'            lastErr = ERR_ExpectedExpres
'            lastErrNum = ERR_ExpectedExpresN
'        End If
'   kod obrade ako nema operatora u izrazu, ako doðe do glupiranja provjeriti
