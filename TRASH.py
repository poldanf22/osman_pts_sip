# iterasi 2 rata-rata - 2
            # rata" MTK ke MTK tambahan dan mapel MTK awal
            ws['CE{}'.format(
                r)] = "=IF($BR${}=0,$AS${},$AS${}-1)".format(r, r, r)
            ws['CE{}'.format(s)] = "=STDEV(CE2:CE{})".format(q)
            ws['CE{}'.format(t)] = "=MAX(CE2:CE{})".format(q)
            ws['CE{}'.format(u)] = "=MIN(CE2:CE{})".format(q)
            # rata" IND ke IND tambahan dan mapel IND awal
            ws['CF{}'.format(
                r)] = "=IF($BS${}=0,$AT${},$AT${}-1)".format(r, r, r)
            ws['CF{}'.format(s)] = "=STDEV(CF2:CF{})".format(q)
            ws['CF{}'.format(t)] = "=MAX(CF2:CF{})".format(q)
            ws['CF{}'.format(u)] = "=MIN(CF2:CF{})".format(q)
            # rata" ENG ke ENG tambahan dan mapel ENG awal
            ws['CG{}'.format(
                r)] = "=IF($BT${}=0,$AU${},$AU${}-1)".format(r, r, r)
            ws['CG{}'.format(s)] = "=STDEV(CG2:CG{})".format(q)
            ws['CG{}'.format(t)] = "=MAX(CG2:CG{})".format(q)
            ws['CG{}'.format(u)] = "=MIN(CG2:CG{})".format(q)
            # rata" SEJ ke SEJ tambahan dan mapel SEJ awal
            ws['CH{}'.format(
                r)] = "=IF($BU${}=0,$AV${},$AV${}-1)".format(r, r, r)
            ws['CH{}'.format(s)] = "=STDEV(CH2:CH{})".format(q)
            ws['CH{}'.format(t)] = "=MAX(CH2:CH{})".format(q)
            ws['CH{}'.format(u)] = "=MIN(CH2:CH{})".format(q)
            # rata" GEO ke GEO tambahan dan mapel GEO awal
            ws['CI{}'.format(
                r)] = "=IF($BV${}=0,$AW${},$AW${}-1)".format(r, r, r)
            ws['CI{}'.format(s)] = "=STDEV(CI2:CI{})".format(q)
            ws['CI{}'.format(t)] = "=MAX(CI2:CI{})".format(q)
            ws['CI{}'.format(u)] = "=MIN(CI2:CI{})".format(q)
            # rata" EKO ke EKO tambahan dan mapel EKO awal
            ws['CJ{}'.format(
                r)] = "=IF($BW${}=0,$AX${},$AX${}-1)".format(r, r, r)
            ws['CJ{}'.format(s)] = "=STDEV(CJ2:CJ{})".format(q)
            ws['CJ{}'.format(t)] = "=MAX(CJ2:CJ{})".format(q)
            ws['CJ{}'.format(u)] = "=MIN(CJ2:CJ{})".format(q)
            # rata" SOS ke SOS tambahan dan mapel SOS awal
            ws['CK{}'.format(
                r)] = "=IF($BX${}=0,$AY${},$AY${}-1)".format(r, r, r)
            ws['CK{}'.format(s)] = "=STDEV(CK2:CK{})".format(q)
            ws['CK{}'.format(t)] = "=MAX(CK2:CK{})".format(q)
            ws['CK{}'.format(u)] = "=MIN(CK2:CK{})".format(q)
            # jml MAPEL
            ws['CL{}'.format(r)] = "=ROUND(AVERAGE(CL2:CL{}),2)".format(q)
            ws['CL{}'.format(t)] = "=MAX(CL2:CL{})".format(q)
            ws['CL{}'.format(u)] = "=MIN(CL2:CL{})".format(q)
            # MAX Z SCORE
            ws['CM{}'.format(r)] = "=MAX(CM2:CM{})".format(q)
            ws['CN{}'.format(r)] = "=MAX(CN2:CN{})".format(q)
            ws['CO{}'.format(r)] = "=MAX(CO2:CO{})".format(q)
            ws['CP{}'.format(r)] = "=MAX(CP2:CP{})".format(q)
            ws['CQ{}'.format(r)] = "=MAX(CQ2:CQ{})".format(q)
            ws['CR{}'.format(r)] = "=MAX(CR2:CR{})".format(q)
            ws['CS{}'.format(r)] = "=MAX(CS2:CS{})".format(q)
            # NILAI STANDAR MTK
            ws['CT{}'.format(r)] = "=MAX(CT2:CT{})".format(q)
            ws['CT{}'.format(s)] = "=MIN(CT2:CT{})".format(q)
            ws['CT{}'.format(t)] = "=ROUND(AVERAGE(CT2:CT{}),2)".format(q)
            # NILAI STANDAR IND
            ws['CU{}'.format(r)] = "=MAX(CU2:CU{})".format(q)
            ws['CU{}'.format(s)] = "=MIN(CU2:CU{})".format(q)
            ws['CU{}'.format(t)] = "=ROUND(AVERAGE(CU2:CU{}),2)".format(q)
            # NILAI STANDAR ENG
            ws['CV{}'.format(r)] = "=MAX(CV2:CV{})".format(q)
            ws['CV{}'.format(s)] = "=MIN(CV2:CV{})".format(q)
            ws['CV{}'.format(t)] = "=ROUND(AVERAGE(CV2:CV{}),2)".format(q)
            # NILAI STANDAR SEJ
            ws['CW{}'.format(r)] = "=MAX(CW2:CW{})".format(q)
            ws['CW{}'.format(s)] = "=MIN(CW2:CW{})".format(q)
            ws['CW{}'.format(t)] = "=ROUND(AVERAGE(CW2:CW{}),2)".format(q)
            # NILAI STANDAR GEO
            ws['CX{}'.format(r)] = "=MAX(CX2:CX{})".format(q)
            ws['CX{}'.format(s)] = "=MIN(CX2:CX{})".format(q)
            ws['CX{}'.format(t)] = "=ROUND(AVERAGE(CX2:CX{}),2)".format(q)
            # NILAI STANDAR EKO
            ws['CY{}'.format(r)] = "=MAX(CY2:CY{})".format(q)
            ws['CY{}'.format(s)] = "=MIN(CY2:CY{})".format(q)
            ws['CY{}'.format(t)] = "=ROUND(AVERAGE(CY2:CY{}),2)".format(q)
            # NILAI STANDAR SOS
            ws['CZ{}'.format(r)] = "=MAX(CZ2:CZ{})".format(q)
            ws['CZ{}'.format(s)] = "=MIN(CZ2:CZ{})".format(q)
            ws['CZ{}'.format(t)] = "=ROUND(AVERAGE(CZ2:CZ{}),2)".format(q)
            # NILAI STANDAR S.JML
            ws['DA{}'.format(r)] = "=MAX(DA2:DA{})".format(q)
            ws['DA{}'.format(s)] = "=MIN(DA2:DA{})".format(q)
            ws['DA{}'.format(t)] = "=ROUND(AVERAGE(DA2:DA{}),2)".format(q)

            # TAMBAHAN
            ws['DD{}'.format(r)] = "=SUM(DD2:DD{})".format(q)
            ws['DE{}'.format(r)] = "=SUM(DE2:DE{})".format(q)
            ws['DF{}'.format(r)] = "=SUM(DF2:DF{})".format(q)
            ws['DG{}'.format(r)] = "=SUM(DG2:DG{})".format(q)
            ws['DH{}'.format(r)] = "=SUM(DH2:DH{})".format(q)
            ws['DI{}'.format(r)] = "=SUM(DI2:DI{})".format(q)
            ws['DJ{}'.format(r)] = "=SUM(DJ2:DJ{})".format(q)

            # iterasi 3 rata-rata - 3
            # rata" MTK ke MTK tambahan dan mapel MTK awal
            ws['DQ{}'.format(
                r)] = "=IF($DD${}=0,$CE${},$CE${}-1)".format(r, r, r)
            ws['DQ{}'.format(s)] = "=STDEV(DQ2:DQ{})".format(q)
            ws['DQ{}'.format(t)] = "=MAX(DQ2:DQ{})".format(q)
            ws['DQ{}'.format(u)] = "=MIN(DQ2:DQ{})".format(q)
            # rata" IND ke IND tambahan dan mapel IND awal
            ws['DR{}'.format(
                r)] = "=IF($DE${}=0,$CF${},$CF${}-1)".format(r, r, r)
            ws['DR{}'.format(s)] = "=STDEV(DR2:DR{})".format(q)
            ws['DR{}'.format(t)] = "=MAX(DR2:DR{})".format(q)
            ws['DR{}'.format(u)] = "=MIN(DR2:DR{})".format(q)
            # rata" ENG ke ENG tambahan dan mapel ENG awal
            ws['DS{}'.format(
                r)] = "=IF($DF${}=0,$CG${},$CG{}-1)".format(r, r, r)
            ws['DS{}'.format(s)] = "=STDEV(DS2:DS{})".format(q)
            ws['DS{}'.format(t)] = "=MAX(DS2:DS{})".format(q)
            ws['DS{}'.format(u)] = "=MIN(DS2:DS{})".format(q)
            # rata" SEJ ke SEJ tambahan dan mapel SEJ awal
            ws['DT{}'.format(
                r)] = "=IF($DG${}=0,$CH${},$CH${}-1)".format(r, r, r)
            ws['DT{}'.format(s)] = "=STDEV(DT2:DT{})".format(q)
            ws['DT{}'.format(t)] = "=MAX(DT2:DT{})".format(q)
            ws['DT{}'.format(u)] = "=MIN(DT2:DT{})".format(q)
            # rata" GEO ke GEO tambahan dan mapel GEO awal
            ws['DU{}'.format(
                r)] = "=IF($DH${}=0,$CI${},$CI${}-1)".format(r, r, r)
            ws['DU{}'.format(s)] = "=STDEV(DU2:DU{})".format(q)
            ws['DU{}'.format(t)] = "=MAX(DU2:DU{})".format(q)
            ws['DU{}'.format(u)] = "=MIN(DU2:DU{})".format(q)
            # rata" EKO ke EKO tambahan dan mapel EKO awal
            ws['DV{}'.format(
                r)] = "=IF($DI${}=0,$CJ${},$CJ${}-1)".format(r, r, r)
            ws['DV{}'.format(s)] = "=STDEV(DV2:DV{})".format(q)
            ws['DV{}'.format(t)] = "=MAX(DV2:DV{})".format(q)
            ws['DV{}'.format(u)] = "=MIN(DV2:DV{})".format(q)
            # rata" SOS ke SOS tambahan dan mapel SOS awal
            ws['DW{}'.format(
                r)] = "=IF($DJ${}=0,$CK${},$CK${}-1)".format(r, r, r)
            ws['DW{}'.format(s)] = "=STDEV(DW2:DW{})".format(q)
            ws['DW{}'.format(t)] = "=MAX(DW2:DW{})".format(q)
            ws['DW{}'.format(u)] = "=MIN(DW2:DW{})".format(q)
            # jml MAPEL
            ws['DX{}'.format(r)] = "=ROUND(AVERAGE(DX2:DX{}),2)".format(q)
            ws['DX{}'.format(t)] = "=MAX(DX2:DX{})".format(q)
            ws['DX{}'.format(u)] = "=MIN(DX2:DX{})".format(q)
            # MAX Z SCORE
            ws['DY{}'.format(r)] = "=MAX(DY2:DY{})".format(q)
            ws['DZ{}'.format(r)] = "=MAX(DZ2:DZ{})".format(q)
            ws['EA{}'.format(r)] = "=MAX(EA2:EA{})".format(q)
            ws['EB{}'.format(r)] = "=MAX(EB2:EB{})".format(q)
            ws['EC{}'.format(r)] = "=MAX(EC2:EC{})".format(q)
            ws['ED{}'.format(r)] = "=MAX(ED2:ED{})".format(q)
            ws['EE{}'.format(r)] = "=MAX(EE2:EE{})".format(q)
            # NILAI STANDAR MTK
            ws['EF{}'.format(r)] = "=MAX(EF2:EF{})".format(q)
            ws['EF{}'.format(s)] = "=MIN(EF2:EF{})".format(q)
            ws['EF{}'.format(t)] = "=ROUND(AVERAGE(EF2:EF{}),2)".format(q)
            # NILAI STANDAR IND
            ws['EG{}'.format(r)] = "=MAX(EG2:EG{})".format(q)
            ws['EG{}'.format(s)] = "=MIN(EG2:EG{})".format(q)
            ws['EG{}'.format(t)] = "=ROUND(AVERAGE(EG2:EG{}),2)".format(q)
            # NILAI STANDAR ENG
            ws['EH{}'.format(r)] = "=MAX(EH2:EH{})".format(q)
            ws['EH{}'.format(s)] = "=MIN(EH2:EH{})".format(q)
            ws['EH{}'.format(t)] = "=ROUND(AVERAGE(EH2:EH{}),2)".format(q)
            # NILAI STANDAR SEJ
            ws['EI{}'.format(r)] = "=MAX(EI2:EI{})".format(q)
            ws['EI{}'.format(s)] = "=MIN(EI2:EI{})".format(q)
            ws['EI{}'.format(t)] = "=ROUND(AVERAGE(EI2:EI{}),2)".format(q)
            # NILAI STANDAR GEO
            ws['EJ{}'.format(r)] = "=MAX(EJ2:EJ{})".format(q)
            ws['EJ{}'.format(s)] = "=MIN(EJ2:EJ{})".format(q)
            ws['EJ{}'.format(t)] = "=ROUND(AVERAGE(EJ2:EJ{}),2)".format(q)
            # NILAI STANDAR EKO
            ws['EK{}'.format(r)] = "=MAX(EK2:EK{})".format(q)
            ws['EK{}'.format(s)] = "=MIN(EK2:EK{})".format(q)
            ws['EK{}'.format(t)] = "=ROUND(AVERAGE(EK2:EK{}),2)".format(q)
            # NILAI STANDAR SOS
            ws['EL{}'.format(r)] = "=MAX(EL2:EL{})".format(q)
            ws['EL{}'.format(s)] = "=MIN(EL2:EL{})".format(q)
            ws['EL{}'.format(t)] = "=ROUND(AVERAGE(EL2:EL{}),2)".format(q)
            # NILAI STANDAR S.JML
            ws['EM{}'.format(r)] = "=MAX(EM2:EM{})".format(q)
            ws['EM{}'.format(s)] = "=MIN(EM2:EM{})".format(q)
            ws['EM{}'.format(t)] = "=ROUND(AVERAGE(EM2:EM{}),2)".format(q)

            # TAMBAHAN
            ws['EP{}'.format(r)] = "=SUM(EP2:EP{})".format(q)
            ws['EQ{}'.format(r)] = "=SUM(EQ2:EQ{})".format(q)
            ws['ER{}'.format(r)] = "=SUM(ER2:ER{})".format(q)
            ws['ES{}'.format(r)] = "=SUM(ES2:ES{})".format(q)
            ws['ET{}'.format(r)] = "=SUM(ET2:ET{})".format(q)
            ws['EU{}'.format(r)] = "=SUM(EU2:EU{})".format(q)
            ws['EV{}'.format(r)] = "=SUM(EV2:EV{})".format(q)

            # iterasi 4 rata-rata - 4
            # rata" MTK ke MTK tambahan dan mapel MTK awal
            ws['FC{}'.format(
                r)] = "=IF($EP${}=0,$DQ${},$DQ${}-1)".format(r, r, r)
            ws['FC{}'.format(s)] = "=STDEV(FC2:FC{})".format(q)
            ws['FC{}'.format(t)] = "=MAX(FC2:FC{})".format(q)
            ws['FC{}'.format(u)] = "=MIN(FC2:FC{})".format(q)
            # rata" IND ke IND tambahan dan mapel IND awal
            ws['FD{}'.format(
                r)] = "=IF($EQ${}=0,$DR${},$DR${}-1)".format(r, r, r)
            ws['FD{}'.format(s)] = "=STDEV(FD2:FD{})".format(q)
            ws['FD{}'.format(t)] = "=MAX(FD2:FD{})".format(q)
            ws['FD{}'.format(u)] = "=MIN(FD2:FD{})".format(q)
            # rata" ENG ke ENG tambahan dan mapel ENG awal
            ws['FE{}'.format(
                r)] = "=IF($ER${}=0,$DS${},$DS{}-1)".format(r, r, r)
            ws['FE{}'.format(s)] = "=STDEV(FE2:FE{})".format(q)
            ws['FE{}'.format(t)] = "=MAX(FE2:FE{})".format(q)
            ws['FE{}'.format(u)] = "=MIN(FE2:FE{})".format(q)
            # rata" SEJ ke SEJ tambahan dan mapel SEJ awal
            ws['FF{}'.format(
                r)] = "=IF($ES${}=0,$DT${},$DT${}-1)".format(r, r, r)
            ws['FF{}'.format(s)] = "=STDEV(FF2:FF{})".format(q)
            ws['FF{}'.format(t)] = "=MAX(FF2:FF{})".format(q)
            ws['FF{}'.format(u)] = "=MIN(FF2:FF{})".format(q)
            # rata" GEO ke GEO tambahan dan mapel GEO awal
            ws['FG{}'.format(
                r)] = "=IF($ET${}=0,$DU${},$DU${}-1)".format(r, r, r)
            ws['FG{}'.format(s)] = "=STDEV(FG2:FG{})".format(q)
            ws['FG{}'.format(t)] = "=MAX(FG2:FG{})".format(q)
            ws['FG{}'.format(u)] = "=MIN(FG2:FG{})".format(q)
            # rata" EKO ke EKO tambahan dan mapel EKO awal
            ws['FH{}'.format(
                r)] = "=IF($EU${}=0,$DV${},$DV${}-1)".format(r, r, r)
            ws['FH{}'.format(s)] = "=STDEV(FH2:FH{})".format(q)
            ws['FH{}'.format(t)] = "=MAX(FH2:FH{})".format(q)
            ws['FH{}'.format(u)] = "=MIN(FH2:FH{})".format(q)
            # rata" SOS ke SOS tambahan dan mapel SOS awal
            ws['FI{}'.format(
                r)] = "=IF($EV${}=0,$DW${},$DW${}-1)".format(r, r, r)
            ws['FI{}'.format(s)] = "=STDEV(FI2:FI{})".format(q)
            ws['FI{}'.format(t)] = "=MAX(FI2:FI{})".format(q)
            ws['FI{}'.format(u)] = "=MIN(FI2:FI{})".format(q)
            # jml MAPEL
            ws['FJ{}'.format(r)] = "=ROUND(AVERAGE(FJ2:FJ{}),2)".format(q)
            ws['FJ{}'.format(t)] = "=MAX(FJ2:FJ{})".format(q)
            ws['FJ{}'.format(u)] = "=MIN(FJ2:FJ{})".format(q)
            # MAX Z SCORE
            ws['FK{}'.format(r)] = "=MAX(FK2:FK{})".format(q)
            ws['FL{}'.format(r)] = "=MAX(FL2:FL{})".format(q)
            ws['FM{}'.format(r)] = "=MAX(FM2:FM{})".format(q)
            ws['FN{}'.format(r)] = "=MAX(FN2:FN{})".format(q)
            ws['FO{}'.format(r)] = "=MAX(FO2:FO{})".format(q)
            ws['FP{}'.format(r)] = "=MAX(FP2:FP{})".format(q)
            ws['FQ{}'.format(r)] = "=MAX(FQ2:FQ{})".format(q)
            # NILAI STANDAR MTK
            ws['FR{}'.format(r)] = "=MAX(FR2:FR{})".format(q)
            ws['FR{}'.format(s)] = "=MIN(FR2:FR{})".format(q)
            ws['FR{}'.format(t)] = "=ROUND(AVERAGE(FR2:FR{}),2)".format(q)
            # NILAI STANDAR IND
            ws['FS{}'.format(r)] = "=MAX(FS2:FS{})".format(q)
            ws['FS{}'.format(s)] = "=MIN(FS2:FS{})".format(q)
            ws['FS{}'.format(t)] = "=ROUND(AVERAGE(FS2:FS{}),2)".format(q)
            # NILAI STANDAR ENG
            ws['FT{}'.format(r)] = "=MAX(FT2:FT{})".format(q)
            ws['FT{}'.format(s)] = "=MIN(FT2:FT{})".format(q)
            ws['FT{}'.format(t)] = "=ROUND(AVERAGE(FT2:FT{}),2)".format(q)
            # NILAI STANDAR SEJ
            ws['FU{}'.format(r)] = "=MAX(FU2:FU{})".format(q)
            ws['FU{}'.format(s)] = "=MIN(FU2:FU{})".format(q)
            ws['FU{}'.format(t)] = "=ROUND(AVERAGE(FU2:FU{}),2)".format(q)
            # NILAI STANDAR GEO
            ws['FV{}'.format(r)] = "=MAX(FV2:FV{})".format(q)
            ws['FV{}'.format(s)] = "=MIN(FV2:FV{})".format(q)
            ws['FV{}'.format(t)] = "=ROUND(AVERAGE(FV2:FV{}),2)".format(q)
            # NILAI STANDAR EKO
            ws['FW{}'.format(r)] = "=MAX(FW2:FW{})".format(q)
            ws['FW{}'.format(s)] = "=MIN(FW2:FW{})".format(q)
            ws['FW{}'.format(t)] = "=ROUND(AVERAGE(FW2:FW{}),2)".format(q)
            # NILAI STANDAR SOS
            ws['FX{}'.format(r)] = "=MAX(FX2:FX{})".format(q)
            ws['FX{}'.format(s)] = "=MIN(FX2:FX{})".format(q)
            ws['FX{}'.format(t)] = "=ROUND(AVERAGE(FX2:FX{}),2)".format(q)
            # NILAI STANDAR S.JML
            ws['FY{}'.format(r)] = "=MAX(FY2:FY{})".format(q)
            ws['FY{}'.format(s)] = "=MIN(FY2:FY{})".format(q)
            ws['FY{}'.format(t)] = "=ROUND(AVERAGE(FY2:FY{}),2)".format(q)

            # TAMBAHAN
            ws['GB{}'.format(r)] = "=SUM(GB2:GB{})".format(q)
            ws['GC{}'.format(r)] = "=SUM(GC2:GC{})".format(q)
            ws['GD{}'.format(r)] = "=SUM(GD2:GD{})".format(q)
            ws['GE{}'.format(r)] = "=SUM(GE2:GE{})".format(q)
            ws['GF{}'.format(r)] = "=SUM(GF2:GF{})".format(q)
            ws['GG{}'.format(r)] = "=SUM(GG2:GG{})".format(q)
            ws['GH{}'.format(r)] = "=SUM(GH2:GH{})".format(q)

            # iterasi 5 rata-rata - 5
            # rata" MTK ke MTK tambahan dan mapel MTK awal
            ws['GO{}'.format(
                r)] = "=IF($GB${}=0,$FC${},$FC${}-1)".format(r, r, r)
            ws['GO{}'.format(s)] = "=STDEV(GO2:GO{})".format(q)
            ws['GO{}'.format(t)] = "=MAX(GO2:GO{})".format(q)
            ws['GO{}'.format(u)] = "=MIN(GO2:GO{})".format(q)
            # rata" IND ke IND tambahan dan mapel IND awal
            ws['GP{}'.format(
                r)] = "=IF($GC${}=0,$FD${},$FD${}-1)".format(r, r, r)
            ws['GP{}'.format(s)] = "=STDEV(GP2:GP{})".format(q)
            ws['GP{}'.format(t)] = "=MAX(GP2:GP{})".format(q)
            ws['GP{}'.format(u)] = "=MIN(GP2:GP{})".format(q)
            # rata" ENG ke ENG tambahan dan mapel ENG awal
            ws['GQ{}'.format(
                r)] = "=IF($GD${}=0,$FE${},$FE{}-1)".format(r, r, r)
            ws['GQ{}'.format(s)] = "=STDEV(GQ2:GQ{})".format(q)
            ws['GQ{}'.format(t)] = "=MAX(GQ2:GQ{})".format(q)
            ws['GQ{}'.format(u)] = "=MIN(GQ2:GQ{})".format(q)
            # rata" SEJ ke SEJ tambahan dan mapel SEJ awal
            ws['GR{}'.format(
                r)] = "=IF($GE${}=0,$FF${},$FF${}-1)".format(r, r, r)
            ws['GR{}'.format(s)] = "=STDEV(GR2:GR{})".format(q)
            ws['GR{}'.format(t)] = "=MAX(GR2:GR{})".format(q)
            ws['GR{}'.format(u)] = "=MIN(GR2:GR{})".format(q)
            # rata" GEO ke GEO tambahan dan mapel GEO awal
            ws['GS{}'.format(
                r)] = "=IF($GF${}=0,$FG${},$FG${}-1)".format(r, r, r)
            ws['GS{}'.format(s)] = "=STDEV(GS2:GS{})".format(q)
            ws['GS{}'.format(t)] = "=MAX(GS2:GS{})".format(q)
            ws['GS{}'.format(u)] = "=MIN(GS2:GS{})".format(q)
            # rata" EKO ke EKO tambahan dan mapel EKO awal
            ws['GT{}'.format(
                r)] = "=IF($GG${}=0,$FH${},$FH${}-1)".format(r, r, r)
            ws['GT{}'.format(s)] = "=STDEV(GT2:GT{})".format(q)
            ws['GT{}'.format(t)] = "=MAX(GT2:GT{})".format(q)
            ws['GT{}'.format(u)] = "=MIN(GT2:GT{})".format(q)
            # rata" SOS ke SOS tambahan dan mapel SOS awal
            ws['GU{}'.format(
                r)] = "=IF($GH${}=0,$FI${},$FI${}-1)".format(r, r, r)
            ws['GU{}'.format(s)] = "=STDEV(GU2:GU{})".format(q)
            ws['GU{}'.format(t)] = "=MAX(GU2:GU{})".format(q)
            ws['GU{}'.format(u)] = "=MIN(GU2:GU{})".format(q)
            # jml MAPEL
            ws['GV{}'.format(r)] = "=ROUND(AVERAGE(GV2:GV{}),2)".format(q)
            ws['GV{}'.format(t)] = "=MAX(GV2:GV{})".format(q)
            ws['GV{}'.format(u)] = "=MIN(GV2:GV{})".format(q)
            # MAX Z SCORE
            ws['GW{}'.format(r)] = "=MAX(GW2:GW{})".format(q)
            ws['GX{}'.format(r)] = "=MAX(GX2:GX{})".format(q)
            ws['GY{}'.format(r)] = "=MAX(GY2:GY{})".format(q)
            ws['GZ{}'.format(r)] = "=MAX(GZ2:GZ{})".format(q)
            ws['HA{}'.format(r)] = "=MAX(HA2:HA{})".format(q)
            ws['HB{}'.format(r)] = "=MAX(HB2:HB{})".format(q)
            ws['HC{}'.format(r)] = "=MAX(HC2:HC{})".format(q)
            # NILAI STANDAR MTK
            ws['HD{}'.format(r)] = "=MAX(HD2:HD{})".format(q)
            ws['HD{}'.format(s)] = "=MIN(HD2:HD{})".format(q)
            ws['HD{}'.format(t)] = "=ROUND(AVERAGE(HD2:HD{}),2)".format(q)
            # NILAI STANDAR IND
            ws['HE{}'.format(r)] = "=MAX(HE2:HE{})".format(q)
            ws['HE{}'.format(s)] = "=MIN(HE2:HE{})".format(q)
            ws['HE{}'.format(t)] = "=ROUND(AVERAGE(HE2:HE{}),2)".format(q)
            # NILAI STANDAR ENG
            ws['HF{}'.format(r)] = "=MAX(HF2:HF{})".format(q)
            ws['HF{}'.format(s)] = "=MIN(HF2:HF{})".format(q)
            ws['HF{}'.format(t)] = "=ROUND(AVERAGE(HF2:HF{}),2)".format(q)
            # NILAI STANDAR SEJ
            ws['HG{}'.format(r)] = "=MAX(HG2:HG{})".format(q)
            ws['HG{}'.format(s)] = "=MIN(HG2:HG{})".format(q)
            ws['HG{}'.format(t)] = "=ROUND(AVERAGE(HG2:HG{}),2)".format(q)
            # NILAI STANDAR GEO
            ws['HH{}'.format(r)] = "=MAX(HH2:HH{})".format(q)
            ws['HH{}'.format(s)] = "=MIN(HH2:HH{})".format(q)
            ws['HH{}'.format(t)] = "=ROUND(AVERAGE(HH2:HH{}),2)".format(q)
            # NILAI STANDAR EKO
            ws['HI{}'.format(r)] = "=MAX(HI2:HI{})".format(q)
            ws['HI{}'.format(s)] = "=MIN(HI2:HI{})".format(q)
            ws['HI{}'.format(t)] = "=ROUND(AVERAGE(HI2:HI{}),2)".format(q)
            # NILAI STANDAR SOS
            ws['HJ{}'.format(r)] = "=MAX(HJ2:HJ{})".format(q)
            ws['HJ{}'.format(s)] = "=MIN(HJ2:HJ{})".format(q)
            ws['HJ{}'.format(t)] = "=ROUND(AVERAGE(HJ2:HJ{}),2)".format(q)
            # NILAI STANDAR S.JML
            ws['HK{}'.format(r)] = "=MAX(HK2:HK{})".format(q)
            ws['HK{}'.format(s)] = "=MIN(HK2:HK{})".format(q)
            ws['HK{}'.format(t)] = "=ROUND(AVERAGE(HK2:HK{}),2)".format(q)

            # TAMBAHAN
            ws['HN{}'.format(r)] = "=SUM(HN2:HN{})".format(q)
            ws['HO{}'.format(r)] = "=SUM(HO2:HO{})".format(q)
            ws['HP{}'.format(r)] = "=SUM(HP2:HP{})".format(q)
            ws['HQ{}'.format(r)] = "=SUM(HQ2:HQ{})".format(q)
            ws['HR{}'.format(r)] = "=SUM(HR2:HR{})".format(q)
            ws['HS{}'.format(r)] = "=SUM(HS2:HS{})".format(q)
            ws['HT{}'.format(r)] = "=SUM(HT2:HT{})".format(q)