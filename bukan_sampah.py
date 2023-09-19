
    if selected_file == "Nilai Std. 8 SMP (KM-MTK SB)":
        # menghilangkan hamburger
        st.markdown("""
        <style>
        .css-1rs6os.edgvbvh3
        {
            visibility:hidden;
        }
        .css-1lsmgbg.egzxvld0
        {
            visibility:hidden;
        }
        </style>
        """, unsafe_allow_html=True)

        image = Image.open('logo resmi nf resize.png')
        st.image(image)

        st.title("Olah Nilai Standar KM")

        st.header("8 SMP-MTK SB")

        col6 = st.container()

        with col6:
            KELAS = st.selectbox(
                "KELAS",
                ("--Pilih Kelas--", "8 SMP SB"))

        col7 = st.container()

        with col7:
            SEMESTER = st.selectbox(
                "SEMESTER",
                ("--Pilih Semester--", "SEMESTER 1", "SEMESTER 2"))

        col8 = st.container()

        with col8:
            PENILAIAN = st.selectbox(
                "PENILAIAN",
                ("--Pilih Penilaian--", "SUMATIF TENGAH SEMESTER"))

        col9 = st.container()

        with col9:
            KURIKULUM = st.selectbox(
                "KURIKULUM",
                ("--Pilih Kurikulum--", "KM"))

        TAHUN = st.text_input("Masukkan Tahun Ajaran",
                              placeholder="contoh: 2022-2023")

        col1, col2, col3, col4, col5 = st.columns(5)

        with col1:
            MTK = st.selectbox(
                "JML. SOAL MAT. SB.",
                ("--Pilih--", 25, 30, 35, 40, 45))

        with col2:
            IND = st.selectbox(
                "JML. SOAL IND.",
                ("--Pilih--", 25, 30, 35, 40, 45))

        with col3:
            ENG = st.selectbox(
                "JML. SOAL ENG.",
                ("--Pilih--", 25, 30, 35, 40, 45))

        with col4:
            IPA = st.selectbox(
                "JML. SOAL IPA.",
                ("--Pilih--", 25, 30, 35, 40, 45))

        with col5:
            IPS = st.selectbox(
                "JML. SOAL IPS.",
                ("--Pilih--", 25, 30, 35, 40, 45))

        JML_SOAL_MAT = MTK
        JML_SOAL_IND = IND
        JML_SOAL_ENG = ENG
        JML_SOAL_IPA = IPA
        JML_SOAL_IPS = IPS

        uploaded_file = st.file_uploader(
            'Letakkan file excel', type='xlsx')

        if uploaded_file is not None:

            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb['Sheet1']

            q = len(ws['K'])
            r = len(ws['K'])+2
            s = len(ws['K'])+3
            t = len(ws['K'])+4
            u = len(ws['K'])+5
            v = len(ws['K'])+6
            w = len(ws['K'])+7
            x = len(ws['K'])+8

            ws['G{}'.format(r)] = "=ROUND(AVERAGE(G2:G{}),2)".format(q)  # mat
            ws['H{}'.format(r)] = "=ROUND(AVERAGE(H2:H{}),2)".format(q)  # ind
            ws['I{}'.format(r)] = "=ROUND(AVERAGE(I2:I{}),2)".format(q)  # eng
            ws['J{}'.format(r)] = "=ROUND(AVERAGE(J2:J{}),2)".format(q)  # ipa
            ws['K{}'.format(r)] = "=ROUND(AVERAGE(K2:K{}),2)".format(q)  # ips
            ws['L{}'.format(r)] = "=ROUND(AVERAGE(L2:L{}),2)".format(q)  # jml
            ws['G{}'.format(s)] = "=STDEV(G2:G{})".format(q)
            ws['H{}'.format(s)] = "=STDEV(H2:H{})".format(q)
            ws['I{}'.format(s)] = "=STDEV(I2:I{})".format(q)
            ws['J{}'.format(s)] = "=STDEV(J2:J{})".format(q)
            ws['K{}'.format(s)] = "=STDEV(K2:K{})".format(q)
            ws['G{}'.format(t)] = "=MAX(G2:G{})".format(q)
            ws['H{}'.format(t)] = "=MAX(H2:H{})".format(q)
            ws['I{}'.format(t)] = "=MAX(I2:I{})".format(q)
            ws['J{}'.format(t)] = "=MAX(J2:J{})".format(q)
            ws['K{}'.format(t)] = "=MAX(K2:K{})".format(q)
            ws['L{}'.format(t)] = "=MAX(L2:L{})".format(q)
            ws['M{}'.format(r)] = "=MAX(M2:M{})".format(q)
            ws['N{}'.format(r)] = "=MAX(N2:N{})".format(q)
            ws['O{}'.format(r)] = "=MAX(O2:O{})".format(q)
            ws['P{}'.format(r)] = "=MAX(P2:P{})".format(q)
            ws['Q{}'.format(r)] = "=MAX(Q2:Q{})".format(q)
            ws['R{}'.format(r)] = "=MAX(R2:R{})".format(q)
            ws['S{}'.format(r)] = "=MAX(S2:S{})".format(q)
            ws['T{}'.format(r)] = "=MAX(T2:T{})".format(q)
            ws['U{}'.format(r)] = "=MAX(U2:U{})".format(q)
            ws['V{}'.format(r)] = "=MAX(V2:V{})".format(q)
            ws['W{}'.format(r)] = "=ROUND(MAX(W2:W{}),2)".format(q)
            ws['G{}'.format(u)] = "=MIN(G2:G{})".format(q)
            ws['H{}'.format(u)] = "=MIN(H2:H{})".format(q)
            ws['I{}'.format(u)] = "=MIN(I2:I{})".format(q)
            ws['J{}'.format(u)] = "=MIN(J2:J{})".format(q)
            ws['K{}'.format(u)] = "=MIN(K2:K{})".format(q)
            ws['L{}'.format(u)] = "=MIN(L2:L{})".format(q)
            ws['R{}'.format(s)] = "=MIN(R2:R{})".format(q)
            ws['S{}'.format(s)] = "=MIN(S2:S{})".format(q)
            ws['T{}'.format(s)] = "=MIN(T2:T{})".format(q)
            ws['U{}'.format(s)] = "=MIN(U2:U{})".format(q)
            ws['V{}'.format(s)] = "=MIN(V2:V{})".format(q)
            ws['W{}'.format(s)] = "=MIN(W2:W{})".format(q)
            ws['R{}'.format(t)] = "=ROUND(AVERAGE(R2:R{}),2)".format(q)
            ws['S{}'.format(t)] = "=ROUND(AVERAGE(S2:S{}),2)".format(q)
            ws['T{}'.format(t)] = "=ROUND(AVERAGE(T2:T{}),2)".format(q)
            ws['U{}'.format(t)] = "=ROUND(AVERAGE(U2:U{}),2)".format(q)
            ws['V{}'.format(t)] = "=ROUND(AVERAGE(V2:V{}),2)".format(q)
            ws['W{}'.format(t)] = "=ROUND(AVERAGE(W2:W{}),2)".format(q)
            ws['X{}'.format(r)] = "=MAX(X2:X{})".format(q)
            ws['Z{}'.format(r)] = "=SUM(Z2:Z{})".format(q)
            ws['AA{}'.format(r)] = "=SUM(AA2:AA{})".format(q)
            ws['AB{}'.format(r)] = "=SUM(AB2:AB{})".format(q)
            ws['AC{}'.format(r)] = "=SUM(AC2:AC{})".format(q)
            ws['AD{}'.format(r)] = "=SUM(AD2:AD{})".format(q)
            # new
            # iterasi 1 rata-rata - 1
            ws['F{}'.format(v)] = 'JUMLAH SOAL'
            ws['G{}'.format(v)] = JML_SOAL_MAT
            ws['H{}'.format(v)] = JML_SOAL_IND
            ws['I{}'.format(v)] = JML_SOAL_ENG
            ws['J{}'.format(v)] = JML_SOAL_IPA
            ws['K{}'.format(v)] = JML_SOAL_IPS
            ws['AK{}'.format(r)] = "=IF($Z${}=0,$G${},$G${}-1)".format(r, r, r)
            ws['AK{}'.format(s)] = "=STDEV(AK2:AK{})".format(q)
            ws['AK{}'.format(t)] = "=MAX(AK2:AK{})".format(q)
            ws['AK{}'.format(u)] = "=MIN(AK2:AK{})".format(q)
            ws['AL{}'.format(
                r)] = "=IF($AA${}=0,$H${},$H${}-1)".format(r, r, r)
            ws['AL{}'.format(s)] = "=STDEV(AL2:AL{})".format(q)
            ws['AL{}'.format(t)] = "=MAX(AL2:AL{})".format(q)
            ws['AL{}'.format(u)] = "=MIN(AL2:AL{})".format(q)
            ws['AM{}'.format(
                r)] = "=IF($AB${}=0,$I${},$I${}-1)".format(r, r, r)
            ws['AM{}'.format(s)] = "=STDEV(AM2:AM{})".format(q)
            ws['AM{}'.format(t)] = "=MAX(AM2:AM{})".format(q)
            ws['AM{}'.format(u)] = "=MIN(AM2:AM{})".format(q)
            ws['AN{}'.format(
                r)] = "=IF($AC${}=0,$J${},$J${}-1)".format(r, r, r)
            ws['AN{}'.format(s)] = "=STDEV(AN2:AN{})".format(q)
            ws['AN{}'.format(t)] = "=MAX(AN2:AN{})".format(q)
            ws['AN{}'.format(u)] = "=MIN(AN2:AN{})".format(q)
            ws['AO{}'.format(
                r)] = "=IF($AD${}=0,$K${},$K${}-1)".format(r, r, r)
            ws['AO{}'.format(s)] = "=STDEV(AO2:AO{})".format(q)
            ws['AO{}'.format(t)] = "=MAX(AO2:AO{})".format(q)
            ws['AO{}'.format(u)] = "=MIN(AO2:AO{})".format(q)
            ws['AP{}'.format(r)] = "=ROUND(AVERAGE(AP2:AP{}),2)".format(q)
            ws['AP{}'.format(t)] = "=MAX(AP2:AP{})".format(q)
            ws['AP{}'.format(u)] = "=MIN(AP2:AP{})".format(q)
            ws['AQ{}'.format(r)] = "=MAX(AQ2:AQ{})".format(q)
            ws['AR{}'.format(r)] = "=MAX(AR2:AR{})".format(q)
            ws['AS{}'.format(r)] = "=MAX(AS2:AS{})".format(q)
            ws['AT{}'.format(r)] = "=MAX(AT2:AT{})".format(q)
            ws['AU{}'.format(r)] = "=MAX(AU2:AU{})".format(q)
            ws['AV{}'.format(r)] = "=MAX(AV2:AV{})".format(q)
            ws['AV{}'.format(s)] = "=MIN(AV2:AV{})".format(q)
            ws['AV{}'.format(t)] = "=ROUND(AVERAGE(AV2:AV{}),2)".format(q)
            ws['AW{}'.format(r)] = "=MAX(AW2:AW{})".format(q)
            ws['AW{}'.format(s)] = "=MIN(AW2:AW{})".format(q)
            ws['AW{}'.format(t)] = "=ROUND(AVERAGE(AW2:AW{}),2)".format(q)
            ws['AX{}'.format(r)] = "=MAX(AX2:AX{})".format(q)
            ws['AX{}'.format(s)] = "=MIN(AX2:AX{})".format(q)
            ws['AX{}'.format(t)] = "=ROUND(AVERAGE(AX2:AX{}),2)".format(q)
            ws['AY{}'.format(r)] = "=MAX(AY2:AY{})".format(q)
            ws['AY{}'.format(s)] = "=MIN(AY2:AY{})".format(q)
            ws['AY{}'.format(t)] = "=ROUND(AVERAGE(AY2:AY{}),2)".format(q)
            ws['AZ{}'.format(r)] = "=MAX(AZ2:AZ{})".format(q)
            ws['AZ{}'.format(s)] = "=MIN(AZ2:AZ{})".format(q)
            ws['AZ{}'.format(t)] = "=ROUND(AVERAGE(AZ2:AZ{}),2)".format(q)
            ws['BA{}'.format(r)] = "=MAX(BA2:BA{})".format(q)
            ws['BA{}'.format(s)] = "=MIN(BA2:BA{})".format(q)
            ws['BA{}'.format(t)] = "=ROUND(AVERAGE(BA2:BA{}),2)".format(q)
            ws['BD{}'.format(r)] = "=SUM(BD2:BD{})".format(q)
            ws['BE{}'.format(r)] = "=SUM(BE2:BE{})".format(q)
            ws['BF{}'.format(r)] = "=SUM(BF2:BF{})".format(q)
            ws['BG{}'.format(r)] = "=SUM(BG2:BG{})".format(q)
            ws['BH{}'.format(r)] = "=SUM(BH2:BH{})".format(q)

            # iterasi 2 rata-rata - 1
            ws['BO{}'.format(
                r)] = "=IF($BD${}=0,$AK${},$AK${}-1)".format(r, r, r)
            ws['BO{}'.format(s)] = "=STDEV(BO2:BO{})".format(q)
            ws['BO{}'.format(t)] = "=MAX(BO2:BO{})".format(q)
            ws['BO{}'.format(u)] = "=MIN(BO2:BO{})".format(q)
            ws['BP{}'.format(
                r)] = "=IF($BE${}=0,$AL${},$AL${}-1)".format(r, r, r)
            ws['BP{}'.format(s)] = "=STDEV(BP2:BP{})".format(q)
            ws['BP{}'.format(t)] = "=MAX(BP2:BP{})".format(q)
            ws['BP{}'.format(u)] = "=MIN(BP2:BP{})".format(q)
            ws['BQ{}'.format(
                r)] = "=IF($BF${}=0,$AM${},$AM${}-1)".format(r, r, r)
            ws['BQ{}'.format(s)] = "=STDEV(BQ2:BQ{})".format(q)
            ws['BQ{}'.format(t)] = "=MAX(BQ2:BQ{})".format(q)
            ws['BQ{}'.format(u)] = "=MIN(BQ2:BQ{})".format(q)
            ws['BR{}'.format(
                r)] = "=IF($BG${}=0,$AN${},$AN${}-1)".format(r, r, r)
            ws['BR{}'.format(s)] = "=STDEV(BR2:BR{})".format(q)
            ws['BR{}'.format(t)] = "=MAX(BR2:BR{})".format(q)
            ws['BR{}'.format(u)] = "=MIN(BR2:BR{})".format(q)
            ws['BS{}'.format(
                r)] = "=IF($BH${}=0,$AO${},$AO${}-1)".format(r, r, r)
            ws['BS{}'.format(s)] = "=STDEV(BS2:BS{})".format(q)
            ws['BS{}'.format(t)] = "=MAX(BS2:BS{})".format(q)
            ws['BS{}'.format(u)] = "=MIN(BS2:BS{})".format(q)
            ws['BT{}'.format(r)] = "=ROUND(AVERAGE(BT2:BT{}),2)".format(q)
            ws['BT{}'.format(t)] = "=MAX(BT2:BT{})".format(q)
            ws['BT{}'.format(u)] = "=MIN(BT2:BT{})".format(q)
            ws['BU{}'.format(r)] = "=MAX(BU2:BU{})".format(q)
            ws['BV{}'.format(r)] = "=MAX(BV2:BV{})".format(q)
            ws['BW{}'.format(r)] = "=MAX(BW2:BW{})".format(q)
            ws['BX{}'.format(r)] = "=MAX(BX2:BX{})".format(q)
            ws['BY{}'.format(r)] = "=MAX(BY2:BY{})".format(q)
            ws['BZ{}'.format(r)] = "=MAX(BZ2:BZ{})".format(q)
            ws['BZ{}'.format(s)] = "=MIN(BZ2:BZ{})".format(q)
            ws['BZ{}'.format(t)] = "=ROUND(AVERAGE(BZ2:BZ{}),2)".format(q)
            ws['CA{}'.format(r)] = "=MAX(CA2:CA{})".format(q)
            ws['CA{}'.format(s)] = "=MIN(CA2:CA{})".format(q)
            ws['CA{}'.format(t)] = "=ROUND(AVERAGE(CA2:CA{}),2)".format(q)
            ws['CB{}'.format(r)] = "=MAX(CB2:CB{})".format(q)
            ws['CB{}'.format(s)] = "=MIN(CB2:CB{})".format(q)
            ws['CB{}'.format(t)] = "=ROUND(AVERAGE(CB2:CB{}),2)".format(q)
            ws['CC{}'.format(r)] = "=MAX(CC2:CC{})".format(q)
            ws['CC{}'.format(s)] = "=MIN(CC2:CC{})".format(q)
            ws['CC{}'.format(t)] = "=ROUND(AVERAGE(CC2:CC{}),2)".format(q)
            ws['CD{}'.format(r)] = "=MAX(CD2:CD{})".format(q)
            ws['CD{}'.format(s)] = "=MIN(CD2:CD{})".format(q)
            ws['CD{}'.format(t)] = "=ROUND(AVERAGE(CD2:CD{}),2)".format(q)
            ws['CE{}'.format(r)] = "=MAX(CE2:CE{})".format(q)
            ws['CE{}'.format(s)] = "=MIN(CE2:CE{})".format(q)
            ws['CE{}'.format(t)] = "=ROUND(AVERAGE(CE2:CE{}),2)".format(q)
            ws['CH{}'.format(r)] = "=SUM(CH2:CH{})".format(q)
            ws['CI{}'.format(r)] = "=SUM(CI2:CI{})".format(q)
            ws['CJ{}'.format(r)] = "=SUM(CJ2:CJ{})".format(q)
            ws['CK{}'.format(r)] = "=SUM(CK2:CK{})".format(q)
            ws['CL{}'.format(r)] = "=SUM(CL2:CL{})".format(q)

            # iterasi 3 rata-rata - 1
            ws['CS{}'.format(
                r)] = "=IF($CH${}=0,$BO${},$BO${}-1)".format(r, r, r)
            ws['CS{}'.format(s)] = "=STDEV(CS2:CS{})".format(q)
            ws['CS{}'.format(t)] = "=MAX(CS2:CS{})".format(q)
            ws['CS{}'.format(u)] = "=MIN(CS2:CS{})".format(q)
            ws['CT{}'.format(
                r)] = "=IF($CI${}=0,$BP${},$BP${}-1)".format(r, r, r)
            ws['CT{}'.format(s)] = "=STDEV(CT2:CT{})".format(q)
            ws['CT{}'.format(t)] = "=MAX(CT2:CT{})".format(q)
            ws['CT{}'.format(u)] = "=MIN(CT2:CT{})".format(q)
            ws['CU{}'.format(
                r)] = "=IF($CJ${}=0,$BQ${},$BQ${}-1)".format(r, r, r)
            ws['CU{}'.format(s)] = "=STDEV(CU2:CU{})".format(q)
            ws['CU{}'.format(t)] = "=MAX(CU2:CU{})".format(q)
            ws['CU{}'.format(u)] = "=MIN(CU2:CU{})".format(q)
            ws['CV{}'.format(
                r)] = "=IF($CK${}=0,$BR${},$BR${}-1)".format(r, r, r)
            ws['CV{}'.format(s)] = "=STDEV(CV2:CV{})".format(q)
            ws['CV{}'.format(t)] = "=MAX(CV2:CV{})".format(q)
            ws['CV{}'.format(u)] = "=MIN(CV2:CV{})".format(q)
            ws['CW{}'.format(
                r)] = "=IF($CL${}=0,$BS${},$BS${}-1)".format(r, r, r)
            ws['CW{}'.format(s)] = "=STDEV(CW2:CW{})".format(q)
            ws['CW{}'.format(t)] = "=MAX(CW2:CW{})".format(q)
            ws['CW{}'.format(u)] = "=MIN(CW2:CW{})".format(q)
            ws['CX{}'.format(r)] = "=ROUND(AVERAGE(CX2:CX{}),2)".format(q)
            ws['CX{}'.format(t)] = "=MAX(CX2:CX{})".format(q)
            ws['CX{}'.format(u)] = "=MIN(CX2:CX{})".format(q)
            ws['CY{}'.format(r)] = "=MAX(CY2:CY{})".format(q)
            ws['CZ{}'.format(r)] = "=MAX(CZ2:CZ{})".format(q)
            ws['DA{}'.format(r)] = "=MAX(DA2:DA{})".format(q)
            ws['DB{}'.format(r)] = "=MAX(DB2:DB{})".format(q)
            ws['DC{}'.format(r)] = "=MAX(DC2:DC{})".format(q)
            ws['DD{}'.format(r)] = "=MAX(DD2:DD{})".format(q)
            ws['DD{}'.format(s)] = "=MIN(DD2:DD{})".format(q)
            ws['DD{}'.format(t)] = "=ROUND(AVERAGE(DD2:DD{}),2)".format(q)
            ws['DE{}'.format(r)] = "=MAX(DE2:DE{})".format(q)
            ws['DE{}'.format(s)] = "=MIN(DE2:DE{})".format(q)
            ws['DE{}'.format(t)] = "=ROUND(AVERAGE(DE2:DE{}),2)".format(q)
            ws['DF{}'.format(r)] = "=MAX(DF2:DF{})".format(q)
            ws['DF{}'.format(s)] = "=MIN(DF2:DF{})".format(q)
            ws['DF{}'.format(t)] = "=ROUND(AVERAGE(DF2:DF{}),2)".format(q)
            ws['DG{}'.format(r)] = "=MAX(DG2:DG{})".format(q)
            ws['DG{}'.format(s)] = "=MIN(DG2:DG{})".format(q)
            ws['DG{}'.format(t)] = "=ROUND(AVERAGE(DG2:DG{}),2)".format(q)
            ws['DH{}'.format(r)] = "=MAX(DH2:DH{})".format(q)
            ws['DH{}'.format(s)] = "=MIN(DH2:DH{})".format(q)
            ws['DH{}'.format(t)] = "=ROUND(AVERAGE(DH2:DH{}),2)".format(q)
            ws['DI{}'.format(r)] = "=MAX(DI2:DI{})".format(q)
            ws['DI{}'.format(s)] = "=MIN(DI2:DI{})".format(q)
            ws['DI{}'.format(t)] = "=ROUND(AVERAGE(DI2:DI{}),2)".format(q)
            ws['DL{}'.format(r)] = "=SUM(DL2:DL{})".format(q)
            ws['DM{}'.format(r)] = "=SUM(DM2:DM{})".format(q)
            ws['DN{}'.format(r)] = "=SUM(DN2:DN{})".format(q)
            ws['DO{}'.format(r)] = "=SUM(DO2:DO{})".format(q)
            ws['DP{}'.format(r)] = "=SUM(DP2:DP{})".format(q)

            # iterasi 4 rata-rata - 1
            ws['DW{}'.format(
                r)] = "=IF($DL${}=0,$CS${},$CS${}-1)".format(r, r, r)
            ws['DW{}'.format(s)] = "=STDEV(DW2:DW{})".format(q)
            ws['DW{}'.format(t)] = "=MAX(DW2:DW{})".format(q)
            ws['DW{}'.format(u)] = "=MIN(DW2:DW{})".format(q)
            ws['DX{}'.format(
                r)] = "=IF($DM${}=0,$CT${},$CT${}-1)".format(r, r, r)
            ws['DX{}'.format(s)] = "=STDEV(DX2:DX{})".format(q)
            ws['DX{}'.format(t)] = "=MAX(DX2:DX{})".format(q)
            ws['DX{}'.format(u)] = "=MIN(DX2:DX{})".format(q)
            ws['DY{}'.format(
                r)] = "=IF($DN${}=0,$CU${},$CU${}-1)".format(r, r, r)
            ws['DY{}'.format(s)] = "=STDEV(DY2:DY{})".format(q)
            ws['DY{}'.format(t)] = "=MAX(DY2:DY{})".format(q)
            ws['DY{}'.format(u)] = "=MIN(DY2:DY{})".format(q)
            ws['DZ{}'.format(
                r)] = "=IF($DO${}=0,$CV${},$CV${}-1)".format(r, r, r)
            ws['DZ{}'.format(s)] = "=STDEV(DZ2:DZ{})".format(q)
            ws['DZ{}'.format(t)] = "=MAX(DZ2:DZ{})".format(q)
            ws['DZ{}'.format(u)] = "=MIN(DZ2:DZ{})".format(q)
            ws['EA{}'.format(
                r)] = "=IF($DP${}=0,$CW${},$CW${}-1)".format(r, r, r)
            ws['EA{}'.format(s)] = "=STDEV(EA2:EA{})".format(q)
            ws['EA{}'.format(t)] = "=MAX(EA2:EA{})".format(q)
            ws['EA{}'.format(u)] = "=MIN(EA2:EA{})".format(q)
            ws['EB{}'.format(r)] = "=ROUND(AVERAGE(EB2:EB{}),2)".format(q)
            ws['EB{}'.format(t)] = "=MAX(EB2:EB{})".format(q)
            ws['EB{}'.format(u)] = "=MIN(EB2:EB{})".format(q)
            ws['EC{}'.format(r)] = "=MAX(EC2:EC{})".format(q)
            ws['ED{}'.format(r)] = "=MAX(ED2:ED{})".format(q)
            ws['EE{}'.format(r)] = "=MAX(EE2:EE{})".format(q)
            ws['EF{}'.format(r)] = "=MAX(EF2:EF{})".format(q)
            ws['EG{}'.format(r)] = "=MAX(EG2:EG{})".format(q)
            ws['EH{}'.format(r)] = "=MAX(EH2:EH{})".format(q)
            ws['EH{}'.format(s)] = "=MIN(EH2:EH{})".format(q)
            ws['EH{}'.format(t)] = "=ROUND(AVERAGE(EH2:EH{}),2)".format(q)
            ws['EI{}'.format(r)] = "=MAX(EI2:EI{})".format(q)
            ws['EI{}'.format(s)] = "=MIN(EI2:EI{})".format(q)
            ws['EI{}'.format(t)] = "=ROUND(AVERAGE(EI2:EI{}),2)".format(q)
            ws['EJ{}'.format(r)] = "=MAX(EJ2:EJ{})".format(q)
            ws['EJ{}'.format(s)] = "=MIN(EJ2:EJ{})".format(q)
            ws['EJ{}'.format(t)] = "=ROUND(AVERAGE(EJ2:EJ{}),2)".format(q)
            ws['EK{}'.format(r)] = "=MAX(EK2:EK{})".format(q)
            ws['EK{}'.format(s)] = "=MIN(EK2:EK{})".format(q)
            ws['EK{}'.format(t)] = "=ROUND(AVERAGE(EK2:EK{}),2)".format(q)
            ws['EL{}'.format(r)] = "=MAX(EL2:EL{})".format(q)
            ws['EL{}'.format(s)] = "=MIN(EL2:EL{})".format(q)
            ws['EL{}'.format(t)] = "=ROUND(AVERAGE(EL2:EL{}),2)".format(q)
            ws['EM{}'.format(r)] = "=MAX(EM2:EM{})".format(q)
            ws['EM{}'.format(s)] = "=MIN(EM2:EM{})".format(q)
            ws['EM{}'.format(t)] = "=ROUND(AVERAGE(EM2:EM{}),2)".format(q)
            ws['EP{}'.format(r)] = "=SUM(EP2:EP{})".format(q)
            ws['EQ{}'.format(r)] = "=SUM(EQ2:EQ{})".format(q)
            ws['ER{}'.format(r)] = "=SUM(ER2:ER{})".format(q)
            ws['ES{}'.format(r)] = "=SUM(ES2:ES{})".format(q)
            ws['ET{}'.format(r)] = "=SUM(ET2:ET{})".format(q)

            # iterasi 5 rata-rata - 1
            ws['FA{}'.format(
                r)] = "=IF($EP${}=0,$DW${},$DW${}-1)".format(r, r, r)
            ws['FA{}'.format(s)] = "=STDEV(FA2:FA{})".format(q)
            ws['FA{}'.format(t)] = "=MAX(FA2:FA{})".format(q)
            ws['FA{}'.format(u)] = "=MIN(FA2:FA{})".format(q)
            ws['FB{}'.format(
                r)] = "=IF($EQ${}=0,$DX${},$DX${}-1)".format(r, r, r)
            ws['FB{}'.format(s)] = "=STDEV(FB2:FB{})".format(q)
            ws['FB{}'.format(t)] = "=MAX(FB2:FB{})".format(q)
            ws['FB{}'.format(u)] = "=MIN(FB2:FB{})".format(q)
            ws['FC{}'.format(
                r)] = "=IF($ER${}=0,$DY${},$DY${}-1)".format(r, r, r)
            ws['FC{}'.format(s)] = "=STDEV(FC2:FC{})".format(q)
            ws['FC{}'.format(t)] = "=MAX(FC2:FC{})".format(q)
            ws['FC{}'.format(u)] = "=MIN(FC2:FC{})".format(q)
            ws['FD{}'.format(
                r)] = "=IF($ES${}=0,$DZ${},$DZ${}-1)".format(r, r, r)
            ws['FD{}'.format(s)] = "=STDEV(FD2:FD{})".format(q)
            ws['FD{}'.format(t)] = "=MAX(FD2:FD{})".format(q)
            ws['FD{}'.format(u)] = "=MIN(FD2:FD{})".format(q)
            ws['FE{}'.format(
                r)] = "=IF($ET${}=0,$EA${},$EA${}-1)".format(r, r, r)
            ws['FE{}'.format(s)] = "=STDEV(FE2:FE{})".format(q)
            ws['FE{}'.format(t)] = "=MAX(FE2:FE{})".format(q)
            ws['FE{}'.format(u)] = "=MIN(FE2:FE{})".format(q)
            ws['FF{}'.format(r)] = "=ROUND(AVERAGE(FF2:FF{}),2)".format(q)
            ws['FF{}'.format(t)] = "=MAX(FF2:FF{})".format(q)
            ws['FF{}'.format(u)] = "=MIN(FF2:FF{})".format(q)
            ws['FG{}'.format(r)] = "=MAX(FG2:FG{})".format(q)
            ws['FH{}'.format(r)] = "=MAX(FH2:FH{})".format(q)
            ws['FI{}'.format(r)] = "=MAX(FI2:FI{})".format(q)
            ws['FJ{}'.format(r)] = "=MAX(FJ2:FJ{})".format(q)
            ws['FK{}'.format(r)] = "=MAX(FK2:FK{})".format(q)
            ws['FL{}'.format(r)] = "=MAX(FL2:FL{})".format(q)
            ws['FL{}'.format(s)] = "=MIN(FL2:FL{})".format(q)
            ws['FL{}'.format(t)] = "=ROUND(AVERAGE(FL2:FL{}),2)".format(q)
            ws['FM{}'.format(r)] = "=MAX(FM2:FM{})".format(q)
            ws['FM{}'.format(s)] = "=MIN(FM2:FM{})".format(q)
            ws['FM{}'.format(t)] = "=ROUND(AVERAGE(FM2:FM{}),2)".format(q)
            ws['FN{}'.format(r)] = "=MAX(FN2:FN{})".format(q)
            ws['FN{}'.format(s)] = "=MIN(FN2:FN{})".format(q)
            ws['FN{}'.format(t)] = "=ROUND(AVERAGE(FN2:FN{}),2)".format(q)
            ws['FO{}'.format(r)] = "=MAX(FO2:FO{})".format(q)
            ws['FO{}'.format(s)] = "=MIN(FO2:FO{})".format(q)
            ws['FO{}'.format(t)] = "=ROUND(AVERAGE(FO2:FO{}),2)".format(q)
            ws['FP{}'.format(r)] = "=MAX(FP2:FP{})".format(q)
            ws['FP{}'.format(s)] = "=MIN(FP2:FP{})".format(q)
            ws['FP{}'.format(t)] = "=ROUND(AVERAGE(FP2:FP{}),2)".format(q)
            ws['FQ{}'.format(r)] = "=MAX(FQ2:FQ{})".format(q)
            ws['FQ{}'.format(s)] = "=MIN(FQ2:FQ{})".format(q)
            ws['FQ{}'.format(t)] = "=ROUND(AVERAGE(FQ2:FQ{}),2)".format(q)
            ws['FT{}'.format(r)] = "=SUM(FT2:FT{})".format(q)
            ws['FU{}'.format(r)] = "=SUM(FU2:FU{})".format(q)
            ws['FV{}'.format(r)] = "=SUM(FV2:FV{})".format(q)
            ws['FW{}'.format(r)] = "=SUM(FW2:FW{})".format(q)
            ws['FX{}'.format(r)] = "=SUM(FX2:FX{})".format(q)

            # Z Score
            ws['B1'] = 'NAMA_SISWA_1'
            ws['C1'] = 'NOMOR_NF_1'
            ws['D1'] = 'KELAS_1'
            ws['E1'] = 'NAMA_SEKOLAH_1'
            ws['F1'] = 'LOKASI_1'
            ws['G1'] = 'MAT_1'
            ws['H1'] = 'IND_1'
            ws['I1'] = 'ENG_1'
            ws['J1'] = 'IPA_1'
            ws['K1'] = 'IPS_1'
            ws['L1'] = 'JML_1'
            ws['M1'] = 'Z_MAT_1'
            ws['N1'] = 'Z_IND_1'
            ws['O1'] = 'Z_ENG_1'
            ws['P1'] = 'Z_IPA_1'
            ws['Q1'] = 'Z_IPS_1'
            ws['R1'] = 'S_MAT_1'
            ws['S1'] = 'S_IND_1'
            ws['T1'] = 'S_ENG_1'
            ws['U1'] = 'S_IPA_1'
            ws['V1'] = 'S_IPS_1'
            ws['W1'] = 'S_JML_1'
            ws['X1'] = 'RANK_NAS._1'
            ws['Y1'] = 'RANK_LOK._1'
            ws['M1'].font = Font(bold=False, name='Calibri', size=11)
            ws['N1'].font = Font(bold=False, name='Calibri', size=11)
            ws['O1'].font = Font(bold=False, name='Calibri', size=11)
            ws['P1'].font = Font(bold=False, name='Calibri', size=11)
            ws['Q1'].font = Font(bold=False, name='Calibri', size=11)
            ws['R1'].font = Font(bold=False, name='Calibri', size=11)
            ws['S1'].font = Font(bold=False, name='Calibri', size=11)
            ws['T1'].font = Font(bold=False, name='Calibri', size=11)
            ws['U1'].font = Font(bold=False, name='Calibri', size=11)
            ws['V1'].font = Font(bold=False, name='Calibri', size=11)
            ws['W1'].font = Font(bold=False, name='Calibri', size=11)
            ws['X1'].font = Font(bold=False, name='Calibri', size=11)
            ws['Y1'].font = Font(bold=False, name='Calibri', size=11)
        # FILL
            ws['B1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['C1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['D1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['E1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['F1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['G1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['H1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['I1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['J1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['K1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['L1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['M1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['N1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['O1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['P1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['Q1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['R1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['S1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['T1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['U1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['V1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['W1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['X1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['Y1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            # tambahan
            ws['Z1'] = 'MAT_20_1'
            ws['AA1'] = 'IND_20_1'
            ws['AB1'] = 'ENG_20_1'
            ws['AC1'] = 'IPA_20_1'
            ws['AD1'] = 'IPS_20_1'
            ws['Z1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AA1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AB1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AC1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AD1'].font = Font(bold=False, name='Calibri', size=11)
            ws['Z1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['AA1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['AB1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['AC1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')
            ws['AD1'].fill = PatternFill(
                fill_type='solid', start_color='00FF6600', end_color='00FF6600')

            for row in range(2, q+1):
                ws['L{}'.format(
                    row)] = '=SUM(G{}:K{})'.format(row, row, row)
                ws['M{}'.format(
                    row)] = '=IFERROR(ROUND(IF(G{}="","",(G{}-G${})/G${}),2),"")'.format(row, row, r, s)
                ws['N{}'.format(
                    row)] = '=IFERROR(ROUND(IF(H{}="","",(H{}-H${})/H${}),2),"")'.format(row, row, r, s)
                ws['O{}'.format(
                    row)] = '=IFERROR(ROUND(IF(I{}="","",(I{}-I${})/I${}),2),"")'.format(row, row, r, s)
                ws['P{}'.format(
                    row)] = '=IFERROR(ROUND(IF(J{}="","",(J{}-J${})/J${}),2),"")'.format(row, row, r, s)
                ws['Q{}'.format(
                    row)] = '=IFERROR(ROUND(IF(K{}="","",(K{}-K${})/K${}),2),"")'.format(row, row, r, s)
                ws['R{}'.format(
                    row)] = '=IFERROR(ROUND(IF(G{}="","",IF(70+30*M{}/$M${}<20,20,70+30*M{}/$M${})),2),"")'.format(row, row, r, row, r)
                ws['S{}'.format(
                    row)] = '=IFERROR(ROUND(IF(H{}="","",IF(70+30*N{}/$N${}<20,20,70+30*N{}/$N${})),2),"")'.format(row, row, r, row, r)
                ws['T{}'.format(
                    row)] = '=IFERROR(ROUND(IF(I{}="","",IF(70+30*O{}/$O${}<20,20,70+30*O{}/$O${})),2),"")'.format(row, row, r, row, r)
                ws['U{}'.format(
                    row)] = '=IFERROR(ROUND(IF(J{}="","",IF(70+30*P{}/$P${}<20,20,70+30*P{}/$P${})),2),"")'.format(row, row, r, row, r)
                ws['V{}'.format(
                    row)] = '=IFERROR(ROUND(IF(K{}="","",IF(70+30*Q{}/$Q${}<20,20,70+30*Q{}/$Q${})),2),"")'.format(row, row, r, row, r)

                ws['W{}'.format(row)] = '=IF(SUM(R{}:V{})=0,"",SUM(R{}:V{}))'.format(
                    row, row, row, row)
                ws['X{}'.format(row)] = '=IF(W{}="","",RANK(W{},$W$2:$W${}))'.format(
                    row, row, q)
                ws['Y{}'.format(
                    row)] = '=IF(X{}="","",COUNTIFS($F$2:$F${},F{},$X$2:$X${},"<"&X{})+1)'.format(row, q, row, q, row)
                # TAMBAHAN
                ws['Z{}'.format(row)] = '=IF($G${}=25,IF(AND(G{}>4,R{}=20),1,""),IF($G${}=30,IF(AND(G{}>5,R{}=20),1,""),IF($G${}=35,IF(AND(G{}>6,R{}=20),1,""),IF($G${}=40,IF(AND(G{}>7,R{}=20),1,""),IF($G${}=45,IF(AND(G{}>8,R{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['AA{}'.format(row)] = '=IF($H${}=25,IF(AND(H{}>4,S{}=20),1,""),IF($H${}=30,IF(AND(H{}>5,S{}=20),1,""),IF($H${}=35,IF(AND(H{}>6,S{}=20),1,""),IF($H${}=40,IF(AND(H{}>7,S{}=20),1,""),IF($H${}=45,IF(AND(H{}>8,S{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['AB{}'.format(row)] = '=IF($I${}=25,IF(AND(I{}>4,T{}=20),1,""),IF($I${}=30,IF(AND(I{}>5,T{}=20),1,""),IF($I${}=35,IF(AND(I{}>6,T{}=20),1,""),IF($I${}=40,IF(AND(I{}>7,T{}=20),1,""),IF($I${}=45,IF(AND(I{}>8,T{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['AC{}'.format(row)] = '=IF($J${}=25,IF(AND(J{}>4,U{}=20),1,""),IF($J${}=30,IF(AND(J{}>5,U{}=20),1,""),IF($J${}=35,IF(AND(J{}>6,U{}=20),1,""),IF($J${}=40,IF(AND(J{}>7,U{}=20),1,""),IF($J${}=45,IF(AND(J{}>8,U{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['AD{}'.format(row)] = '=IF($K${}=25,IF(AND(K{}>4,V{}=20),1,""),IF($K${}=30,IF(AND(K{}>5,V{}=20),1,""),IF($K${}=35,IF(AND(K{}>6,V{}=20),1,""),IF($K${}=40,IF(AND(K{}>7,V{}=20),1,""),IF($K${}=45,IF(AND(K{}>8,V{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)

        # new Z Score
            ws['AF1'] = 'NAMA_SISWA_2'
            ws['AG1'] = 'NOMOR_NF_2'
            ws['AH1'] = 'KELAS_2'
            ws['AI1'] = 'NAMA_SEKOLAH_2'
            ws['AJ1'] = 'LOKASI_2'
            ws['AK1'] = 'MAT_2'
            ws['AL1'] = 'IND_2'
            ws['AM1'] = 'ENG_2'
            ws['AN1'] = 'IPA_2'
            ws['AO1'] = 'IPS_2'
            ws['AP1'] = 'JML_2'
            ws['AQ1'] = 'Z_MAT_2'
            ws['AR1'] = 'Z_IND_2'
            ws['AS1'] = 'Z_ENG_2'
            ws['AT1'] = 'Z_IPA_2'
            ws['AU1'] = 'Z_IPS_2'
            ws['AV1'] = 'S_MAT_2'
            ws['AW1'] = 'S_IND_2'
            ws['AX1'] = 'S_ENG_2'
            ws['AY1'] = 'S_IPA_2'
            ws['AZ1'] = 'S_IPS_2'
            ws['BA1'] = 'S_JML_2'
            ws['BB1'] = 'RANK_NAS._2'
            ws['BC1'] = 'RANK_LOK._2'
            ws['AQ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AR1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AS1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AT1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AU1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AV1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AW1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AX1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AY1'].font = Font(bold=False, name='Calibri', size=11)
            ws['AZ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BA1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BB1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BC1'].font = Font(bold=False, name='Calibri', size=11)

            # FILL
            ws['AF1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AG1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AH1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AI1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AJ1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AK1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AL1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AM1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AN1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AO1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AP1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AQ1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AR1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AS1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AT1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AU1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AV1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AW1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AX1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AY1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['AZ1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BA1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BB1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BC1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')

            # tambahan
            ws['BD1'] = 'MAT_20_2'
            ws['BE1'] = 'IND_20_2'
            ws['BF1'] = 'ENG_20_2'
            ws['BG1'] = 'IPA_20_2'
            ws['BH1'] = 'IPS_20_2'
            ws['BD1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BE1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BF1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BG1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BH1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BD1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BE1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BF1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BG1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')
            ws['BH1'].fill = PatternFill(
                fill_type='solid', start_color='31E1F7', end_color='31E1F7')

            for row in range(2, q+1):
                ws['AF{}'.format(row)] = '=B{}'.format(row)
                ws['AG{}'.format(row)] = '=C{}'.format(row, row)
                ws['AH{}'.format(row)] = '=D{}'.format(row, row)
                ws['AI{}'.format(row)] = '=E{}'.format(row, row)
                ws['AJ{}'.format(row)] = '=F{}'.format(row, row)
                ws['AK{}'.format(row)] = '=IF(G{}="","",G{})'.format(row, row)
                ws['AL{}'.format(row)] = '=IF(H{}="","",H{})'.format(row, row)
                ws['AM{}'.format(row)] = '=IF(I{}="","",I{})'.format(row, row)
                ws['AN{}'.format(row)] = '=IF(J{}="","",J{})'.format(row, row)
                ws['AO{}'.format(row)] = '=IF(K{}="","",K{})'.format(row, row)
                ws['AP{}'.format(row)] = '=IF(L{}="","",L{})'.format(row, row)
                ws['AQ{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AK{}="","",(AK{}-AK${})/AK${}),2),"")'.format(row, row, r, s)
                ws['AR{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AL{}="","",(AL{}-AL${})/AL${}),2),"")'.format(row, row, r, s)
                ws['AS{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AM{}="","",(AM{}-AM${})/AM${}),2),"")'.format(row, row, r, s)
                ws['AT{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AN{}="","",(AN{}-AN${})/AN${}),2),"")'.format(row, row, r, s)
                ws['AU{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AO{}="","",(AO{}-AO${})/AO${}),2),"")'.format(row, row, r, s)
                ws['AV{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AK{}="","",IF(70+30*AQ{}/$AQ${}<20,20,70+30*AQ{}/$AQ${})),2),"")'.format(row, row, r, row, r)
                ws['AW{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AL{}="","",IF(70+30*AR{}/$AR${}<20,20,70+30*AR{}/$AR${})),2),"")'.format(row, row, r, row, r)
                ws['AX{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AM{}="","",IF(70+30*AS{}/$AS${}<20,20,70+30*AS{}/$AS${})),2),"")'.format(row, row, r, row, r)
                ws['AY{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AN{}="","",IF(70+30*AT{}/$AT${}<20,20,70+30*AT{}/$AT${})),2),"")'.format(row, row, r, row, r)
                ws['AZ{}'.format(
                    row)] = '=IFERROR(ROUND(IF(AO{}="","",IF(70+30*AU{}/$AU${}<20,20,70+30*AU{}/$AU${})),2),"")'.format(row, row, r, row, r)

                ws['BA{}'.format(row)] = '=IF(SUM(AV{}:AZ{})=0,"",SUM(AV{}:AZ{}))'.format(
                    row, row, row, row)
                ws['BB{}'.format(row)] = '=IF(BA{}="","",RANK(BA{},$BA$2:$BA${}))'.format(
                    row, row, q)
                ws['BC{}'.format(
                    row)] = '=IF(BB{}="","",COUNTIFS($AJ$2:$AJ${},F{},$BB$2:$BB${},"<"&BB{})+1)'.format(row, q, row, q, row)
            #     TAMBAHAN
                ws['BD{}'.format(row)] = '=IF($G${}=25,IF(AND(AK{}>4,AV{}=20),1,""),IF($G${}=30,IF(AND(AK{}>5,AV{}=20),1,""),IF($G${}=35,IF(AND(AK{}>6,AV{}=20),1,""),IF($G${}=40,IF(AND(AK{}>7,AV{}=20),1,""),IF($G${}=45,IF(AND(AK{}>8,AV{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['BE{}'.format(row)] = '=IF($H${}=25,IF(AND(AL{}>4,AW{}=20),1,""),IF($H${}=30,IF(AND(AL{}>5,AW{}=20),1,""),IF($H${}=35,IF(AND(AL{}>6,AW{}=20),1,""),IF($H${}=40,IF(AND(AL{}>7,AW{}=20),1,""),IF($H${}=45,IF(AND(AL{}>8,AW{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['BF{}'.format(row)] = '=IF($I${}=25,IF(AND(AM{}>4,AX{}=20),1,""),IF($I${}=30,IF(AND(AM{}>5,AX{}=20),1,""),IF($I${}=35,IF(AND(AM{}>6,AX{}=20),1,""),IF($I${}=40,IF(AND(AM{}>7,AX{}=20),1,""),IF($I${}=45,IF(AND(AM{}>8,AX{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['BG{}'.format(row)] = '=IF($J${}=25,IF(AND(AN{}>4,AY{}=20),1,""),IF($J${}=30,IF(AND(AN{}>5,AY{}=20),1,""),IF($J${}=35,IF(AND(AN{}>6,AY{}=20),1,""),IF($J${}=40,IF(AND(AN{}>7,AY{}=20),1,""),IF($J${}=45,IF(AND(AN{}>8,AY{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['BH{}'.format(row)] = '=IF($K${}=25,IF(AND(AO{}>4,AZ{}=20),1,""),IF($K${}=30,IF(AND(AO{}>5,AZ{}=20),1,""),IF($K${}=35,IF(AND(AO{}>6,AZ{}=20),1,""),IF($K${}=40,IF(AND(AO{}>7,AZ{}=20),1,""),IF($K${}=45,IF(AND(AO{}>8,AZ{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)

                # new Z Score [2]
            ws['BJ1'] = 'NAMA_SISWA_3'
            ws['BK1'] = 'NOMOR_NF_3'
            ws['BL1'] = 'KELAS_3'
            ws['BM1'] = 'NAMA_SEKOLAH_3'
            ws['BN1'] = 'LOKASI_3'
            ws['BO1'] = 'MAT_3'
            ws['BP1'] = 'IND_3'
            ws['BQ1'] = 'ENG_3'
            ws['BR1'] = 'IPA_3'
            ws['BS1'] = 'IPS_3'
            ws['BT1'] = 'JML_3'
            ws['BU1'] = 'Z_MAT_3'
            ws['BV1'] = 'Z_IND_3'
            ws['BW1'] = 'Z_ENG_3'
            ws['BX1'] = 'Z_IPA_3'
            ws['BY1'] = 'Z_IPS_3'
            ws['BZ1'] = 'S_MAT_3'
            ws['CA1'] = 'S_IND_3'
            ws['CB1'] = 'S_ENG_3'
            ws['CC1'] = 'S_IPA_3'
            ws['CD1'] = 'S_IPS_3'
            ws['CE1'] = 'S_JML_3'
            ws['CF1'] = 'RANK_NAS._3'
            ws['CG1'] = 'RANK_LOK._3'
            ws['BU1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BV1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BW1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BX1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BY1'].font = Font(bold=False, name='Calibri', size=11)
            ws['BZ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CA1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CB1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CC1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CD1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CE1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CF1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CG1'].font = Font(bold=False, name='Calibri', size=11)

            # FILL
            ws['BJ1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BK1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BL1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BM1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BN1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BO1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BP1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BQ1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BR1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BS1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BT1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BU1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BV1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BW1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BX1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BY1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['BZ1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CA1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CB1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CC1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CD1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CE1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CF1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CG1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')

            # tambahan
            ws['CH1'] = 'MAT_20_3'
            ws['CI1'] = 'IND_20_3'
            ws['CJ1'] = 'ENG_20_3'
            ws['CK1'] = 'IPA_20_3'
            ws['CL1'] = 'IPS_20_3'
            ws['CH1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CI1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CJ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CK1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CL1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CH1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CI1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CJ1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CK1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')
            ws['CL1'].fill = PatternFill(
                fill_type='solid', start_color='A1C298', end_color='A1C298')

            for row in range(2, q+1):
                ws['BJ{}'.format(row)] = '=B{}'.format(row)
                ws['BK{}'.format(row)] = '=C{}'.format(row, row)
                ws['BL{}'.format(row)] = '=D{}'.format(row, row)
                ws['BM{}'.format(row)] = '=E{}'.format(row, row)
                ws['BN{}'.format(row)] = '=F{}'.format(row, row)
                ws['BO{}'.format(row)] = '=IF(G{}="","",G{})'.format(row, row)
                ws['BP{}'.format(row)] = '=IF(H{}="","",H{})'.format(row, row)
                ws['BQ{}'.format(row)] = '=IF(I{}="","",I{})'.format(row, row)
                ws['BR{}'.format(row)] = '=IF(J{}="","",J{})'.format(row, row)
                ws['BS{}'.format(row)] = '=IF(K{}="","",K{})'.format(row, row)
                ws['BT{}'.format(row)] = '=IF(L{}="","",L{})'.format(row, row)
                ws['BU{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BO{}="","",(BO{}-BO${})/BO${}),2),"")'.format(row, row, r, s)
                ws['BV{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BP{}="","",(BP{}-BP${})/BP${}),2),"")'.format(row, row, r, s)
                ws['BW{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BQ{}="","",(BQ{}-BQ${})/BQ${}),2),"")'.format(row, row, r, s)
                ws['BX{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BR{}="","",(BR{}-BR${})/BR${}),2),"")'.format(row, row, r, s)
                ws['BY{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BS{}="","",(BS{}-BS${})/BS${}),2),"")'.format(row, row, r, s)
                ws['BZ{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BO{}="","",IF(70+30*BU{}/$BU${}<20,20,70+30*BU{}/$BU${})),2),"")'.format(row, row, r, row, r)
                ws['CA{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BP{}="","",IF(70+30*BV{}/$BV${}<20,20,70+30*BV{}/$BV${})),2),"")'.format(row, row, r, row, r)
                ws['CB{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BQ{}="","",IF(70+30*BW{}/$BW${}<20,20,70+30*BW{}/$BW${})),2),"")'.format(row, row, r, row, r)
                ws['CC{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BR{}="","",IF(70+30*BX{}/$BX${}<20,20,70+30*BX{}/$BX${})),2),"")'.format(row, row, r, row, r)
                ws['CD{}'.format(
                    row)] = '=IFERROR(ROUND(IF(BS{}="","",IF(70+30*BY{}/$BY${}<20,20,70+30*BY{}/$BY${})),2),"")'.format(row, row, r, row, r)

                ws['CE{}'.format(row)] = '=IF(SUM(BZ{}:CD{})=0,"",SUM(BZ{}:CD{}))'.format(
                    row, row, row, row)
                ws['CF{}'.format(row)] = '=IF(CE{}="","",RANK(CE{},$CE$2:$CE${}))'.format(
                    row, row, q)
                ws['CG{}'.format(
                    row)] = '=IF(CF{}="","",COUNTIFS($BN$2:$BN${},F{},$CF$2:$CF${},"<"&CF{})+1)'.format(row, q, row, q, row)
                #     TAMBAHAN
                ws['CH{}'.format(row)] = '=IF($G${}=25,IF(AND(BO{}>4,BZ{}=20),1,""),IF($G${}=30,IF(AND(BO{}>5,BZ{}=20),1,""),IF($G${}=35,IF(AND(BO{}>6,BZ{}=20),1,""),IF($G${}=40,IF(AND(BO{}>7,BZ{}=20),1,""),IF($G${}=45,IF(AND(BO{}>8,BZ{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['CI{}'.format(row)] = '=IF($H${}=25,IF(AND(BP{}>4,CA{}=20),1,""),IF($H${}=30,IF(AND(BP{}>5,CA{}=20),1,""),IF($H${}=35,IF(AND(BP{}>6,CA{}=20),1,""),IF($H${}=40,IF(AND(BP{}>7,CA{}=20),1,""),IF($H${}=45,IF(AND(BP{}>8,CA{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['CJ{}'.format(row)] = '=IF($I${}=25,IF(AND(BQ{}>4,CB{}=20),1,""),IF($I${}=30,IF(AND(BQ{}>5,CB{}=20),1,""),IF($I${}=35,IF(AND(BQ{}>6,CB{}=20),1,""),IF($I${}=40,IF(AND(BQ{}>7,CB{}=20),1,""),IF($I${}=45,IF(AND(BQ{}>8,CB{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['CK{}'.format(row)] = '=IF($J${}=25,IF(AND(BR{}>4,CC{}=20),1,""),IF($J${}=30,IF(AND(BR{}>5,CC{}=20),1,""),IF($J${}=35,IF(AND(BR{}>6,CC{}=20),1,""),IF($J${}=40,IF(AND(BR{}>7,CC{}=20),1,""),IF($J${}=45,IF(AND(BR{}>8,CC{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['CL{}'.format(row)] = '=IF($K${}=25,IF(AND(BS{}>4,CD{}=20),1,""),IF($K${}=30,IF(AND(BS{}>5,CD{}=20),1,""),IF($K${}=35,IF(AND(BS{}>6,CD{}=20),1,""),IF($K${}=40,IF(AND(BS{}>7,CD{}=20),1,""),IF($K${}=45,IF(AND(BS{}>8,CD{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)

                # new Z Score [3]
            ws['CN1'] = 'NAMA_SISWA_4'
            ws['CO1'] = 'NOMOR_NF_4'
            ws['CP1'] = 'KELAS_4'
            ws['CQ1'] = 'NAMA_SEKOLAH_4'
            ws['CR1'] = 'LOKASI_4'
            ws['CS1'] = 'MAT_4'
            ws['CT1'] = 'IND_4'
            ws['CU1'] = 'ENG_4'
            ws['CV1'] = 'IPA_4'
            ws['CW1'] = 'IPS_4'
            ws['CX1'] = 'JML_4'
            ws['CY1'] = 'Z_MAT_4'
            ws['CZ1'] = 'Z_IND_4'
            ws['DA1'] = 'Z_ENG_4'
            ws['DB1'] = 'Z_IPA_4'
            ws['DC1'] = 'Z_IPS_4'
            ws['DD1'] = 'S_MAT_4'
            ws['DE1'] = 'S_IND_4'
            ws['DF1'] = 'S_ENG_4'
            ws['DG1'] = 'S_IPA_4'
            ws['DH1'] = 'S_IPS_4'
            ws['DI1'] = 'S_JML_4'
            ws['DJ1'] = 'RANK_NAS._4'
            ws['DK1'] = 'RANK_LOK._4'
            ws['CY1'].font = Font(bold=False, name='Calibri', size=11)
            ws['CZ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DA1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DB1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DC1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DD1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DE1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DF1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DG1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DH1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DI1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DJ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DK1'].font = Font(bold=False, name='Calibri', size=11)

            # FILL
            ws['CN1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CO1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CP1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CQ1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CR1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CS1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CT1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CU1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CV1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CW1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CX1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CY1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['CZ1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DA1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DB1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DC1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DD1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DE1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DF1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DG1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DH1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DI1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DJ1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DK1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')

            # tambahan
            ws['DL1'] = 'MAT_20_4'
            ws['DM1'] = 'IND_20_4'
            ws['DN1'] = 'ENG_20_4'
            ws['DO1'] = 'IPA_20_4'
            ws['DP1'] = 'IPS_20_4'
            ws['DL1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DM1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DN1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DO1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DP1'].font = Font(bold=False, name='Calibri', size=11)
            ws['DL1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DM1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DN1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DO1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')
            ws['DP1'].fill = PatternFill(
                fill_type='solid', start_color='FFE9A0', end_color='FFE9A0')

            for row in range(2, q+1):
                ws['CN{}'.format(row)] = '=B{}'.format(row)
                ws['CO{}'.format(row)] = '=C{}'.format(row, row)
                ws['CP{}'.format(row)] = '=D{}'.format(row, row)
                ws['CQ{}'.format(row)] = '=E{}'.format(row, row)
                ws['CR{}'.format(row)] = '=F{}'.format(row, row)
                ws['CS{}'.format(row)] = '=IF(G{}="","",G{})'.format(row, row)
                ws['CT{}'.format(row)] = '=IF(H{}="","",H{})'.format(row, row)
                ws['CU{}'.format(row)] = '=IF(I{}="","",I{})'.format(row, row)
                ws['CV{}'.format(row)] = '=IF(J{}="","",J{})'.format(row, row)
                ws['CW{}'.format(row)] = '=IF(K{}="","",K{})'.format(row, row)
                ws['CX{}'.format(row)] = '=IF(L{}="","",L{})'.format(row, row)
                ws['CY{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CS{}="","",(CS{}-CS${})/CS${}),2),"")'.format(row, row, r, s)
                ws['CZ{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CT{}="","",(CT{}-CT${})/CT${}),2),"")'.format(row, row, r, s)
                ws['DA{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CU{}="","",(CU{}-CU${})/CU${}),2),"")'.format(row, row, r, s)
                ws['DB{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CV{}="","",(CV{}-CV${})/CV${}),2),"")'.format(row, row, r, s)
                ws['DC{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CW{}="","",(CW{}-CW${})/CW${}),2),"")'.format(row, row, r, s)
                ws['DD{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CY{}="","",IF(70+30*CY{}/$CY${}<20,20,70+30*CY{}/$CY${})),2),"")'.format(row, row, r, row, r)
                ws['DE{}'.format(
                    row)] = '=IFERROR(ROUND(IF(CZ{}="","",IF(70+30*CZ{}/$CZ${}<20,20,70+30*CZ{}/$CZ${})),2),"")'.format(row, row, r, row, r)
                ws['DF{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DA{}="","",IF(70+30*DA{}/$DA${}<20,20,70+30*DA{}/$DA${})),2),"")'.format(row, row, r, row, r)
                ws['DG{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DB{}="","",IF(70+30*DB{}/$DB${}<20,20,70+30*DB{}/$DB${})),2),"")'.format(row, row, r, row, r)
                ws['DH{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DC{}="","",IF(70+30*DC{}/$DC${}<20,20,70+30*DC{}/$DC${})),2),"")'.format(row, row, r, row, r)

                ws['DI{}'.format(row)] = '=IF(SUM(DD{}:DH{})=0,"",SUM(DD{}:DH{}))'.format(
                    row, row, row, row)
                ws['DJ{}'.format(row)] = '=IF(DI{}="","",RANK(DI{},$DI$2:$DI${}))'.format(
                    row, row, q)
                ws['DK{}'.format(
                    row)] = '=IF(DJ{}="","",COUNTIFS($CR$2:$CR${},F{},$DJ$2:$DJ${},"<"&DJ{})+1)'.format(row, q, row, q, row)
                #     TAMBAHAN
                ws['DL{}'.format(row)] = '=IF($G${}=25,IF(AND(CS{}>4,DD{}=20),1,""),IF($G${}=30,IF(AND(CS{}>5,DD{}=20),1,""),IF($G${}=35,IF(AND(CS{}>6,DD{}=20),1,""),IF($G${}=40,IF(AND(CS{}>7,DD{}=20),1,""),IF($G${}=45,IF(AND(CS{}>8,DD{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['DM{}'.format(row)] = '=IF($H${}=25,IF(AND(CT{}>4,DE{}=20),1,""),IF($H${}=30,IF(AND(CT{}>5,DE{}=20),1,""),IF($H${}=35,IF(AND(CT{}>6,DE{}=20),1,""),IF($H${}=40,IF(AND(CT{}>7,DE{}=20),1,""),IF($H${}=45,IF(AND(CT{}>8,DE{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['DN{}'.format(row)] = '=IF($I${}=25,IF(AND(CU{}>4,DF{}=20),1,""),IF($I${}=30,IF(AND(CU{}>5,DF{}=20),1,""),IF($I${}=35,IF(AND(CU{}>6,DF{}=20),1,""),IF($I${}=40,IF(AND(CU{}>7,DF{}=20),1,""),IF($I${}=45,IF(AND(CU{}>8,DF{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['DO{}'.format(row)] = '=IF($J${}=25,IF(AND(CV{}>4,DG{}=20),1,""),IF($J${}=30,IF(AND(CV{}>5,DG{}=20),1,""),IF($J${}=35,IF(AND(CV{}>6,DG{}=20),1,""),IF($J${}=40,IF(AND(CV{}>7,DG{}=20),1,""),IF($J${}=45,IF(AND(CV{}>8,DG{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['DP{}'.format(row)] = '=IF($K${}=25,IF(AND(CW{}>4,DH{}=20),1,""),IF($K${}=30,IF(AND(CW{}>5,DH{}=20),1,""),IF($K${}=35,IF(AND(CW{}>6,DH{}=20),1,""),IF($K${}=40,IF(AND(CW{}>7,DH{}=20),1,""),IF($K${}=45,IF(AND(CW{}>8,DH{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)

            # new Z Score [4]
            ws['DR1'] = 'NAMA_SISWA_5'
            ws['DS1'] = 'NOMOR_NF_5'
            ws['DT1'] = 'KELAS_5'
            ws['DU1'] = 'NAMA_SEKOLAH_5'
            ws['DV1'] = 'LOKASI_5'
            ws['DW1'] = 'MAT_5'
            ws['DX1'] = 'IND_5'
            ws['DY1'] = 'ENG_5'
            ws['DZ1'] = 'IPA_5'
            ws['EA1'] = 'IPS_5'
            ws['EB1'] = 'JML_5'
            ws['EC1'] = 'Z_MAT_5'
            ws['ED1'] = 'Z_IND_5'
            ws['EE1'] = 'Z_ENG_5'
            ws['EF1'] = 'Z_IPA_5'
            ws['EG1'] = 'Z_IPS_5'
            ws['EH1'] = 'S_MAT_5'
            ws['EI1'] = 'S_IND_5'
            ws['EJ1'] = 'S_ENG_5'
            ws['EK1'] = 'S_IPA_5'
            ws['EL1'] = 'S_IPS_5'
            ws['EM1'] = 'S_JML_5'
            ws['EN1'] = 'RANK_NAS._5'
            ws['EO1'] = 'RANK_LOK._5'
            ws['EC1'].font = Font(bold=False, name='Calibri', size=11)
            ws['ED1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EE1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EF1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EG1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EH1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EI1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EJ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EK1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EL1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EM1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EN1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EO1'].font = Font(bold=False, name='Calibri', size=11)

            # FILL
            ws['DR1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DS1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DT1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DU1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DV1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DW1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DX1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DY1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['DZ1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EA1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EB1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EC1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['ED1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EE1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EF1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EG1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EH1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EI1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EJ1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EK1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EL1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EM1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EN1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EO1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')

            # tambahan
            ws['EP1'] = 'MAT_20_5'
            ws['EQ1'] = 'IND_20_5'
            ws['ER1'] = 'ENG_20_5'
            ws['ES1'] = 'IPA_20_5'
            ws['ET1'] = 'IPS_20_5'
            ws['EP1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EQ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['ER1'].font = Font(bold=False, name='Calibri', size=11)
            ws['ES1'].font = Font(bold=False, name='Calibri', size=11)
            ws['ET1'].font = Font(bold=False, name='Calibri', size=11)
            ws['EP1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['EQ1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['ER1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['ES1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')
            ws['ET1'].fill = PatternFill(
                fill_type='solid', start_color='ECC5FB', end_color='ECC5FB')

            for row in range(2, q+1):
                ws['DR{}'.format(row)] = '=B{}'.format(row)
                ws['DS{}'.format(row)] = '=C{}'.format(row, row)
                ws['DT{}'.format(row)] = '=D{}'.format(row, row)
                ws['DU{}'.format(row)] = '=E{}'.format(row, row)
                ws['DV{}'.format(row)] = '=F{}'.format(row, row)
                ws['DW{}'.format(row)] = '=IF(G{}="","",G{})'.format(row, row)
                ws['DX{}'.format(row)] = '=IF(H{}="","",H{})'.format(row, row)
                ws['DY{}'.format(row)] = '=IF(I{}="","",I{})'.format(row, row)
                ws['DZ{}'.format(row)] = '=IF(J{}="","",J{})'.format(row, row)
                ws['EA{}'.format(row)] = '=IF(K{}="","",K{})'.format(row, row)
                ws['EB{}'.format(row)] = '=IF(L{}="","",L{})'.format(row, row)
                ws['EC{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DW{}="","",(DW{}-DW${})/DW${}),2),"")'.format(row, row, r, s)
                ws['ED{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DX{}="","",(DX{}-DX${})/DX${}),2),"")'.format(row, row, r, s)
                ws['EE{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DY{}="","",(DY{}-DY${})/DY${}),2),"")'.format(row, row, r, s)
                ws['EF{}'.format(
                    row)] = '=IFERROR(ROUND(IF(DZ{}="","",(DZ{}-DZ${})/DZ${}),2),"")'.format(row, row, r, s)
                ws['EG{}'.format(
                    row)] = '=IFERROR(ROUND(IF(EA{}="","",(EA{}-EA${})/EA${}),2),"")'.format(row, row, r, s)
                ws['EH{}'.format(
                    row)] = '=IFERROR(ROUND(IF(EC{}="","",IF(70+30*EC{}/$EC${}<20,20,70+30*EC{}/$EC${})),2),"")'.format(row, row, r, row, r)
                ws['EI{}'.format(
                    row)] = '=IFERROR(ROUND(IF(ED{}="","",IF(70+30*ED{}/$ED${}<20,20,70+30*ED{}/$ED${})),2),"")'.format(row, row, r, row, r)
                ws['EJ{}'.format(
                    row)] = '=IFERROR(ROUND(IF(EE{}="","",IF(70+30*EE{}/$EE${}<20,20,70+30*EE{}/$EE${})),2),"")'.format(row, row, r, row, r)
                ws['EK{}'.format(
                    row)] = '=IFERROR(ROUND(IF(EF{}="","",IF(70+30*EF{}/$EF${}<20,20,70+30*EF{}/$EF${})),2),"")'.format(row, row, r, row, r)
                ws['EL{}'.format(
                    row)] = '=IFERROR(ROUND(IF(EG{}="","",IF(70+30*EG{}/$EG${}<20,20,70+30*EG{}/$EG${})),2),"")'.format(row, row, r, row, r)

                ws['EM{}'.format(row)] = '=IF(SUM(EH{}:EL{})=0,"",SUM(EH{}:EL{}))'.format(
                    row, row, row, row)
                ws['EN{}'.format(row)] = '=IF(EM{}="","",RANK(EM{},$EM$2:$EM${}))'.format(
                    row, row, q)
                ws['EO{}'.format(
                    row)] = '=IF(EN{}="","",COUNTIFS($DV$2:$DV${},F{},$EN$2:$EN${},"<"&EN{})+1)'.format(row, q, row, q, row)
                #     TAMBAHAN
                ws['EP{}'.format(row)] = '=IF($G${}=25,IF(AND(DW{}>4,EH{}=20),1,""),IF($G${}=30,IF(AND(DW{}>5,EH{}=20),1,""),IF($G${}=35,IF(AND(DW{}>6,EH{}=20),1,""),IF($G${}=40,IF(AND(DW{}>7,EH{}=20),1,""),IF($G${}=45,IF(AND(DW{}>8,EH{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['EQ{}'.format(row)] = '=IF($H${}=25,IF(AND(DX{}>4,EI{}=20),1,""),IF($H${}=30,IF(AND(DX{}>5,EI{}=20),1,""),IF($H${}=35,IF(AND(DX{}>6,EI{}=20),1,""),IF($H${}=40,IF(AND(DX{}>7,EI{}=20),1,""),IF($H${}=45,IF(AND(DX{}>8,EI{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['ER{}'.format(row)] = '=IF($I${}=25,IF(AND(DY{}>4,EJ{}=20),1,""),IF($I${}=30,IF(AND(DY{}>5,EJ{}=20),1,""),IF($I${}=35,IF(AND(DY{}>6,EJ{}=20),1,""),IF($I${}=40,IF(AND(DY{}>7,EJ{}=20),1,""),IF($I${}=45,IF(AND(DY{}>8,EJ{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['ES{}'.format(row)] = '=IF($J${}=25,IF(AND(DZ{}>4,EK{}=20),1,""),IF($J${}=30,IF(AND(DZ{}>5,EK{}=20),1,""),IF($J${}=35,IF(AND(DZ{}>6,EK{}=20),1,""),IF($J${}=40,IF(AND(DZ{}>7,EK{}=20),1,""),IF($J${}=45,IF(AND(DZ{}>8,EK{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['ET{}'.format(row)] = '=IF($K${}=25,IF(AND(EA{}>4,EL{}=20),1,""),IF($K${}=30,IF(AND(EA{}>5,EL{}=20),1,""),IF($K${}=35,IF(AND(EA{}>6,EL{}=20),1,""),IF($K${}=40,IF(AND(EA{}>7,EL{}=20),1,""),IF($K${}=45,IF(AND(EA{}>8,EL{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)

                # new Z Score [5]
            ws['EV1'] = 'NAMA SISWA'
            ws['EW1'] = 'NOMOR NF'
            ws['EX1'] = 'KELAS'
            ws['EY1'] = 'NAMA SEKOLAH'
            ws['EZ1'] = 'LOKASI'
            ws['FA1'] = 'MAT'
            ws['FB1'] = 'IND'
            ws['FC1'] = 'ENG'
            ws['FD1'] = 'IPA'
            ws['FE1'] = 'IPS'
            ws['FF1'] = 'JML'
            ws['FG1'] = 'Z_MAT'
            ws['FH1'] = 'Z_IND'
            ws['FI1'] = 'Z_ENG'
            ws['FJ1'] = 'Z_IPA'
            ws['FK1'] = 'Z_IPS'
            ws['FL1'] = 'S_MAT'
            ws['FM1'] = 'S_IND'
            ws['FN1'] = 'S_ENG'
            ws['FO1'] = 'S_IPA'
            ws['FP1'] = 'S_IPS'
            ws['FQ1'] = 'S_JML'
            ws['FR1'] = 'RANK NAS.'
            ws['FS1'] = 'RANK LOK.'
            ws['FG1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FH1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FI1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FJ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FK1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FL1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FM1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FN1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FO1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FP1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FQ1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FR1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FS1'].font = Font(bold=False, name='Calibri', size=11)

            # FILL
            ws['EV1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['EW1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['EX1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['EY1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['EZ1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FA1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FB1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FC1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FD1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FE1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FF1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FG1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FH1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FI1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FJ1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FK1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FL1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FM1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FN1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FO1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FP1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FQ1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FR1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FS1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')

            # tambahan
            ws['FT1'] = 'MAT_20'
            ws['FU1'] = 'IND_20'
            ws['FV1'] = 'ENG_20'
            ws['FW1'] = 'IPA_20'
            ws['FX1'] = 'IPS_20'
            ws['FT1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FU1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FV1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FW1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FX1'].font = Font(bold=False, name='Calibri', size=11)
            ws['FT1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FU1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FV1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FW1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')
            ws['FX1'].fill = PatternFill(
                fill_type='solid', start_color='E1FFEE', end_color='E1FFEE')

            for row in range(2, q+1):
                ws['EV{}'.format(row)] = '=B{}'.format(row)
                ws['EW{}'.format(row)] = '=C{}'.format(row, row)
                ws['EX{}'.format(row)] = '=D{}'.format(row, row)
                ws['EY{}'.format(row)] = '=E{}'.format(row, row)
                ws['EZ{}'.format(row)] = '=F{}'.format(row, row)
                ws['FA{}'.format(row)] = '=IF(G{}="","",G{})'.format(row, row)
                ws['FB{}'.format(row)] = '=IF(H{}="","",H{})'.format(row, row)
                ws['FC{}'.format(row)] = '=IF(I{}="","",I{})'.format(row, row)
                ws['FD{}'.format(row)] = '=IF(J{}="","",J{})'.format(row, row)
                ws['FE{}'.format(row)] = '=IF(K{}="","",K{})'.format(row, row)
                ws['FF{}'.format(row)] = '=IF(L{}="","",L{})'.format(row, row)
                ws['FG{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FA{}="","",(FA{}-FA${})/FA${}),2),"")'.format(row, row, r, s)
                ws['FH{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FB{}="","",(FB{}-FB${})/FB${}),2),"")'.format(row, row, r, s)
                ws['FI{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FC{}="","",(FC{}-FC${})/FC${}),2),"")'.format(row, row, r, s)
                ws['FJ{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FD{}="","",(FD{}-FD${})/FD${}),2),"")'.format(row, row, r, s)
                ws['FK{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FE{}="","",(FE{}-FE${})/FE${}),2),"")'.format(row, row, r, s)
                ws['FL{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FG{}="","",IF(70+30*FG{}/$FG${}<20,20,70+30*FG{}/$FG${})),2),"")'.format(row, row, r, row, r)
                ws['FM{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FH{}="","",IF(70+30*FH{}/$FH${}<20,20,70+30*FH{}/$FH${})),2),"")'.format(row, row, r, row, r)
                ws['FN{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FI{}="","",IF(70+30*FI{}/$FI${}<20,20,70+30*FI{}/$FI${})),2),"")'.format(row, row, r, row, r)
                ws['FO{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FJ{}="","",IF(70+30*FJ{}/$FJ${}<20,20,70+30*FJ{}/$FJ${})),2),"")'.format(row, row, r, row, r)
                ws['FP{}'.format(
                    row)] = '=IFERROR(ROUND(IF(FK{}="","",IF(70+30*FK{}/$FK${}<20,20,70+30*FK{}/$FK${})),2),"")'.format(row, row, r, row, r)

                ws['FQ{}'.format(row)] = '=IF(SUM(FL{}:FP{})=0,"",SUM(FL{}:FP{}))'.format(
                    row, row, row, row)
                ws['FR{}'.format(row)] = '=IF(FQ{}="","",RANK(FQ{},$FQ$2:$FQ${}))'.format(
                    row, row, q)
                ws['FS{}'.format(
                    row)] = '=IF(FR{}="","",COUNTIFS($EZ$2:$EZ${},F{},$FR$2:$FR${},"<"&FR{})+1)'.format(row, q, row, q, row)
                #     TAMBAHAN
                ws['FT{}'.format(row)] = '=IF($G${}=25,IF(AND(FA{}>4,FL{}=20),1,""),IF($G${}=30,IF(AND(FA{}>5,FL{}=20),1,""),IF($G${}=35,IF(AND(FA{}>6,FL{}=20),1,""),IF($G${}=40,IF(AND(FA{}>7,FL{}=20),1,""),IF($G${}=45,IF(AND(FA{}>8,FL{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['FU{}'.format(row)] = '=IF($H${}=25,IF(AND(FB{}>4,FM{}=20),1,""),IF($H${}=30,IF(AND(FB{}>5,FM{}=20),1,""),IF($H${}=35,IF(AND(FB{}>6,FM{}=20),1,""),IF($H${}=40,IF(AND(FB{}>7,FM{}=20),1,""),IF($H${}=45,IF(AND(FB{}>8,FM{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['FV{}'.format(row)] = '=IF($I${}=25,IF(AND(FC{}>4,FN{}=20),1,""),IF($I${}=30,IF(AND(FC{}>5,FN{}=20),1,""),IF($I${}=35,IF(AND(FC{}>6,FN{}=20),1,""),IF($I${}=40,IF(AND(FC{}>7,FN{}=20),1,""),IF($I${}=45,IF(AND(FC{}>8,FN{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['FW{}'.format(row)] = '=IF($J${}=25,IF(AND(FD{}>4,FO{}=20),1,""),IF($J${}=30,IF(AND(FD{}>5,FO{}=20),1,""),IF($J${}=35,IF(AND(FD{}>6,FO{}=20),1,""),IF($J${}=40,IF(AND(FD{}>7,FO{}=20),1,""),IF($J${}=45,IF(AND(FD{}>8,FO{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)
                ws['FX{}'.format(row)] = '=IF($K${}=25,IF(AND(FE{}>4,FP{}=20),1,""),IF($K${}=30,IF(AND(FE{}>5,FP{}=20),1,""),IF($K${}=35,IF(AND(FE{}>6,FP{}=20),1,""),IF($K${}=40,IF(AND(FE{}>7,FP{}=20),1,""),IF($K${}=45,IF(AND(FE{}>8,FP{}=20),1,""))))))'.format(
                    v, row, row, v, row, row, v, row, row, v, row, row, v, row, row)

            # Mengubah 'KELAS' sesuai dengan nilai yang dipilih dari selectbox 'KELAS'
            kelas = KELAS.lower().replace(" ", "")
            semester = SEMESTER.lower()
            tahun = TAHUN.replace("-", "")
            penilaian = PENILAIAN.lower()
            kurikulum = KURIKULUM.lower()

            path_file = f"{kelas}_{penilaian}_{semester}_{kurikulum}_{tahun}_nilai_std.xlsx"

            # Simpan file ke direktori temporer
            temp_dir = tempfile.gettempdir()
            file_path = temp_dir + '/' + path_file
            wb.save(file_path)

            st.success(
                "File siap diunduh!")

            # Tombol unduh file
            with open(file_path, "rb") as f:
                bytes_data = f.read()
            st.download_button(label="Unduh File", data=bytes_data,
                               file_name=path_file)

            st.warning(
                "Buka file unduhan, klik 'Enable Editing' dan 'Save'")  