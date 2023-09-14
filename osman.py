import pickle
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
from PIL import Image
import pandas as pd
from streamlit_option_menu import option_menu
import openpyxl
from openpyxl.styles import Font, PatternFill
import tempfile

# User Authentication
names = ["TI Polda NF 1", "TI Polda NF 2"]
usernames = ["admin1", "admin2"]

# load hashed kd_akses
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_kd_akses = pickle.load(file)

authenticator = stauth.Authenticate(
    names, usernames, hashed_kd_akses, "lookup", "abcdef")
name, authentication_status, username = authenticator.login("Login", "main")

if authentication_status == False:
    st.error("Username/kode akses salah")

if authentication_status == None:
    st.warning("Silahkan masukan username dan kode akses")

url = "https://osman2-8bdgvgq3z54.streamlit.app/"

if authentication_status:
    authenticator.logout("Logout", "sidebar")
    with st.sidebar:
        with st.sidebar:
            st.markdown(
                f'''<a href={url}><button style="background-color:GreenYellow;">Untuk Lok.</button></a>''', unsafe_allow_html=True)
        selected_file = option_menu(
            menu_title="Pilih file:",
            options=["Pivot PTS", "Nilai Std. SD, SMP"],
        )
    toUmum_tahun = "0123-24"
    toUnik_tahun = "0323-24"
    tahun = "23-24"
    if selected_file == "Pivot PTS":
        # kurikulum - kelas - mapel
        # 4sd k13
        k13_4sd_mat = 'M4d1O'+toUmum_tahun+'K13'
        k13_4sd_ind = 'I4d1O'+toUmum_tahun+'K13'
        k13_4sd_eng = 'E4d1O'+toUmum_tahun+'K13'
        k13_4sd_ipa = 'A4d1O'+toUmum_tahun+'K13'
        k13_4sd_ips = 'Z4d1O'+toUmum_tahun+'K13'
        k13_4sd = [k13_4sd_mat, k13_4sd_ind,
                   k13_4sd_eng, k13_4sd_ipa, k13_4sd_ips]
        column_order_k13_4sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_4SD', 'IND_4SD',
                                'ENG_4SD', 'IPA_4SD', 'IPS_4SD']

        # 5sd k13
        k13_5sd_mat = 'M5d1O'+toUmum_tahun+'K13'
        k13_5sd_ind = 'I5d1O'+toUmum_tahun+'K13'
        k13_5sd_eng = 'E5d1O'+toUmum_tahun+'K13'
        k13_5sd_ipa = 'A5d1O'+toUmum_tahun+'K13'
        k13_5sd_ips = 'Z5d1O'+toUmum_tahun+'K13'
        k13_5sd = [k13_5sd_mat, k13_5sd_ind,
                   k13_5sd_eng, k13_5sd_ipa, k13_5sd_ips]
        column_order_k13_5sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_5SD', 'IND_5SD',
                                'ENG_5SD', 'IPA_5SD', 'IPS_5SD']

        # 6sd k13
        k13_6sd_mat = 'M6d1O'+toUmum_tahun+'K13'
        k13_6sd_ind = 'I6d1O'+toUmum_tahun+'K13'
        k13_6sd_eng = 'E6d1O'+toUmum_tahun+'K13'
        k13_6sd_ipa = 'A6d1O'+toUmum_tahun+'K13'
        k13_6sd_ips = 'Z6d1O'+toUmum_tahun+'K13'
        k13_6sd = [k13_6sd_mat, k13_6sd_ind,
                   k13_6sd_eng, k13_6sd_ipa, k13_6sd_ips]
        column_order_k13_6sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_6SD', 'IND_6SD',
                                'ENG_6SD', 'IPA_6SD', 'IPS_6SD']

        # 7smp k13
        k13_7smp_mat = 'M1p1O'+toUmum_tahun+'K13'
        k13_7smp_ind = 'I1p1O'+toUmum_tahun+'K13'
        k13_7smp_eng = 'E1p1O'+toUmum_tahun+'K13'
        k13_7smp_ipa = '4161A1'+tahun
        k13_7smp_ips = 'G1p1O'+toUmum_tahun+'K13'
        k13_7smp = [k13_7smp_mat, k13_7smp_ind,
                    k13_7smp_eng, k13_7smp_ipa, k13_7smp_ips]
        column_order_k13_7smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_7SMP', 'IND_7SMP',
                                 'ENG_7SMP', 'IPA_7SMP', 'IPS_7SMP']

        # 8smp k13
        k13_8smp_mat = 'M2p1O'+toUmum_tahun+'K13'
        k13_8smp_ind = 'I2p1O'+toUmum_tahun+'K13'
        k13_8smp_eng = 'E2p1O'+toUmum_tahun+'K13'
        k13_8smp_ipa = '5161A1'+tahun
        k13_8smp_ips = 'G2p1O'+toUmum_tahun+'K13'
        k13_8smp = [k13_8smp_mat, k13_8smp_ind,
                    k13_8smp_eng, k13_8smp_ipa, k13_8smp_ips]
        column_order_k13_8smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_8SMP', 'IND_8SMP',
                                 'ENG_8SMP', 'IPA_8SMP', 'IPS_8SMP']

        # 9smp k13
        k13_9smp_mat = 'M3p1O'+toUmum_tahun+'K13'
        k13_9smp_ind = 'I3p1O'+toUmum_tahun+'K13'
        k13_9smp_eng = 'E3p1O'+toUmum_tahun+'K13'
        k13_9smp_ipa = '6161A1'+tahun
        k13_9smp_ips = 'G3p1O'+toUmum_tahun+'K13'
        k13_9smp = [k13_9smp_mat, k13_9smp_ind,
                    k13_9smp_eng, k13_9smp_ipa, k13_9smp_ips]
        column_order_k13_9smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_9SMP', 'IND_9SMP',
                                 'ENG_9SMP', 'IPA_9SMP', 'IPS_9SMP']

        # PPLS IPA
        ppls_ipa_mat = 'M9a1O'+toUmum_tahun+'PPLS'
        ppls_ipa_fis = 'F9a1O'+toUmum_tahun+'PPLS'
        ppls_ipa_kim = 'K9a1O'+toUmum_tahun+'PPLS'
        ppls_ipa_bio = 'B9a1O'+toUmum_tahun+'PPLS'
        ppls_ipa = [ppls_ipa_mat, ppls_ipa_bio,
                    ppls_ipa_fis, ppls_ipa_kim]
        column_order_ppls_ipa = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_PPLS_IPA',
                                 'FIS_PPLS_IPA', 'KIM_PPLS_IPA', 'BIO_PPLS_IPA',]

        # PPLS IPS
        ppls_ips_geo = 'G9s1O'+toUmum_tahun+'PPLS'
        ppls_ips_eko = 'O9s1O'+toUmum_tahun+'PPLS'
        ppls_ips_sej = 'S9s1O'+toUmum_tahun+'PPLS'
        ppls_ips_sos = 'L9s1O'+toUmum_tahun+'PPLS'
        ppls_ips = [ppls_ips_geo, ppls_ips_eko,
                    ppls_ips_sej, ppls_ips_sos]
        column_order_ppls_ips = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'GEO_PPLS_IPS',
                                 'EKO_PPLS_IPS', 'SEJ_PPLS_IPS', 'SOS_PPLS_IPS',]

        # 4sd km
        km_4sd_mat = 'M4d1O'+toUmum_tahun+'KM'
        km_4sd_ind = 'I4d1O'+toUmum_tahun+'KM'
        km_4sd_eng = 'E4d1O'+toUmum_tahun+'KM'
        km_4sd_ipas = '1281D1'+tahun
        km_4sd = [km_4sd_mat, km_4sd_ind,
                  km_4sd_eng, km_4sd_ipas]
        column_order_km_4sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_4SD', 'IND_4SD',
                               'ENG_4SD', 'IPAS_4SD']

        # 5sd km
        km_5sd_mat = 'M5d1O'+toUmum_tahun+'KM'
        km_5sd_ind = 'I5d1O'+toUmum_tahun+'KM'
        km_5sd_eng = 'E5d1O'+toUmum_tahun+'KM'
        km_5sd_ipas = '2281D123-24'
        km_5sd = [km_5sd_mat, km_5sd_ind,
                  km_5sd_eng, km_5sd_ipas]
        column_order_km_5sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_5SD', 'IND_5SD',
                               'ENG_5SD', 'IPAS_5SD']

        # 7smp km
        km_7smp_mat = 'M1p1O'+toUmum_tahun+'KM'
        km_7smp_ind = 'I1p1O'+toUmum_tahun+'KM'
        km_7smp_eng = 'E1p1O'+toUmum_tahun+'KM'
        km_7smp_ipa = '4281A1'+tahun
        km_7smp_ips = '4281S1'+tahun
        km_7smp = [km_7smp_mat, km_7smp_ind,
                   km_7smp_eng, km_7smp_ipa, km_7smp_ips]
        column_order_km_7smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_7SMP', 'IND_7SMP',
                                'ENG_7SMP', 'IPA_7SMP', 'IPS_7SMP']

        # 8smp km
        km_8smp_mat = 'M2p1O'+toUmum_tahun+'KM'
        km_8smp_ind = 'I2p1O'+toUmum_tahun+'KM'
        km_8smp_eng = 'E2p1O'+toUmum_tahun+'KM'
        km_8smp_ipa = 'B2p1O'+toUmum_tahun+'KM'
        km_8smp_ips = '5281S1'+tahun
        km_8smp_mat_new = 'M2p1O'+toUnik_tahun+'KM'
        km_8smp = [km_8smp_mat, km_8smp_ind,
                   km_8smp_eng, km_8smp_ipa, km_8smp_ips, km_8smp_mat_new]
        column_order_km_8smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_8SMP', 'IND_8SMP',
                                'ENG_8SMP', 'IPA_8SMP', 'IPS_8SMP', 'MAT_NEW_8SMP']

        image = Image.open('logo resmi nf resize.png')
        st.image(image)

        st.title("PIVOT - PTS")

        col1 = st.container()
        with col1:
            KURIKULUM = st.selectbox(
                "KURIKULUM",
                ("--Pilih Kurikulum--", "K13", "KM", "PPLS"))

        col2 = st.container()
        with col2:
            KELAS = st.selectbox(
                "KELAS",
                ("--Pilih Kelas--", "4 SD", "5 SD", "6 SD", "7 SMP", "8 SMP", "9 SMP", "PPLS IPA", "PPLS IPS"))

        col3 = st.container()
        with col3:
            SEMESTER = st.selectbox(
                "SEMESTER",
                ("--Pilih Semester--", "SEMESTER 1", "SEMESTER 2"))

        TAHUN = st.text_input("Masukkan Tahun Ajaran",
                              placeholder="contoh: 2022-2023")

        uploaded_detail = st.file_uploader(
            'Letakkan file excel Detail Siswa', type='xlsx')
        uploaded_to_pts = st.file_uploader(
            'Letakkan file excel TO PTS', type='xlsx')

        detail = None
        to_pts = None

        if uploaded_detail is not None:
            detail = pd.read_excel(uploaded_detail)

        if uploaded_to_pts is not None:
            to_pts = pd.read_excel(uploaded_to_pts)

        if detail is not None and to_pts is not None:
            detail = detail.drop(['user_id', 'is_test_access', 'no_hp', 'lokasi_id', 'jenjang_id',
                                  'riwayat_jenjang', 'jenjang_dipilih_id', 'kode_level', 'kode_kelas',
                                  'tempat_lahir', 'tanggal_lahir', 'semester', 'tahun_ajar',
                                  'program', 'pin', 'join_skolla', 'created_at', 'updated_at'], axis=1)  # Menghilangkan kolom sebelum dilakukan merge

            result = pd.merge(detail, to_pts[['no_nf', 'kode_paket', 'tahun_ajaran', 'kelas_id',
                                              'lokasi_id', 'jumlah_benar']], on='no_nf', how='left')
            # Menghapus nilai NaN dari kolom 'kode_paket'
            result = result.dropna(subset=['kode_paket'])

            # k13
            if KELAS == "4 SD" and KURIKULUM == "K13":
                kode_kls_kur = k13_4sd
                column_order = column_order_k13_4sd
            elif KELAS == "5 SD" and KURIKULUM == "K13":
                kode_kls_kur = k13_5sd
                column_order = column_order_k13_5sd
            elif KELAS == "6 SD" and KURIKULUM == "K13":
                kode_kls_kur = k13_6sd
                column_order = column_order_k13_6sd
            elif KELAS == "7 SMP" and KURIKULUM == "K13":
                kode_kls_kur = k13_7smp
                column_order = column_order_k13_7smp
            elif KELAS == "8 SMP" and KURIKULUM == "K13":
                kode_kls_kur = k13_8smp
                column_order = column_order_k13_8smp
            elif KELAS == "9 SMP" and KURIKULUM == "K13":
                kode_kls_kur = k13_9smp
                column_order = column_order_k13_9smp
            # km
            elif KELAS == "4 SD" and KURIKULUM == "KM":
                kode_kls_kur = km_4sd
                column_order = column_order_km_4sd
            elif KELAS == "5 SD" and KURIKULUM == "KM":
                kode_kls_kur = km_5sd
                column_order = column_order_km_5sd
            elif KELAS == "7 SMP" and KURIKULUM == "KM":
                kode_kls_kur = km_7smp
                column_order = column_order_km_7smp
            elif KELAS == "8 SMP" and KURIKULUM == "KM":
                kode_kls_kur = km_8smp
                column_order = column_order_km_8smp
            # ppls
            elif KELAS == "PPLS IPA" and KURIKULUM == "PPLS":
                kode_kls_kur = ppls_ipa
                column_order = column_order_ppls_ipa
            elif KELAS == "PPLS IPS" and KURIKULUM == "PPLS":
                kode_kls_kur = ppls_ips
                column_order = column_order_ppls_ips

            result_filtered = result[result['kode_paket'].isin(kode_kls_kur)]
            result_filtered.drop_duplicates(
                subset=['name', 'kode_paket'], keep='first', inplace=True)

            # Menggunakan pivot_table untuk menjadikan konten kolom 'kode_paket' sebagai header dan menghilangkan duplikat
            result_pivot = pd.pivot_table(result_filtered, index=[
                'name', 'no_nf', 'lokasi_id', 'sekolah', 'kelas_id', 'tahun_ajaran'], columns='kode_paket', values='jumlah_benar', aggfunc='first')
            result_pivot.reset_index(inplace=True)  # Mengatur ulang indeks

            # Ubah nama kolom
            result_pivot = result_pivot.rename(
                columns={'name': 'NAMA', 'no_nf': 'NONF', 'lokasi_id': 'KD_LOK', 'sekolah': 'NAMA_SKLH', 'kelas_id': 'KELAS', 'tahun_ajaran': 'IDTAHUN',
                         'M4d1O'+toUmum_tahun+'K13': 'MAT_4SD', 'I4d1O'+toUmum_tahun+'K13': 'IND_4SD', 'E4d1O'+toUmum_tahun+'K13': 'ENG_4SD', 'A4d1O'+toUmum_tahun+'K13': 'IPA_4SD', 'Z4d1O'+toUmum_tahun+'K13': 'IPS_4SD',
                         'M5d1O'+toUmum_tahun+'K13': 'MAT_5SD', 'I5d1O'+toUmum_tahun+'K13': 'IND_5SD', 'E5d1O'+toUmum_tahun+'K13': 'ENG_5SD', 'A5d1O'+toUmum_tahun+'K13': 'IPA_5SD', 'Z5d1O'+toUmum_tahun+'K13': 'IPS_5SD',
                         'M6d1O'+toUmum_tahun+'K13': 'MAT_6SD', 'I6d1O'+toUmum_tahun+'K13': 'IND_6SD', 'E6d1O'+toUmum_tahun+'K13': 'ENG_6SD', 'A6d1O'+toUmum_tahun+'K13': 'IPA_6SD', 'Z6d1O'+toUmum_tahun+'K13': 'IPS_6SD',
                         'M1p1O'+toUmum_tahun+'K13': 'MAT_7SMP', 'I1p1O'+toUmum_tahun+'K13': 'IND_7SMP', 'E1p1O'+toUmum_tahun+'K13': 'ENG_7SMP', '4161A1'+tahun: 'IPA_7SMP', 'G1p1O'+toUmum_tahun+'K13': 'IPS_7SMP',
                         'M2p1O'+toUmum_tahun+'K13': 'MAT_8SMP', 'I2p1O'+toUmum_tahun+'K13': 'IND_8SMP', 'E2p1O'+toUmum_tahun+'K13': 'ENG_8SMP', '5161A1'+tahun: 'IPA_8SMP', 'G2p1O'+toUmum_tahun+'K13': 'IPS_8SMP',
                         'M3p1O'+toUmum_tahun+'K13': 'MAT_9SMP', 'I3p1O'+toUmum_tahun+'K13': 'IND_9SMP', 'E3p1O'+toUmum_tahun+'K13': 'ENG_9SMP', '6161A1'+tahun: 'IPA_9SMP', 'G3p1O'+toUmum_tahun+'K13': 'IPS_9SMP',
                         'M4d1O'+toUmum_tahun+'KM': 'MAT_4SD', 'I4d1O'+toUmum_tahun+'KM': 'IND_4SD', 'E4d1O'+toUmum_tahun+'KM': 'ENG_4SD', '1281D1'+tahun: 'IPAS_4SD',
                         'M5d1O'+toUmum_tahun+'KM': 'MAT_5SD', 'I5d1O'+toUmum_tahun+'KM': 'IND_5SD', 'E5d1O'+toUmum_tahun+'KM': 'ENG_5SD', '2281D1'+tahun: 'IPAS_5SD',
                         'M1p1O'+toUmum_tahun+'KM': 'MAT_7SMP', 'I1p1O'+toUmum_tahun+'KM': 'IND_7SMP', 'E1p1O'+toUmum_tahun+'KM': 'ENG_7SMP', '4281A1'+tahun: 'IPA_7SMP', '4281S1'+tahun: 'IPS_7SMP',
                         'M2p1O'+toUmum_tahun+'KM': 'MAT_8SMP', 'I2p1O'+toUmum_tahun+'KM': 'IND_8SMP', 'E2p1O'+toUmum_tahun+'KM': 'ENG_8SMP', 'B2p1O'+toUmum_tahun+'KM': 'IPA_8SMP', '5281S1'+tahun: 'IPS_8SMP', 'M2p1O'+toUnik_tahun+'KM': 'MAT_NEW_8SMP',
                         'M9a1O'+toUmum_tahun+'PPLS': 'MAT_PPLS_IPA', 'F9a1O'+toUmum_tahun+'PPLS': 'FIS_PPLS_IPA', 'K9a1O'+toUmum_tahun+'PPLS': 'KIM_PPLS_IPA', 'B9a1O'+toUmum_tahun+'PPLS': 'BIO_PPLS_IPA',
                         'G9s1O'+toUmum_tahun+'PPLS': 'GEO_PPLS_IPS', 'O9s1O'+toUmum_tahun+'PPLS': 'EKO_PPLS_IPS', 'S9s1O'+toUmum_tahun+'PPLS': 'SEJ_PPLS_IPS', 'L9s1O'+toUmum_tahun+'PPLS': 'SOS_PPLS_IPS'})

            result_pivot = result_pivot.reindex(columns=column_order)

            kelas = KELAS.lower().replace(" ", "")
            kurikulum = KURIKULUM.lower()
            tahun = TAHUN.replace("-", "")
            semester = SEMESTER.lower()

            path_file = f"{kelas}_pts_{semester}_{kurikulum}_{tahun}_pivot.xlsx"

            # Simpan file ke direktori temporer
            temp_dir = tempfile.gettempdir()
            file_path = temp_dir + '/' + path_file
            # wb.save(file_path)

            # Menyimpan DataFrame ke file Excel
            result_pivot.to_excel(file_path, index=False)
            st.success("File siap diunduh!")

            # Tombol unduh file
            with open(file_path, "rb") as f:
                bytes_data = f.read()
            st.download_button(label="Unduh File", data=bytes_data,
                               file_name=path_file)
