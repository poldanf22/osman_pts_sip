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
        k13_4sd_mat = 'M4d1O{toUmum_tahun}K13'
        k13_4sd_ind = 'I4d1O{toUmum_tahun}K13'
        k13_4sd_eng = 'E4d1O{toUmum_tahun}K13'
        k13_4sd_ipa = 'A4d1O{toUmum_tahun}K13'
        k13_4sd_ips = 'Z4d1O{toUmum_tahun}K13'
        k13_4sd = [k13_4sd_mat, k13_4sd_ind,
                   k13_4sd_eng, k13_4sd_ipa, k13_4sd_ips]
        column_order_k13_4sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_4SD', 'IND_4SD',
                                'ENG_4SD', 'IPA_4SD', 'IPS_4SD']

        # 5sd k13
        k13_5sd_mat = 'M5d1O{toUmum_tahun}K13'
        k13_5sd_ind = 'I5d1O{toUmum_tahun}K13'
        k13_5sd_eng = 'E5d1O{toUmum_tahun}K13'
        k13_5sd_ipa = 'A5d1O{toUmum_tahun}K13'
        k13_5sd_ips = 'Z5d1O{toUmum_tahun}K13'
        k13_5sd = [k13_5sd_mat, k13_5sd_ind,
                   k13_5sd_eng, k13_5sd_ipa, k13_5sd_ips]
        column_order_k13_5sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_5SD', 'IND_5SD',
                                'ENG_5SD', 'IPA_5SD', 'IPS_5SD']

        # 6sd k13
        k13_6sd_mat = 'M6d1O{toUmum_tahun}K13'
        k13_6sd_ind = 'I6d1O{toUmum_tahun}K13'
        k13_6sd_eng = 'E6d1O{toUmum_tahun}K13'
        k13_6sd_ipa = 'A6d1O{toUmum_tahun}K13'
        k13_6sd_ips = 'Z6d1O{toUmum_tahun}K13'
        k13_6sd = [k13_6sd_mat, k13_6sd_ind,
                   k13_6sd_eng, k13_6sd_ipa, k13_6sd_ips]
        column_order_k13_6sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_6SD', 'IND_6SD',
                                'ENG_6SD', 'IPA_6SD', 'IPS_6SD']

        # 7smp k13
        k13_7smp_mat = 'M1p1O{toUmum_tahun}K13'
        k13_7smp_ind = 'I1p1O{toUmum_tahun}K13'
        k13_7smp_eng = 'E1p1O{toUmum_tahun}K13'
        k13_7smp_ipa = '4161A1{tahun}'
        k13_7smp_ips = 'G1p1O{toUmum_tahun}K13'
        k13_7smp = [k13_7smp_mat, k13_7smp_ind,
                    k13_7smp_eng, k13_7smp_ipa, k13_7smp_ips]
        column_order_k13_7smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_7SMP', 'IND_7SMP',
                                 'ENG_7SMP', 'IPA_7SMP', 'IPS_7SMP']

        # 8smp k13
        k13_8smp_mat = 'M2p1O{toUmum_tahun}K13'
        k13_8smp_ind = 'I2p1O{toUmum_tahun}K13'
        k13_8smp_eng = 'E2p1O{toUmum_tahun}K13'
        k13_8smp_ipa = '5161A1{tahun}'
        k13_8smp_ips = 'G1p1O{toUmum_tahun}K13'
        k13_8smp_mat_new = 'M2p1O{toUnik_tahun}K13'
        k13_8smp = [k13_8smp_mat, k13_8smp_ind,
                    k13_8smp_eng, k13_8smp_ipa, k13_8smp_ips]
        column_order_k13_8smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_8SMP', 'IND_8SMP',
                                 'ENG_8SMP', 'IPA_8SMP', 'IPS_8SMP', 'MAT_NEW_8SMP']

        # 9smp k13
        k13_9smp_mat = 'M3p1O{toUmum_tahun}K13'
        k13_9smp_ind = 'I3p1O{toUmum_tahun}K13'
        k13_9smp_eng = 'E3p1O{toUmum_tahun}K13'
        k13_9smp_ipa = '6161A123-24'
        k13_9smp_ips = 'G3p1O{toUmum_tahun}K13'
        k13_9smp = [k13_9smp_mat, k13_9smp_ind,
                    k13_9smp_eng, k13_9smp_ipa, k13_9smp_ips]
        column_order_k13_9smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_9SMP', 'IND_9SMP',
                                 'ENG_9SMP', 'IPA_9SMP', 'IPS_9SMP']

        image = Image.open('logo resmi nf resize.png')
        st.image(image)

        st.title("PIVOT - PTS")

        col1 = st.container()
        with col1:
            KURIKULUM = st.selectbox(
                "KURIKULUM",
                ("--Pilih Kurikulum--", "K13", "KM"))

        col2 = st.container()
        with col2:
            KELAS = st.selectbox(
                "KELAS",
                ("--Pilih Kelas--", "4 SD", "5 SD", "6 SD", "7 SMP", "8 SMP", "9 SMP"))

        col3 = st.container()
        with col3:
            SEMESTER = st.selectbox(
                "SEMESTER",
                ("--Pilih Semester--", "SEMESTER 1", "SEMESTER 2"))

        TAHUN = st.text_input("Masukkan Tahun Ajaran",
                              placeholder="contoh: 2022-2023")
