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
