import pandas as pd
import streamlit as st
import pickle
from pathlib import Path
import plotly.express as px
from PIL import Image
import matplotlib.pyplot as plt
import streamlit_authenticator as stauth
from docxtpl import DocxTemplate
import base64
from io import BytesIO
from docx import Document
import io

names = ['Salvador Jair Ocampo','Rodrigo Manzano']
usernames = ['jocampo', 'rmanzano']

file_path = Path(__file__).parent / 'hashed_pw.pkl'
with file_path.open('rb') as file:
    hashed_passwords = pickle.load(file)

authenticator = stauth.Authenticate(names, usernames, hashed_passwords,
                                    'principal_dashboard', 'abcdef', cookie_expiry_days=2/24)

name, authentication_status, username = authenticator.login('Login','main')

if authentication_status == False:
    st.error('El Usuario/Constraseña es incorrecta')

if authentication_status == None:
    st.warning('Complete todo los campos')

if authentication_status == True:
    st.sidebar.title(f'Bienvenido {name}')
    authenticator.logout('Logout', 'sidebar')
    imagen = "foto4.jpg"
    imagen2 = 'Foto2.jpg'
    st.markdown(
    f'<div style="display: flex; justify-content: space-between;">'
    f'    <img src="data:image2/png;base64,{base64.b64encode(open(imagen2, "rb").read()).decode()}" '
    f'        style="float: left; margin-right: 10px; margin-top: 10px;" />'
    f'    <img src="data:image/png;base64,{base64.b64encode(open(imagen, "rb").read()).decode()}" '
    f'        style="float: right; margin-right: 10px; margin-top: 10px;" />'
    f'</div>',
    unsafe_allow_html=True
    )

    st.markdown("<h1 style='text-align: center;'>Seguimiento De Estudiantes en RC y E</h1>", unsafe_allow_html=True)

    st.header('Estadisticos del Excel')
    #excel_file = 'PORCENTAJES.xlsx'
    #df = pd.read_excel(excel_file, header=4)
    st.subheader('Subir archivo Excel')
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx", "xls","csv"])

    # Si se carga un archivo, cargarlo en un DataFrame de Pandas
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file,header=4)
    else:
        # Si no se carga un archivo, usar un DataFrame vacío
        df = pd.DataFrame()

    if not df.empty:    
        st.markdown('---')

        df = df.fillna(0)
        df = df.dropna()

        st.dataframe(df)

        st.markdown('---')
        count_columna1 = (df['AC'] == 'X').sum()
        count_columna2 = (df['NA'] == 'X').sum()
        # Crear una gráfica de barras con ambas columnas
        fig, ax = plt.subplots(figsize=(10, 5))
        bar_labels = ['ACREDITADOS', 'NO ACREDITADOS']
        bar_counts = [count_columna1, count_columna2]
        bars = ax.bar(bar_labels, bar_counts, color=['blue', 'green'])
        ax.set_ylabel('Estudiantes')
        ax.set_title('RESULTADOS ACREDITACIONES DE ALUMNOS ASESORADOS')
        # Agregar el número total dentro de cada barra
        for bar, count in zip(bars, bar_counts):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), count, ha='center', va='bottom')
        plt.tight_layout()
        st.pyplot(plt)

        st.markdown('---')
        columnas = ['MATERIA', 'AC', 'NA','Hrs. Asesoría']
        df = df[columnas]
        df['NA'] = df['NA'].eq('X').groupby(df['MATERIA']).transform('sum')
        df['AC'] = df['AC'].eq('X').groupby(df['MATERIA']).transform('sum')
        df['Hrs. Impartidas'] = df['Hrs. Asesoría'].groupby(df['MATERIA']).transform('sum')
        df['Estudiantes'] = df['NA'] + df['AC']
        df = df.drop_duplicates(subset=['MATERIA'])
        df = df[['MATERIA','AC', 'NA', 'Hrs. Impartidas','Estudiantes']]
        st.dataframe(df)


        if st.button('Descargar Resultados en Excel'):
            with st.spinner('Generando archivo, por favor espera...'):
                excel_buffer = io.BytesIO()
                df.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)
                excel_base64 = base64.b64encode(excel_buffer.read()).decode()
                download_link_excel = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_base64}" download="resultados.xlsx">Descargar Excel</a>'
                st.markdown(download_link_excel, unsafe_allow_html=True)
            st.success('Archivo generado exitosamente.')


        st.markdown('---')
        fig, ax = plt.subplots(figsize=(12, 8))
        materias = df['MATERIA']
        na_values = df['NA']
        ac_values = df['AC']
        bar_width = 0.35
        index = range(len(materias))
        bar1 = ax.bar([i - bar_width / 2 for i in index], na_values, bar_width, label='NA', color='blue')
        bar2 = ax.bar([i + bar_width / 2 for i in index], ac_values, bar_width, label='AC', color='green')
        ax.set_xlabel('Asignaturas')
        ax.set_ylabel('Estudiantes')
        ax.set_title('ACREDITACIONES DE ALUMNOS ASESORADOS POR MATERIA')
        ax.set_xticks(index)
        ax.set_xticklabels(materias, rotation=45, ha='right')
        ax.legend()
        for i, j in zip(index, na_values):
            ax.text(i - bar_width / 2, j, str(j), ha='center', va='bottom')
        for i, j in zip(index, ac_values):
            ax.text(i + bar_width / 2, j, str(j), ha='center', va='bottom')
        plt.tight_layout()
        st.pyplot(plt)

        st.markdown('---')
        ig, ax = plt.subplots(figsize=(12, 8))
        materias = df['MATERIA']
        na_values = df['Hrs. Impartidas']

        bar_width = 0.35
        index = range(len(materias))
        bar1 = ax.bar([i - bar_width / 2 for i in index], na_values, bar_width, label='Hrs. Impartidas', color='blue')
        ax.set_xlabel('Asignaturas')
        ax.set_ylabel('Estudiantes')
        ax.set_title('Hrs. Asesoría impartidas por materia.(Total de horas de asesoría: 42 hrs.)')
        ax.set_xticks(index)
        ax.set_xticklabels(materias, rotation=45, ha='right')
        ax.legend()
        for i, j in zip(index, na_values):
            ax.text(i - bar_width / 2, j, str(j), ha='center', va='bottom')
        plt.tight_layout()
        st.pyplot(plt)
