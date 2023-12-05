import pickle
from pathlib import Path
import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import os
import io
import datetime
from docx import Document
from docxtpl import DocxTemplate
import base64
import zipfile
#from docx2pdf import convert
#import win32com.client

#win32com.client.Dispatch("WScript.Shell")
fecha_actual = datetime.datetime.now().strftime("%d/%m/%Y")
st.set_page_config(page_title='Sistema', page_icon='🌍', layout='wide')
names = ['Salvador Jair Ocampo','Rodrigo Manzano']
usernames = ['jocampo', 'rmanzano']
file_path = Path(__file__).parent / 'hashed_pw.pkl'
with file_path.open('rb') as file:
    hashed_passwords = pickle.load(file)

authenticator = stauth.Authenticate(names, usernames, hashed_passwords,
                                    'principal_dashboard', 'abcdef', cookie_expiry_days=2/24)  # 2 horas


name, authentication_status, username = authenticator.login('Login','main')

if authentication_status == False:
    st.error('El Usuario/Constraseña es incorrecta')

if authentication_status == None:
    st.warning('Complete todo los campos')

if authentication_status:    
    imagen = "Foto4.jpg"
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

    #st.sidebar.success("Selecciona la Opcion Arriba")
    st.sidebar.title(f'Bienvenido {name}')
    authenticator.logout('Logout', 'sidebar')
    
    st.subheader('Actualizar archivo de Docentes')
    new_word_file = st.file_uploader("Cargar o Reemplazar Archivo Word Docentes ", type=["docx"])

    if new_word_file is not None:
        # Procesar el nuevo archivo Word cargado y reemplazar el archivo existente
        with st.spinner('Procesando el nuevo archivo, por favor espera...'):
            # Leer el nuevo archivo Word
            new_doc = Document(new_word_file)

            new_file_path = "docentes.docx"
            new_doc.save(new_file_path)
            
            st.success('Archivo Word actualizado exitosamente.')

    st.subheader('Actualizar archivo de Tutores')
    new_word_file = st.file_uploader("Cargar o Reemplazar Archivo Word Tutores ", type=["docx"])

    if new_word_file is not None:
        # Procesar el nuevo archivo Word cargado y reemplazar el archivo existente
        with st.spinner('Procesando el nuevo archivo, por favor espera...'):
            # Leer el nuevo archivo Word
            new_doc = Document(new_word_file)

            new_file_path = "tutorados.docx"
            new_doc.save(new_file_path)
            
            st.success('Archivo Word actualizado exitosamente.')

    st.header('Filtrado del Excel')
    st.subheader('Subir archivo Excel')
    # carga de excel
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx", "xls","csv"])

    # Si se carga un archivo, cargarlo en un DataFrame de Pandas
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file).dropna()
    else:
        # Si no se carga un archivo, usar un DataFrame vacío
        df = pd.DataFrame()

    if not df.empty:
        st.markdown('---')
        st.subheader('Excel Original')
        st.dataframe(df)

        st.sidebar.header('Filtros:')

        profesor = df['Docente'].unique().tolist()
        tutor = df['TUTOR'].unique().tolist()
        materia = df['Materia'].unique().tolist()
        curso = df['Curso'].unique().tolist()

        maestro_select = st.sidebar.multiselect('Docente:',
                                        profesor,
                                        default=profesor)

        tutor_select = st.sidebar.multiselect('Tutor:',
                                        tutor,
                                        default=tutor)

        materia_select = st.sidebar.multiselect('Materia:',
                                        materia,
                                        default=materia)

        curso_select = st.sidebar.multiselect('Curso:',
                                        curso,
                                        default=curso)

        st.markdown('---')
        
        mask = (df['Docente'].isin(maestro_select)) & (df['TUTOR'].isin(tutor_select)) & (df['Materia'].isin(materia_select)) & (df['Curso'].isin(curso_select)) 
        result = df[mask].shape[0]
        st.markdown(f'*Todal de alumnos que cumples las condiciones: {result}*')

        st.subheader('Excel filtrado')
        st.dataframe(df[mask])

        if st.button("Generar Archivos para Tutores"):
            with st.spinner('Generando archivos, por favor espera...'):
                zip_data = io.BytesIO()
                with zipfile.ZipFile(zip_data, 'w') as zip_file:
                    for index, row in df[mask].iterrows():
                        template = DocxTemplate("tutorados.docx")
                        context = {
                            "nombre": row['Estudiante'],
                            "control": row['Control'],
                            "semestre": row['Semestre'],
                            "materia": row['Materia'],
                            "gpo": row['Grupo'],
                            "curso": row['Curso'],
                            "docente": row['Docente'],
                            "tutor": row['TUTOR'],
                            "date": fecha_actual
                        }
                        template.render(context)
                        output = io.BytesIO()
                        template.save(output)
                        zip_file.writestr(f"{row['Estudiante']}.docx",output.getvalue())
                    zip_data.seek(0)
                    download_link_zip = f'[Descargar ZIP de Tutores](data:application/zip;base64,{base64.b64encode(zip_data.read()).decode()})'
                    st.markdown(download_link_zip, unsafe_allow_html=True)
            st.success('Archivos generados exitosamente.')

        if st.button("Generar Archivos para Docentes"):
            with st.spinner('Generando archivos, por favor espera...'):
                zip_data = io.BytesIO()
                with zipfile.ZipFile(zip_data, "w") as zip_file:
                    for index, row in df[mask].iterrows():
                        template = DocxTemplate("docentes.docx")
                        context = {
                            "nombre": row['Estudiante'],
                            "control": row['Control'],
                            "semestre": row['Semestre'],
                            "materia": row['Materia'],
                            "gpo": row['Grupo'],
                            "curso": row['Curso'],
                            "docente": row['Docente'],
                            "tutor": row['TUTOR'],
                            "date": fecha_actual
                        }
                        template.render(context)
                        output = io.BytesIO()
                        template.save(output)
                        zip_file.writestr(f"{row['Estudiante']}.docx", output.getvalue())
                zip_data.seek(0)
                download_link_zip = f'[Descargar ZIP de Docentes](data:application/zip;base64,{base64.b64encode(zip_data.read()).decode()})'
                st.markdown(download_link_zip, unsafe_allow_html=True)
            st.success('Archivos generados exitosamente.')

        if st.button("Generar Archivos Excel por Docente"):
            with st.spinner('Generando archivos, por favor espera...'):
                columns_to_save = ['Estudiante', 'Control', 'Semestre', 'Curso', 'Materia']
                zip_data_excel = io.BytesIO()
                with zipfile.ZipFile(zip_data_excel, "w") as zip_file_excel:
                    for docente_value in df['Docente'].unique():
                        docente_mask = df['Docente'] == docente_value
                        df_for_docente = df[docente_mask][columns_to_save]                        
                        excel_file_name = f"{docente_value}.xlsx"
                        excel_file_path = Path(excel_file_name)
                        df_for_docente.to_excel(excel_file_path, index=False)
                        zip_file_excel.write(excel_file_path, excel_file_name)
                        os.remove(excel_file_path)
            zip_data_excel.seek(0)
            download_link_zip_excel = f'[Descargar ZIP para Excels de Docentes]' \
                                    f'(data:application/zip;base64,{base64.b64encode(zip_data_excel.read()).decode()})'
            st.markdown(download_link_zip_excel, unsafe_allow_html=True)
            st.success('Archivos generados exitosamente.')
