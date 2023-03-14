import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
from PIL import Image

@st.cache
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='DATA')
    workbook = writer.book
    worksheet = writer.sheets['DATA']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


@st.cache(allow_output_mutation=True)
def load_data(uploaded_file):
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
        except:
            pass

        return df


st.set_page_config(page_title='An Outcome')

st.header('Outcome 20**')
st.subheader("Subheader")

uploaded_data = st.file_uploader('Upload Data', key='1')

if uploaded_data:
    uploaded_data_df = load_data(uploaded_data)
    sheet_name = 'DATA'
    dv = pd.read_excel(uploaded_data,
                       sheet_name=sheet_name,
                       usecols='B:D',
                       header=3, )

    df_participants = pd.read_excel(uploaded_data,
                                    sheet_name=sheet_name,
                                    usecols='F:G',
                                    header=3,
                                    dtype='str')

    df_participants.dropna(inplace=True)

    department = dv['Department'].unique().tolist()
    ages = dv['Age'].unique().tolist()

    age_selection = st.slider('Age:',
                              min_value=min(ages),
                              max_value=max(ages),
                              value=(min(ages), max(ages)))

    department_selection = st.multiselect('Department:',
                                          department,
                                          default=department)
    mask = (dv['Age'].between(*age_selection)) & (dv['Department'].isin(department_selection))
    number_of_result = dv[mask].shape[0]
    st.markdown(f'*Available Results" {number_of_result}*')

    df_gathering = dv[mask].groupby(by=['Rating']).count()[['Age']]
    df_gathering = df_gathering.rename(columns={'Age': 'Votes'})
    df_gathering = df_gathering.reset_index()

    bar_chart = px.bar(df_gathering,
                       x='Rating',
                       y='Votes',
                       text='Votes',
                       color_discrete_sequence=['#F63366'] * len(df_gathering),
                       template='plotly_white')
    st.plotly_chart(bar_chart)

    st.dataframe(dv)
    # st.dataframe(df_participants)

    pie_chart = px.pie(df_participants,
                       title='Total No. of Participants',
                       values='Participants',
                       names='Departments')
    st.plotly_chart(pie_chart)

    image = Image.open("D:/Dadash/Images/362755.jpg")
    st.image(image,
             caption='Designed by Dadash Guluzade',
             use_column_width=True)

    # final_df = uploaded_data_df.copy()
