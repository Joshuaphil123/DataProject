import streamlit as st
import pandas as pd
import openai
from dotenv import load_dotenv
from openai import OpenAI
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.chart.data import ChartData
from pptx.util import Inches
load_dotenv()
import os
import hmac
import gdown

st.title("Data Report Generator 📈")

# Direct link to the file
url = 'https://drive.google.com/uc?id=1ybc7M6N0_8AmgTlgVas8WM2Nj5qSI5Ob'
output = 'Revenue Analysis.pptx'
gdown.download(url, output, quiet=False)

# def check_password():
#     """Returns `True` if the user had a correct password or not."""

#     def login_form():
#         """Form with widgets to collect user information"""
#         with st.form("Credentials"):
#             st.text_input("Username", key="username")
#             st.text_input("Password", type="password", key="password")
#             st.form_submit_button("Log in", on_click=password_entered)

#     def password_entered():
#         """Checks whether a password entered by the user is correct."""
#         if st.session_state["username"] in st.secrets[
#             "passwords"
#         ] and hmac.compare_digest(
#             st.session_state["password"],
#             st.secrets.passwords[st.session_state["username"]],
#         ):
#             st.session_state["password_correct"] = True
#             del st.session_state["password"]  # Don't store the username or password.
#             del st.session_state["username"]
#         else:
#             st.session_state["password_correct"] = False

#     # Return True if the username + password is validated.
#     if st.session_state.get("password_correct", False):
#         return True

#     # Show inputs for username + password.
#     login_form()
#     if "password_correct" in st.session_state:
#         st.error("😕 User not known or password incorrect")
#     return False

# if not check_password():
#     st.stop()

#Variable Declaration

sales_data = ""
target_data = ""

# Importing the datasets
sales_Data_name = st.file_uploader("Upload your Data File in XLSX format", type=['xlsx'])
target_Data_name = st.file_uploader("Upload your Target file in XLSX format", type=['xlsx'])
  

year1 = st.selectbox(
"Select the Year",
(2023,2024))  

month1 = st.selectbox(
"Select the month",
(1,2,3,4,5,6,7,8,9,10,11,12)
)

if st.button("Generate Report"):
    with st.spinner("Loading .. 🔃"):
        if sales_Data_name is not None and target_Data_name is not None:
            try:
                sales_data = pd.read_excel(sales_Data_name)
                target_data = pd.read_excel(target_Data_name)
            except Exception as e:
                st.write("Error reading excel files - ", e)  
        st.title("Data Report ")


        

        # Convert the 'Date' columns to datetime
        sales_data['Date'] = pd.to_datetime(sales_data['Date'])
        target_data['Date'] = pd.to_datetime(target_data['Date'])
        sales= sales_data[(sales_data['Date'].dt.year == 2023) & (sales_data['Date'].dt.month == 11)]
        previous_sales= sales_data[(sales_data['Date'].dt.year == 2022) & (sales_data['Date'].dt.month == 11)]
        Current_year_to_date_sales= sales_data[(sales_data['Date'].dt.year == 2023) & (sales_data['Date'].dt.month >=10)& (sales_data['Date'].dt.month <=11)]
        Previous_year_to_date_sales= sales_data[(sales_data['Date'].dt.year == 2022) & (sales_data['Date'].dt.month >=10)& (sales_data['Date'].dt.month <=11)]

      # Filter target data for November 2023 and get the Company Total
        target= target_data[(target_data['Date'].dt.year == 2023) & (target_data['Date'].dt.month == 11)]
        company_total_target = target[target['Attributes'] == 'Company Total']['Value'].sum()
        total_sales = round(sales['Sales_FC'].sum())

        # print(company_total_target) 
        difference_sales_target =total_sales-company_total_target
        # print(difference_sales_target )
        sales_to_target_ratio = total_sales / company_total_target
        # print(sales_to_target_ratio)
        total_previous_sales = round(previous_sales['Sales_FC'].sum())

        current_year_to_date_sales = round(Current_year_to_date_sales['Sales_FC'].sum())
        previous_year_to_date_sales = round(Previous_year_to_date_sales['Sales_FC'].sum())
        
        growth_yoy = round((total_sales - total_previous_sales)/total_previous_sales,1)
        # print(growth_yoy)
        growth_year_to_date = round((current_year_to_date_sales-previous_year_to_date_sales)/previous_year_to_date_sales,1)
        # print(growth_year_to_date)

                # Filter target data for November 2023 and get the Company Total
        target= target_data[(target_data['Date'].dt.year == 2023) & (target_data['Date'].dt.month == 11)]
        company_total_target = round(target[target['Attributes'] == 'Company Total']['Value'].sum())

        # print(company_total_target)
        difference_sales_target =round(total_sales-company_total_target)
        
        # print(difference_sales_target )
        # print(sales_to_target_ratio)
          
       # Store the variables
        data = {
            'Total Sales': total_sales,
            'Company Total Target': company_total_target,
            'Difference (Sales - Target)': difference_sales_target ,
            'Sales/Target Ratio': sales_to_target_ratio,
            'Total Previous Sales': total_previous_sales,
            'Growth':growth_yoy,
            'Current Year to Date Sales':current_year_to_date_sales,
            'Previous Year to Date Sales':previous_year_to_date_sales,
            'Growth Year to Date':growth_year_to_date
        }   

        
        prompt = f"""
        Here is the sales data for November 2023:
        Total Sales: {data['Total Sales']}
        Company Total Target: {data['Company Total Target']}
        Difference (Sales - Target): {data['Difference (Sales - Target)']}
        Sales to target ratio :{data['Sales/Target Ratio']}


        Please provide a brief two line observation based on the summary of this data and include the difference of Sales - target into the observation
        """


        # In[30]:


        client= openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


        # In[31]:


        completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a data analyst with ability to generate quality observation from the given data"},
            {"role": "user", "content": prompt}
        ]
        )


        # Print the observation
        observation=completion.choices[0].message.content
        print(observation)


        # In[32]:


        prompt2 = f"""
        Here is the sales data for the current year and year on year:
        Current Month Sales: {data['Total Sales']}
        Year on Year Sales: {data['Total Previous Sales']}
        Growth/Decline from year on year sales to current month sales: {data['Growth']}


        Please provide a brief two line observation based on the summary of this data.
        """


        # In[33]:


        completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a data analyst with ability to generate quality observation from the given data"},
            {"role": "user", "content": prompt2}
        ]
        )


        # Print the observation
        observation2=completion.choices[0].message.content
        print(observation2)


        # In[34]:


        prompt3 = f"""
        Here is the current year to date sales, previous year to date sales and growth/decline of previous to current year to date sales :
        Current year to date sales: {data['Current Year to Date Sales']}
        Preivous year to date sales: {data['Previous Year to Date Sales']}
        Growth/Decline of previous to current year to date sales: {data['Growth Year to Date']}


        Please provide a brief two line observation based on the summary of this data.
        """


        # In[35]:


        completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a data analyst with ability to generate quality observation from the given data"},
            {"role": "user", "content": prompt3}
        ]
        )


        # Print the observation
        observation3=completion.choices[0].message.content
        print(observation3)


        prompt4 = f"""
        Here are the three observations :
        Target preformance: {['observation']}
        Year on Year sales performance: {['observation2']}
        Year to Date Performance: {['observation3']}


        Please provide a two line observation in one single para based on the three observations
        """

        completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a data analyst with ability to generate quality observation from the given data"},
            {"role": "user", "content": prompt4}
        ]
        )


        # Print the observation
        observation4=completion.choices[0].message.content
        # print(observation4)

        df = pd.DataFrame(list(data.items()), columns=['Metric', 'Value'])

        st.write(df)

        with open("Revenue Analysis.pptx", "rb") as file:
            btn = st.download_button(
                label="Download Summary",
                data=file,
                file_name="Revenue Analysis.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

