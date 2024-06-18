import pandas as pd
import streamlit as st
import openpyxl
from PIL import Image

# Image Variables
icon = Image.open('PTLogo.jpeg')  # Original icon
sidebar_image = Image.open('PTLogo3.png')  # New image for the sidebar

# Set Page width and Title
st.set_page_config(page_title="Practical Tools Cutter Correlation Chart", page_icon=icon, layout="wide")

# Hide Streamlit default menu and footer
hide_default_format = """
<style>
#MainMenu {visibility: hidden; }
footer {visibility: hidden;}
.reportview-container .main .block-container {
    padding-top: 2rem;
}
</style>
"""
st.markdown(hide_default_format, unsafe_allow_html=True)

# Display the logo at the top of the sidebar
st.sidebar.image(sidebar_image, use_column_width=True)

# Load the data from the Excel file
file_path = 'Cutter_Correlation_Chart5H.xlsx'  # Replace with your file path

# Define a function to process each sheet
def load_and_process_sheet(sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Strip leading/trailing whitespaces from column names and normalize
    df.columns = df.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '').str.replace(' ', '_')

    if sheet_name == 'Excelta':
        if 'Size' in df.columns:
            df['Size'] = df['Size'].str.strip().str.title()
        if 'Millimeter_Low' in df.columns:
            df['Millimeter_Low'] = df['Millimeter_Low'].str.replace('mm', '').astype(float)
        if 'Millimeter_High' in df.columns:
            df['Millimeter_High'] = df['Millimeter_High'].str.replace('mm', '').astype(float)
        if 'Inches_Low' in df.columns:
            df['Inches_Low'] = df['Inches_Low'].str.replace('"', '').astype(float)
        if 'Inches_High' in df.columns:
            df['Inches_High'] = df['Inches_High'].str.replace('"', '').astype(float)
        if 'AWG_High' in df.columns:
            df['AWG_High'] = df['AWG_High'].str.replace('AWG', '').astype(int)
        if 'AWG_Low' in df.columns:
            df['AWG_Low'] = df['AWG_Low'].str.replace('AWG', '').astype(int)
    elif sheet_name == 'ideal-tek':
        if 'Head_Width_Millimeter' in df.columns:
            df['Head_Width_Millimeter'] = df['Head_Width_Millimeter'].str.replace('mm', '').astype(float)
        if 'Head_Width__inches' in df.columns:
            df['Head_Width__inches'] = df['Head_Width__inches'].str.replace('"', '').astype(float)
        if 'Lowest_Cutting_Capacity_Millimeter' in df.columns:
            df['Lowest_Cutting_Capacity_Millimeter'] = df['Lowest_Cutting_Capacity_Millimeter'].str.replace('mm', '').astype(float)
        if 'Highest_Cutting_Capacity__Millimeter' in df.columns:
            df['Highest_Cutting_Capacity__Millimeter'] = df['Highest_Cutting_Capacity__Millimeter'].str.replace('mm', '').astype(float)
        if 'Highest_AWG' in df.columns:
            df['Highest_AWG'] = pd.to_numeric(df['Highest_AWG'].str.replace('AWG', ''), errors='coerce').astype('Int64')
        if 'Lowest_AWG' in df.columns:
            df['Lowest_AWG'] = pd.to_numeric(df['Lowest_AWG'].str.replace('AWG', ''), errors='coerce').astype('Int64')
        if 'OAL_Millimeter' in df.columns:
            df['OAL_Millimeter'] = df['OAL_Millimeter'].str.replace('mm', '').astype(float)
        if 'OAL__inches' in df.columns:
            df['OAL__inches'] = df['OAL__inches'].str.replace('"', '').astype(float)
    elif sheet_name == 'Swanstrom':
        if 'Lowest_Cutting_Capacity_Inches' in df.columns:
            df['Lowest_Cutting_Capacity_Inches'] = df['Lowest_Cutting_Capacity_Inches'].str.replace('"', '').astype(float)
        if 'Highest_Cutting_Capacity_Inches' in df.columns:
            df['Highest_Cutting_Capacity_Inches'] = df['Highest_Cutting_Capacity_Inches'].str.replace('"', '').astype(float)
        if 'AWG_Low' in df.columns:
            df['AWG_Low'] = pd.to_numeric(df['AWG_Low'].str.replace('AWG', '').str.replace('AWH', ''), errors='coerce').astype('Int64')
        if 'AWG_High' in df.columns:
            df['AWG_High'] = pd.to_numeric(df['AWG_High'].str.replace('AWG', '').str.replace('AWH', ''), errors='coerce').astype('Int64')
    elif sheet_name == 'EREM':
        if 'Cutting_Capacity_Copper_Low' in df.columns:
            df['Cutting_Capacity_Copper_Low'] = df['Cutting_Capacity_Copper_Low'].str.replace('"', '').astype(float)
        if 'Cutting_Capacity_Copper_High' in df.columns:
            df['Cutting_Capacity_Copper_High'] = df['Cutting_Capacity_Copper_High'].str.replace('"', '').astype(float)
        if 'Cutting_Capacity_Medium_Wire_Low' in df.columns:
            df['Cutting_Capacity_Medium_Wire_Low'] = df['Cutting_Capacity_Medium_Wire_Low'].str.replace('"', '').astype(float)
        if 'Cutting_Capacity_Medium_Wire_High' in df.columns:
            df['Cutting_Capacity_Medium_Wire_High'] = df['Cutting_Capacity_Medium_Wire_High'].str.replace('"', '').astype(float)
        if 'Cutting_Capacity_Hard_Wire_Low' in df.columns:
            df['Cutting_Capacity_Hard_Wire_Low'] = df['Cutting_Capacity_Hard_Wire_Low'].str.replace('"', '').astype(float)
        if 'Cutting_Capacity_Hard_Wire_High' in df.columns:
            df['Cutting_Capacity_Hard_Wire_High'] = df['Cutting_Capacity_Hard_Wire_High'].str.replace('"', '').astype(float)
    return df

# Function to create display-friendly column names
def get_display_column_mapping(df):
    return {col: col.replace('_', ' ') for col in df.columns}

# Load all sheets
df_excelta = load_and_process_sheet('Excelta')
df_idealtek = load_and_process_sheet('ideal-tek')
df_swanstrom = load_and_process_sheet('Swanstrom')
df_erem = load_and_process_sheet('EREM')

# Custom CSS to widen the data tables
st.markdown(
    """
    <style>
    .reportview-container .main .block-container {
        max-width: 2400px;
        padding: 1rem;
    }
    .stDataFrame {
        max-width: 2000px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Create a Streamlit app
st.title('Cutter Correlation Information')
st.subheader('Select Tool Parameters from the Sidebar')

# Sidebar for user inputs
st.sidebar.header('Filter Options')

# Allow the user to select which sheet to view
sheet_selection = st.sidebar.selectbox('Select Brand', ['Excelta', 'ideal-tek', 'Swanstrom', 'EREM'])

# Select the appropriate DataFrame based on the user's choice
if sheet_selection == 'Excelta':
    df = df_excelta
    part_number_column = 'Part_#'
elif sheet_selection == 'ideal-tek':
    df = df_idealtek
    part_number_column = 'Part_Number'
elif sheet_selection == 'Swanstrom':
    df = df_swanstrom
    part_number_column = 'Model_#'
elif sheet_selection == 'EREM':
    df = df_erem
    part_number_column = 'Part_Number'

# Get unique values for dropdowns based on selected sheet
if sheet_selection == 'Excelta':
    head_shapes = ['None'] + sorted(df['Size'].dropna().unique().tolist())
    types_of_cut = ['None'] + sorted(df['Cut'].dropna().unique().tolist())
    material_strengths = ['None'] + sorted(df['Wire'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_head_shape = st.sidebar.selectbox('Select Head Shape', head_shapes, index=0)
    selected_type_of_cut = st.sidebar.selectbox('Select Type of Cut', types_of_cut, index=0)
    selected_material_strength = st.sidebar.selectbox('Select Material Strength', material_strengths, index=0)
elif sheet_selection == 'ideal-tek':
    cutter_hardness = ['None'] + sorted(df['Type'].dropna().unique().tolist())
    head_shapes = ['None'] + sorted(df['Head_Shape'].dropna().unique().tolist())
    head_sizes = ['None'] + sorted(df['Head_Size'].dropna().unique().tolist())
    types_of_cut = ['None'] + sorted(df['Cutting_Edge'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_cutter_hardness = st.sidebar.selectbox('Select Cutter Hardness', cutter_hardness, index=0)
    selected_head_shape = st.sidebar.selectbox('Select Head Shape', head_shapes, index=0)
    selected_head_size = st.sidebar.selectbox('Select Head Size', head_sizes, index=0)
    selected_type_of_cut = st.sidebar.selectbox('Select Type of Cut', types_of_cut, index=0)
elif sheet_selection == 'Swanstrom':
    types_of_cut = ['None'] + sorted(df['Cut'].dropna().unique().tolist())
    material_strengths = ['None'] + sorted(df['Type_of_Cut'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_type_of_cut = st.sidebar.selectbox('Select Type of Cut', types_of_cut, index=0)
    selected_material_strength = st.sidebar.selectbox('Select Material Strength', material_strengths, index=0)
elif sheet_selection == 'EREM':
    series_of_cutters = ['None'] + sorted(df['Series_of_Cutter'].dropna().unique().tolist())
    types_of_cut = ['None'] + sorted(df['Cut_Type'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_series_of_cutter = st.sidebar.selectbox('Select Series of Cutter', series_of_cutters, index=0)
    selected_type_of_cut = st.sidebar.selectbox('Select Type of Cut', types_of_cut, index=0)

# Option to enter part number directly
st.sidebar.subheader('Direct Part Search')
part_number_input = st.sidebar.text_input('Enter Part Number (if known)')

# Show appropriate dimension filters based on sheet selection
st.sidebar.subheader('Dimension Filters')
if sheet_selection != 'EREM':
    if sheet_selection != 'Swanstrom':
        mm_input = st.sidebar.text_input('Enter Millimeter Value')
    inches_input = st.sidebar.text_input('Enter Inches Value')
    awg_input = st.sidebar.text_input('Enter AWG Value')
else:
    copper_input = st.sidebar.text_input('Enter Copper Wire Value')
    medium_input = st.sidebar.text_input('Enter Medium Wire Value')
    hard_input = st.sidebar.text_input('Enter Hard Wire Value')

# Function to convert input to float
def to_float(value):
    try:
        return float(value.strip())
    except ValueError:
        return None

# Function to convert input to int
def to_int(value):
    try:
        return int(value.strip())
    except ValueError:
        return None

# Convert inputs to appropriate types if they are provided
if sheet_selection != 'EREM':
    if sheet_selection != 'Swanstrom':
        mm_value = to_float(mm_input) if mm_input else None
    awg_value = to_int(awg_input) if awg_input else None
    inches_value = to_float(inches_input) if inches_input else None
else:
    copper_value = to_float(copper_input) if copper_input else None
    medium_value = to_float(medium_input) if medium_input else None
    hard_value = to_float(hard_input) if hard_input else None

# Initialize filtered DataFrame as the full DataFrame
filtered_df = df.copy()

# Apply filters only if user has entered values
if sheet_selection == 'Excelta':
    if selected_head_shape != 'None':
        filtered_df = filtered_df[filtered_df['Size'] == selected_head_shape]
    if selected_type_of_cut != 'None':
        filtered_df = filtered_df[filtered_df['Cut'] == selected_type_of_cut]
    if selected_material_strength != 'None':
        filtered_df = filtered_df[filtered_df['Wire'] == selected_material_strength]
elif sheet_selection == 'ideal-tek':
    if selected_cutter_hardness != 'None':
        filtered_df = filtered_df[filtered_df['Type'] == selected_cutter_hardness]
    if selected_head_shape != 'None':
        filtered_df = filtered_df[filtered_df['Head_Shape'] == selected_head_shape]
    if selected_head_size != 'None':
        filtered_df = filtered_df[filtered_df['Head_Size'] == selected_head_size]
    if selected_type_of_cut != 'None':
        filtered_df = filtered_df[filtered_df['Cutting_Edge'] == selected_type_of_cut]
elif sheet_selection == 'Swanstrom':
    if selected_type_of_cut != 'None':
        filtered_df = filtered_df[filtered_df['Cut'] == selected_type_of_cut]
    if selected_material_strength != 'None':
        filtered_df = filtered_df[filtered_df['Type_of_Cut'] == selected_material_strength]
elif sheet_selection == 'EREM':
    if selected_series_of_cutter != 'None':
        filtered_df = filtered_df[filtered_df['Series_of_Cutter'] == selected_series_of_cutter]
    if selected_type_of_cut != 'None':
        filtered_df = filtered_df[filtered_df['Cut_Type'] == selected_type_of_cut]

if part_number_input:
    filtered_df = filtered_df[filtered_df[part_number_column] == part_number_input]

if sheet_selection != 'EREM':
    if sheet_selection != 'Swanstrom':
        if mm_value is not None:
            if sheet_selection == 'Excelta':
                filtered_df = filtered_df[(filtered_df['Millimeter_Low'] <= mm_value) & (filtered_df['Millimeter_High'] >= mm_value)]
            elif sheet_selection == 'ideal-tek':
                filtered_df = filtered_df[(filtered_df['Lowest_Cutting_Capacity_Millimeter'] <= mm_value) & (filtered_df['Highest_Cutting_Capacity__Millimeter'] >= mm_value)]
    if inches_value is not None:
        if sheet_selection == 'Excelta':
            filtered_df = filtered_df[(filtered_df['Inches_Low'] <= inches_value) & (filtered_df['Inches_High'] >= inches_value)]
        elif sheet_selection == 'ideal-tek':
            filtered_df = filtered_df[(filtered_df['Head_Width__inches'] <= inches_value) & (filtered_df['Head_Width__inches'] >= inches_value)]
        elif sheet_selection == 'Swanstrom':
            filtered_df = filtered_df[(filtered_df['Lowest_Cutting_Capacity_Inches'] <= inches_value) & (filtered_df['Highest_Cutting_Capacity_Inches'] >= inches_value)]
if sheet_selection == 'EREM':
    if copper_value is not None:
        filtered_df = filtered_df[(filtered_df['Cutting_Capacity_Copper_Low'] <= copper_value) & (filtered_df['Cutting_Capacity_Copper_High'] >= copper_value)]
    if medium_value is not None:
        filtered_df = filtered_df[(filtered_df['Cutting_Capacity_Medium_Wire_Low'] <= medium_value) & (filtered_df['Cutting_Capacity_Medium_Wire_High'] >= medium_value)]
    if hard_value is not None:
        filtered_df = filtered_df[(filtered_df['Cutting_Capacity_Hard_Wire_Low'] <= hard_value) & (filtered_df['Cutting_Capacity_Hard_Wire_High'] >= hard_value)]

if sheet_selection != 'EREM' and awg_value is not None:
    if sheet_selection == 'Excelta':
        filtered_df = filtered_df[(filtered_df['AWG_Low'] <= awg_value) & (filtered_df['AWG_High'] >= awg_value)]
    elif sheet_selection == 'ideal-tek':
        filtered_df = filtered_df[(filtered_df['Lowest_AWG'] <= awg_value) & (filtered_df['Highest_AWG'] >= awg_value)]
    elif sheet_selection == 'Swanstrom':
        filtered_df = filtered_df[(filtered_df['AWG_Low'] <= awg_value) & (filtered_df['AWG_High'] >= awg_value)]

# Display the filtered DataFrame with display-friendly column names
st.subheader('Filtered Parts Information')
display_column_mapping = get_display_column_mapping(filtered_df)
filtered_df_display = filtered_df.rename(columns=display_column_mapping)

# Apply some styling to the DataFrame
styled_df = filtered_df_display.style.set_properties(**{'text-align': 'left'}).set_table_styles([dict(selector='th', props=[('text-align', 'left')])])

# Apply formatting to remove extra zeros
styled_df = styled_df.format(precision=3, na_rep='')

st.dataframe(styled_df, width=2000)

# Ensure there are rows in the filtered DataFrame before proceeding
if not filtered_df.empty:
    # Get unique part numbers from the filtered DataFrame
    part_numbers = filtered_df[part_number_column].unique()
    selected_part = st.selectbox('Select Part Number', part_numbers)

    # Show selected part number and part information
    st.subheader(f'Selected Part Number: {selected_part}')

    part_info = filtered_df[filtered_df[part_number_column] == selected_part]
    part_info_display = part_info.rename(columns=display_column_mapping)
    st.subheader('Details for Selected Part')
    styled_part_info = part_info_display.style.set_properties(**{'text-align': 'left'}).set_table_styles([dict(selector='th', props=[('text-align', 'left')])])

    # Apply formatting to remove extra zeros
    styled_part_info = styled_part_info.format(precision=3, na_rep='')

    st.dataframe(styled_part_info, width=2000)
else:
    st.write('No parts match the selected criteria.')

# To run the app, save this script and run `streamlit run script_name.py` in the terminal
