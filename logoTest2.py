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
file_path = 'Cutter_Correlation_Chart4.xlsx'  # Replace with your file path

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
        if 'OAL_inches' in df.columns:
            df['OAL_inches'] = df['OAL_inches'].str.replace('"', '').astype(float)
        if 'Blade_Length_inches' in df.columns:
            df['Blade_Length_inches'] = df['Blade_Length_inches'].str.replace('"', '').astype(float)
        if 'Body_Width_inches' in df.columns:
            df['Body_Width_inches'] = df['Body_Width_inches'].str.replace('"', '').astype(float)
        if 'Tip_inches' in df.columns:
            df['Tip_inches'] = df['Tip_inches'].str.replace('"', '').astype(float)
        if 'Type' in df.columns:
            df['Type'] = df['Type'].str.strip().str.title()
        if 'Sub-Type' in df.columns:
            df['Sub-Type'] = df['Sub-Type'].str.strip().str.title()
    return df

# Function to create display-friendly column names
def get_display_column_mapping(df):
    return {col: col.replace('_', ' ') for col in df.columns}

# Load all sheets
df_excelta = load_and_process_sheet('Excelta')
df_idealtek = load_and_process_sheet('ideal-tek')
df_swanstrom = load_and_process_sheet('Swanstrom')

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
sheet_selection = st.sidebar.selectbox('Select Brand', ['Excelta', 'ideal-tek', 'Swanstrom'])

# Select the appropriate DataFrame based on the user's choice
if sheet_selection == 'Excelta':
    df = df_excelta
    part_number_column = 'Part_#'
elif sheet_selection == 'ideal-tek':
    df = df_idealtek
    part_number_column = 'Part_Number'
elif sheet_selection == 'Swanstrom':
    df = df_swanstrom
    part_number_column = 'Model'

# Get unique values for dropdowns based on selected sheet
if sheet_selection == 'Excelta':
    sizes = ['None'] + sorted(df['Size'].dropna().unique().tolist())
    cuts = ['None'] + sorted(df['Cut'].dropna().unique().tolist())
    wires = ['None'] + sorted(df['Wire'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_size = st.sidebar.selectbox('Select Size', sizes, index=0)
    selected_cut = st.sidebar.selectbox('Select Cut', cuts, index=0)
    selected_wire = st.sidebar.selectbox('Select Wire', wires, index=0)
elif sheet_selection == 'ideal-tek':
    types = ['None'] + sorted(df['Type'].dropna().unique().tolist())
    head_shapes = ['None'] + sorted(df['Head_Shape'].dropna().unique().tolist())
    head_sizes = ['None'] + sorted(df['Head_Size'].dropna().unique().tolist())
    cutting_edges = ['None'] + sorted(df['Cutting_Edge'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_type = st.sidebar.selectbox('Select Type', types, index=0)
    selected_head_shape = st.sidebar.selectbox('Select Head Shape', head_shapes, index=0)
    selected_head_size = st.sidebar.selectbox('Select Head Size', head_sizes, index=0)
    selected_cutting_edge = st.sidebar.selectbox('Select Cutting Edge', cutting_edges, index=0)
elif sheet_selection == 'Swanstrom':
    types = ['None'] + sorted(df['Type'].dropna().unique().tolist())
    handle_types = ['None'] + sorted(df['Handle_Type'].dropna().unique().tolist())
    sub_types = ['None'] + sorted(df['Sub-Type'].dropna().unique().tolist())

    st.sidebar.subheader('Filter by Attributes')
    # User selections
    selected_type = st.sidebar.selectbox('Select Type', types, index=0)
    selected_handle_type = st.sidebar.selectbox('Select Handle Type', handle_types, index=0)
    selected_sub_type = st.sidebar.selectbox('Select Sub-Type', sub_types, index=0)

# Option to enter part number directly
st.sidebar.subheader('Direct Part Search')
part_number_input = st.sidebar.text_input('Enter Part Number (if known)')

if sheet_selection != 'Swanstrom':
    st.sidebar.subheader('Dimension Filters')
    mm_input = st.sidebar.text_input('Enter Millimeter Value')
    inches_input = st.sidebar.text_input('Enter Inches Value')
    awg_input = st.sidebar.text_input('Enter AWG Value')
else:
    st.sidebar.subheader('Dimension Filters')
    oal_input = st.sidebar.text_input('Enter OAL (if known)')
    blade_length_input = st.sidebar.text_input('Enter Blade Length (if known)')
    body_width_input = st.sidebar.text_input('Enter Body Width (if known)')
    tip_input = st.sidebar.text_input('Enter Tip (if known)')

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
if sheet_selection != 'Swanstrom':
    mm_value = to_float(mm_input) if mm_input else None
    inches_value = to_float(inches_input) if inches_input else None
    awg_value = to_int(awg_input) if awg_input else None
else:
    oal_value = to_float(oal_input) if oal_input else None
    blade_length_value = to_float(blade_length_input) if blade_length_input else None
    body_width_value = to_float(body_width_input) if body_width_input else None
    tip_value = to_float(tip_input) if tip_input else None

# Initialize filtered DataFrame as the full DataFrame
filtered_df = df.copy()

# Apply filters only if user has entered values
if sheet_selection == 'Excelta':
    if selected_size != 'None':
        filtered_df = filtered_df[filtered_df['Size'] == selected_size]
    if selected_cut != 'None':
        filtered_df = filtered_df[filtered_df['Cut'] == selected_cut]
    if selected_wire != 'None':
        filtered_df = filtered_df[filtered_df['Wire'] == selected_wire]
elif sheet_selection == 'ideal-tek':
    if selected_type != 'None':
        filtered_df = filtered_df[filtered_df['Type'] == selected_type]
    if selected_head_shape != 'None':
        filtered_df = filtered_df[filtered_df['Head_Shape'] == selected_head_shape]
    if selected_head_size != 'None':
        filtered_df = filtered_df[filtered_df['Head_Size'] == selected_head_size]
    if selected_cutting_edge != 'None':
        filtered_df = filtered_df[filtered_df['Cutting_Edge'] == selected_cutting_edge]
elif sheet_selection == 'Swanstrom':
    if selected_type != 'None':
        filtered_df = filtered_df[filtered_df['Type'] == selected_type]
    if selected_handle_type != 'None':
        filtered_df = filtered_df[filtered_df['Handle_Type'] == selected_handle_type]
    if selected_sub_type != 'None':
        filtered_df = filtered_df[filtered_df['Sub-Type'] == selected_sub_type]

if part_number_input:
    filtered_df = filtered_df[filtered_df[part_number_column] == part_number_input]
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
    if awg_value is not None:
        if sheet_selection == 'Excelta':
            filtered_df = filtered_df[(filtered_df['AWG_Low'] <= awg_value) & (filtered_df['AWG_High'] >= awg_value)]
        elif sheet_selection == 'ideal-tek':
            filtered_df = filtered_df[(filtered_df['Lowest_AWG'] <= awg_value) & (filtered_df['Highest_AWG'] >= awg_value)]
else:
    if oal_value is not None:
        filtered_df = filtered_df[(filtered_df['OAL_inches'] == oal_value)]
    if blade_length_value is not None:
        filtered_df = filtered_df[(filtered_df['Blade_Length_inches'] == blade_length_value)]
    if body_width_value is not None:
        filtered_df = filtered_df[(filtered_df['Body_Width_inches'] == body_width_value)]
    if tip_value is not None:
        filtered_df = filtered_df[(filtered_df['Tip_inches'] == tip_value)]

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
