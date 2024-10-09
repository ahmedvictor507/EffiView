from tkinter import filedialog
import pandas as pd
import plotly.express as px


def read_excel_file(sheet_name):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    if file_path:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    return None


def process_data_em(df_em):
    print("Columns in df_em:", df_em.columns)  # Debugging line

    # Define the desired column and their order for the DataFrame
    desired_data_em = [
        'Job',
        'Item',
        'Description',
        'Employee',
        'Operation',
        'Planned Hours',
        'Total Clocked Hours'
    ]

    # Check if required columns exist
    missing_columns = [col for col in desired_data_em if col not in df_em.columns]

    if missing_columns:
        raise ValueError(f"Missing columns in employee data: {', '.join(missing_columns)}")

    # Filter out data where 'Planned Hours' or 'Total Clocked Hours' are less than or equal to 0.5
    df_em = df_em[df_em['Planned Hours'] > 0.5]
    df_em = df_em[df_em['Total Clocked Hours'] > 0.5]

    # Select only the columns specified in 'desired_data_em'
    df_em = df_em[desired_data_em]

    # Combine 'Planned Hours' and 'Total Clocked Hours' and store them in 'Efficiency (%)'
    df_em['Efficiency (%)'] = (df_em['Planned Hours'] / df_em['Total Clocked Hours']) * 100

    # Specify the names that need to be excluded from the DataFrame
    names_to_exclude = ['AMIRUDDIN  BIN BIDEN  AMIR', 'Amalan  Arul Alphonse', 'Anand Chauhan',
                        'Ananthan Baskar', 'DATAR  SINGH', 'DATAR SINGH', 'HANIF BIN MUHAMAD ISA',
                        'Ibrahim  Bin Baharun', 'JENORIKI ANAK LINGONG', 'JUWAHIR BIN SALLEH',
                        'Krishnakumar  Jayaraman', 'Krishnakumar Jayaraman', 'LOVEPREET .',
                        'MOHM AZREEQ BILAL BIN NORDIN', 'MUHAMAD AZUARI BIN MUHAMMAD NOR',
                        'MUNIYAN PICHAIKARAN PREM KUMAR', 'Mustapha Kamall',
                        'PREM KUMAR MUNIYAN PICHAIKARAN', 'Ricky Anak Anthony Ginyan',
                        'SANTOSH KUMAR', 'Sujith Pillai', 'VEDAIYAN SINGARAVELU',
                        'SHAHRIZAL BIN SAID', 'Vivek Lingapandi', 'Joseph Kulandai Stanislaus',
                        'Nur Syafiq  Bin Kamal']

    # Remove leading and trailing whitespace from the 'Employee' column
    df_em['Employee'] = df_em['Employee'].str.strip()

    # Exclude employees that are in 'names_to_exclude'
    # If you want to do the opposite, create names to_include and delete '~' from this line
    df_em = df_em[~df_em['Employee'].isin(names_to_exclude)]
    # Filter out the 'Efficiency (%)' below 120
    # df_em = df_em[df_em['Efficiency (%)'] < 120]
    # Reset the index of the DataFrame, dropping the old index
    df_em.reset_index(drop=True, inplace=True)
    return df_em


def plot_employee_efficiency(df_em):

    # Calculate the average efficiency percentage for each employee
    average_efficiency = df_em.groupby('Employee')['Efficiency (%)'].mean()
    # Calculate the average planned hours for each employee
    average_estimated_hours = df_em.groupby('Employee')['Planned Hours'].mean()
    # Calculate the average total clocked hours for each employee
    average_spent_hours = df_em.groupby('Employee')['Total Clocked Hours'].mean()

    # Create a DataFrame for plotting, consolidating performance metrics for each employee
    df_em_plot = pd.DataFrame({
        'Employee': average_efficiency.index,
        'average_efficiency': average_efficiency.values,
        'average_estimated_hours': average_estimated_hours.values,
        'average_spent_hours': average_spent_hours.values
    })

    # Add custom text for hover
    df_em_plot['custom_text'] = (
        df_em_plot['Employee'] +
        '<br>Average Efficiency: ' + df_em_plot['average_efficiency'].round(2).astype(str) + '%' +
        '<br>Average Estimated Hours: ' + df_em_plot['average_estimated_hours'].round(2).astype(str) +
        '<br>Average Spent Hours: ' + df_em_plot['average_spent_hours'].round(2).astype(str)
    )

    # Create a bar graph that uses df_em_plot data
    fig = px.bar(df_em_plot,
                 x='Employee',
                 y='average_efficiency',
                 labels={'Employee': 'Employee name', 'average_efficiency': 'Average Efficiency (%)'},
                 title='Average Efficiency for Each Employee',
                 text='average_efficiency',
                 custom_data=['custom_text'])

    # Update traces in the figure for better visualization
    fig.update_traces(texttemplate='%{text:.2f}',
                      hovertemplate='%{customdata[0]}<extra></extra>')

    # Adjust the layout of the x-axis for better readability
    fig.update_layout(xaxis_tickangle=-45)
    fig.show()


def process_data_op(df_op):
    # Define the desired column and their order for the DataFrame
    desired_data_op = [
        'Job',
        'Make Item',
        'Operation',
        'Estimated Total Hours',
        'Actual Total Hours',
        'Total Hours Variance',
        'Job Completed On'
    ]

    # Select only the columns specified in 'desired_data_op'
    df_op = df_op[desired_data_op]

    # Calculate the efficiency as a percentage by dividing estimated hours by actual hours and multiplying by 100
    df_op.loc[:, 'Efficiency (%)'] = (df_op['Estimated Total Hours']/df_op['Actual Total Hours'])*100
    # Filter out rows where actual total hours are less than or equal to 0.5 to ensure valid efficiency calculations
    df_op = df_op[df_op['Actual Total Hours'] > 0.5]
    # Further filter the DataFrame to exclude operations with efficiency greater than or equal to 120%
    df_op = df_op[df_op['Efficiency (%)'] < 120]
    # Reset the index of the DataFrame, dropping the old index
    df_op.reset_index(drop=True, inplace=True)

    return df_op


def plot_operation_efficiency(df_op):
    # Calculate average efficiency, estimated total hours, and actual total hours for each operation
    average_efficiency = df_op.groupby('Operation')['Efficiency (%)'].mean()
    average_estimated_hours = df_op.groupby('Operation')['Estimated Total Hours'].mean()
    average_spent_hours = df_op.groupby('Operation')['Actual Total Hours'].mean()

    # Create a DataFrame to hold the average metrics for each operation
    df_op_plot = pd.DataFrame({
        'Operation': average_efficiency.index,
        'average_efficiency': average_efficiency.values,
        'average_estimated_hours': average_estimated_hours.values,
        'average_spent_hours': average_spent_hours.values
    })

    # Create a custom text field for hover information in the plot
    df_op_plot['custom_text'] = (
        df_op_plot['Operation'] +
        '<br>Average Efficiency: ' + df_op_plot['average_efficiency'].round(2).astype(str) + '%' +
        '<br>Average Estimated Hours: ' + df_op_plot['average_estimated_hours'].round(2).astype(str) +
        '<br>Average Spent Hours: ' + df_op_plot['average_spent_hours'].round(2).astype(str)
    )

    # Create a bar plot using Plotly Express
    fig = px.bar(df_op_plot,
                 x='Operation',
                 y='average_efficiency',
                 labels={'Operation': 'Operation Name', 'average_efficiency': 'Average Efficiency (%)'},
                 title='Average Efficiency for Each Operation',
                 text='average_efficiency',
                 custom_data=['custom_text'])

    # Update the traces of the figure for text formatting and hover information
    fig.update_traces(texttemplate='%{text:.2f}',
                      hovertemplate='%{customdata[0]}<extra></extra>')
    # Rotate x-axis labels for better readability
    fig.update_layout(xaxis_tickangle=-45)

    fig.show()