import requests  # Used to extract data from api
import pandas as pd  # Pandas
import plotly.express as px  # Used to graph the Gantt chart
import plotly.graph_objects as go  # Used to plot the Gantt chart
from datetime import datetime  # Used to extract date & time format from strings
import pytz  # Adjust timezone when extracting today's date & time
from concurrent.futures import ThreadPoolExecutor, as_completed


def extract_data_from_api():
    # API URLs needed
    all_jobs_list_url_template = 'https://api.fulcrumpro.com/api/jobs/list'  # URL used to extract job Ids
    operations_url_template = 'https://api.fulcrumpro.com/api/jobs/{jobId}/operations/list'  # URL used to extract operations provided the jobId
    job_list_url_template = 'https://api.fulcrumpro.com/api/jobs/{jobId}'  # URL to extract info related to a job provided the jobId
    sales_list_url_template = 'https://api.fulcrumpro.com/api/sales-orders/{salesOrderId}'  # URL to extract sale info provided the jobId

    # Header that provides authentication and content type
    headers_all_jobs_list = {
        # Authorization key (very important, share with authorized people only)
        'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJOYW1lIjoiRGF0YSBSZXRyaWV2ZWwiLCJSZXZvY2F0aW9uSWQiOiJlYTJjNjQ2OS1lY2VjLTQ2ODQtYjlkNS01NDFiNGZlMmZjZjIiLCJleHAiOjE3NTEyMTI4MDAsImlzcyI6Im1laWJhbi5mdWxjcnVtcHJvLmNvbSIsImF1ZCI6Im1laWJhbiJ9.ariY2Q8msoV9lmYa2WHNb7nt5uC-yiERQvCdjRCmMuM',
        'Content-Type': 'application/json-patch+json'
    }

    # Header that provides authentication and content type
    headers = {
        # Authorization key (very important, share with authorized people only)
        'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJOYW1lIjoiRGF0YSBSZXRyaWV2ZWwiLCJSZXZvY2F0aW9uSWQiOiJlYTJjNjQ2OS1lY2VjLTQ2ODQtYjlkNS01NDFiNGZlMmZjZjIiLCJleHAiOjE3NTEyMTI4MDAsImlzcyI6Im1laWJhbi5mdWxjcnVtcHJvLmNvbSIsImF1ZCI6Im1laWJhbiJ9.ariY2Q8msoV9lmYa2WHNb7nt5uC-yiERQvCdjRCmMuM',
        'Content-Type': 'application/json'
    }

    # Empty list to store data
    all_data = []

    # Function to extract all job ids
    def get_all_job_ids(status, take, skip):
        url = f"{all_jobs_list_url_template}?take={take}&skip={skip}"
        # Filters job ids extracted according to status (Note : when left empty api doesn't extract all jobs)
        payload = {"status": status}

        # Make a POST request from api URl, provide authorization in headers
        response = requests.post(all_jobs_list_url_template, headers=headers_all_jobs_list, json=payload)
        if response.status_code == 200:  # status_code = 200 means extraction was successful
            jobs_data = response.json()
            job_ids = [job['id'] for job in jobs_data]
            return job_ids
        else:
            print(f"Failed to get job list with status code {response.status_code}")
            return []

    # Call the function and store job ids in job_ids list
    job_ids = get_all_job_ids("inProgress", 100, 0)

    # Function to extract data using the jobId extracted
    def fetch_data(jobId):
        try:
            job_url = job_list_url_template.format(jobId=jobId)  # Assign the jobId to the job_list_url
            print(f"Processing Job ID : {jobId}")
            response1 = requests.get(job_url, headers=headers)  # Make a GET request from the API
            print(f"Response status code : {response1.status_code}")
            if response1.status_code != 200:
                return []

            # Store the data extracted in json format in job_data list
            job_data = response1.json()
            # List to extract certain data from the json file ( Can be found under response sample in the api info website)
            job_data = {
                'Job': job_data['name'],
                #'job_status': job_data['status'],
                'job_salesOrderId': job_data['salesOrderId'],
                'Date Created': job_data['createdUtc'],
                'Job Est. Completion': job_data['scheduledEndUtc']
            }

            sales_url = sales_list_url_template.format(
                salesOrderId=job_data['job_salesOrderId'])  # Assign the jobId to the sales_list_url
            print(f"Processing Sales Order ID: {jobId}")
            response2 = requests.get(sales_url, headers=headers)  # Make a GET request from the API
            print(f"Response status code : {response2.status_code}")
            if response2.status_code != 200:
                return []

            # Store the data extracted in json format in sales_data list
            sales_data = response2.json()
            # List to extract certain data from the json file ( Can be found under response sample in the api info website)
            sales_data = {
                #'sales_id': sales_data['id'],
                'Sales Order': sales_data['number'],
                #'sales_orderedDate': sales_data['orderedDate'],
                'Delivery Due Date': sales_data['deliveryDueDate']
            }

            operations_url = operations_url_template.format(jobId=jobId)  # Assign the jobId to the operation_url
            print(f"Processing Operations Order ID: {jobId}")
            response3 = requests.post(operations_url, headers=headers)  # Make a POST request from the API
            print(f"Response status code : {response3.status_code}")
            if response3.status_code != 200:
                return []

            # Store the data extracted in json format in operations_data list
            operations_data = response3.json()

            # Make an empty list to store extracted data and flatten it
            flattened_data = []

            # Loop to extract data categories from operations_data ( Refer to response sample in https://developers.fulcrumpro.com/api-schema#tag/Job-Operation)
            for item in operations_data:
                itemToMake = item['itemToMake']
                operation = item['operation']

                item_to_make_data = {
                    'Job Item': itemToMake['itemReference']['number'],
                    'Job Item Description': itemToMake['itemReference']['description']
                }

                # Extract operation data
                operation_data = {
                    'Status': operation['status'],
                    'Step': operation['order'],
                    'Actual Start': operation['scheduledStartUtc'],
                    'Instructions': operation['instructions'],
                    'Actual End': operation['completedOnUtc'],
                    'Operation': operation['name'],
                }

                # Combine all data extracted into one list
                combined_data = {**job_data, **sales_data, **item_to_make_data, **operation_data}

                # Add contents from combined_data to flattened_data
                flattened_data.append(combined_data)

            return flattened_data

        except Exception as e:
            print(f"An error occurred while processing job ID {jobId}: {e}")
            return []

    # Create a ThreadPoolExecutor to manage a pool of threads, allowing up to 20 threads to run concurrently
    with ThreadPoolExecutor(max_workers=50) as executor:
        # Submit tasks to the executor for each jobId in the job_ids list
        futures = [executor.submit(fetch_data, jobId) for jobId in job_ids]

        # Iterate over the futures as they complete, allowing you to handle each result as soon as it's ready
        for future in as_completed(futures):

            # Retrieve the result from the future object (the output of fetch_data)
            result = future.result()

            # Check if the result is not empty (or None); if it has valid data, extend the all_data list with the result
            if result:
                all_data.extend(result)

    # Store data into a dataframe from all_data list
    df = pd.DataFrame(all_data)

    # Convert Job from string format to numeric format
    df['Job'] = pd.to_numeric(df['Job'], errors='coerce')

    # Sort values in the dataframe according to sales_deliveryDueDate in ascending order
    df = df.sort_values(by='Delivery Due Date', ascending=True)

    return df


def process_df(df_job, df_op):
    # Extract data from the api and store it in dataframe
    df_api = extract_data_from_api()

    # Create a copy of the dataframe
    df_api_copy = df_api.copy()
    # Filter the DataFrame to include only rows where the status is 'complete'
    df_api_copy = df_api_copy[df_api_copy["Status"] == 'complete']
    # Drop unnecessary columns that are no longer required for further analysis or processing
    df_api_copy = df_api_copy.drop(columns=['job_salesOrderId'])

    # Convert the 'Actual Start' column to datetime format, coercing errors (invalid parsing will result in NaT)
    df_api_copy['Actual Start'] = pd.to_datetime(df_api_copy['Actual Start'], errors='coerce')
    # Format the 'Actual Start' datetime values into a string with the format 'DD/MM/YYYY HH:MM:SS'
    df_api_copy['Actual Start'] = df_api_copy['Actual Start'].dt.strftime('%d/%m/%Y %H:%M:%S')

    # Convert the 'Actual End' column to datetime format, coercing errors (invalid parsing will result in NaT)
    df_api_copy['Actual End'] = pd.to_datetime(df_api_copy['Actual End'], errors='coerce')
    # Format the 'Actual End' datetime values into a string with the format 'DD/MM/YYYY HH:MM:SS'
    df_api_copy['Actual End'] = df_api_copy['Actual End'].dt.strftime('%d/%m/%Y %H:%M:%S')

    # Convert the 'Deliver Due Date' column to datetime format, coercing errors (invalid parsing will result in NaT)
    df_api_copy['Delivery Due Date'] = pd.to_datetime(df_api_copy['Delivery Due Date'], errors='coerce')

    # Make a copy of the operation dataframe
    df_op_copy = df_op.copy()
    # Make a copy of the job dataframe
    df_job_copy = df_job.copy()
    # Drop unnecessary columns that are no longer required for further analysis or processing
    df_op_copy = df_op_copy.drop(columns=['Customer PO', 'Planned Setup Hours',
                                          'Actual Setup Hours', 'Planned Labor Hours',
                                          'Actual Labor Hours', 'Planned Machine Hours',
                                          'Actual Machine Hours', 'Ready To Collect From Previous Operation',
                                          'Quantity Collected From Previous Operation',
                                          'Scheduled Department', 'Scheduled Equipment', 'Scheduled Work Center'])
    # Change the name of the status column
    df_job_copy['Job Status'] = df_job_copy['Status']
    # Drop unnecessary columns that are no longer required for further analysis or processing
    df_job_copy = df_job_copy.drop(columns=['Job Item', 'Customer', 'Current Item', 'Activity Date', 'User',
                                            'Operation', 'Latest Activity', 'Job Est. Completion',
                                            'Production Due Date',
                                            'Job Item Description', 'Planned Quantity', 'Quantity Completed',
                                            'Sales Order', 'Customer PO', 'Current Item Description', 'Log Type',
                                            'Status'])

    # Merge operation dataframe with job dataframe
    df_op_copy1 = pd.merge(df_op_copy, df_job_copy, on='Job', how='left')

    # Concatenate the two DataFrames (df_api_copy and df_op_copy1) into one DataFrame and reset the index
    df_combined = pd.concat([df_api_copy, df_op_copy1], ignore_index=True)

    # Use regex to remove non-numeric characters from "Sales Order" and "Job"
    df_combined["Sales Order"] = df_combined["Sales Order"].astype(str).str.extract(r'(\d+)').astype(float)
    df_combined["Job"] = df_combined["Job"].astype(str).str.extract(r'(\d+)').astype(float)

    # Sort the combined DataFrame by the "Job" column to ensure all data is ordered by job numbers
    df_combined = df_combined.sort_values(by="Job")

    # Enable future behavior for silent downcasting warnings globally in pandas
    pd.set_option('future.no_silent_downcasting', True)
    # Forward and backward fill missing values for 'Date Created' within each 'Job' group
    df_combined['Date Created'] = (df_combined.groupby('Job')['Date Created']
                                   .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Delivery Due Date' within each 'Job' group
    df_combined['Delivery Due Date'] = (df_combined.groupby('Job')['Delivery Due Date']
                                        .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Make Item' within each 'Job' group
    df_combined['Make Item'] = (df_combined.groupby('Job')['Make Item']
                                .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Customer' within each 'Job' group
    df_combined['Customer'] = (df_combined.groupby('Job')['Customer']
                               .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Production Due Date' within each 'Job' group
    df_combined['Production Due Date'] = (df_combined.groupby('Job')['Production Due Date']
                                          .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Planned Quantity' within each 'Job' group
    df_combined['Planned Quantity'] = (df_combined.groupby('Job')['Planned Quantity']
                                       .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Job Est. Completion' within each 'Job' group
    df_combined['Job Est. Completion'] = (df_combined.groupby('Job')['Job Est. Completion']
                                          .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))
    # Forward and backward fill missing values for 'Job Status' within each 'Job' group
    df_combined['Job Status'] = (df_combined.groupby('Job')['Job Status']
                                 .transform(lambda x: x.ffill().bfill()).infer_objects(copy=False))

    # Fill missing 'Quantity Completed' values with the 'Planned Quantity' value
    df_combined['Quantity Completed'] = df_combined['Quantity Completed'].fillna(df_combined['Planned Quantity'])
    # Fill missing 'Start' values with 'Actual Start' or fallback to 'Scheduled Start'
    df_combined['Start'] = df_combined['Actual Start'].fillna(df_combined['Scheduled Start'])
    # Fill missing 'End' values with 'Actual End' or fallback to 'Scheduled End'
    df_combined['End'] = df_combined['Actual End'].fillna(df_combined['Scheduled End'])

    # Combine 'Quantity Completed' and 'Planned Quantity' as strings and create a new column 'Qty'
    # in the format "Completed/Planned"
    df_combined["Qty"] = (
            df_combined["Quantity Completed"].astype(str) + "/" + df_combined["Planned Quantity"].astype(str)
    )

    # Replace any occurrences of 'nan' in the 'Qty' column with an empty string
    df_combined["Qty"] = df_combined["Qty"].str.replace('nan', '')

    # Combine 'Sales Order' and 'Job' as strings and create a new column 'SO/WO' # in the format 'Sale Order'/'Job'
    df_combined["SO/WO"] = "SO" + df_combined["Sales Order"].astype('Int64').astype(str) + "/" + "WO" + df_combined[
        "Job"].astype('Int64').astype(str)
    # Drop empty rows if a 'Sale Order' column is empty
    df_combined.dropna(subset=['Sales Order'], inplace=True)
    # Define the desired column order for the DataFrame
    desired_order = [
        'SO/WO',
        'Job',
        'Job Status',
        'Date Created',
        'Delivery Due Date',
        'Production Due Date',
        'Job Est. Completion',
        'Job Item',
        'Job Item Description',
        'Make Item',
        'Step',
        'Operation',
        'Status',
        'Actual Start',
        'Actual End',
        'Scheduled Start',
        'Scheduled End',
        'Start',
        'End',
        'Qty'
    ]

    # Reorder the DataFrame columns based on the defined desired order
    df = df_combined[desired_order]

    return df


def generate_gantt_chart(df):
    # Set the timezone to Asia/Kuala_Lumpur
    timezone = pytz.timezone('Asia/Kuala_Lumpur')

    # Sort the DataFrame by Delivery Due Date
    sorted_so_wo = df.sort_values(by='Delivery Due Date')['SO/WO'].tolist()

    # First, ensure all relevant columns are converted to datetime format
    df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
    df['Delivery Due Date'] = pd.to_datetime(df['Delivery Due Date'], format='%Y/%m/%d', errors='coerce')
    df['Job Est. Completion'] = pd.to_datetime(df['Job Est. Completion'], errors='coerce')
    df['Start'] = pd.to_datetime(df['Start'], dayfirst=True, errors='coerce')
    df['End'] = pd.to_datetime(df['End'], dayfirst=True, errors='coerce')

    # Define a function to check tz-naive or tz-aware and localize/convert
    def localize_or_convert_timezone(series, timezone):
        if series.dt.tz is None:
            return series.dt.tz_localize(timezone, nonexistent='shift_forward', ambiguous='NaT')
        else:
            return series.dt.tz_convert(timezone)

    # Now safely apply the function to your columns
    df['Date Created'] = localize_or_convert_timezone(df['Date Created'], 'Asia/Kuala_Lumpur')
    df['Delivery Due Date'] = localize_or_convert_timezone(df['Delivery Due Date'], 'Asia/Kuala_Lumpur')
    df['Job Est. Completion'] = localize_or_convert_timezone(df['Job Est. Completion'], 'Asia/Kuala_Lumpur')
    df['Start'] = localize_or_convert_timezone(df['Start'], 'Asia/Kuala_Lumpur')
    df['End'] = localize_or_convert_timezone(df['End'], 'Asia/Kuala_Lumpur')

    # Dataframe to determine the color of the job depending on if it's late or not
    # The color is stored as a string to be used later
    df['Job Situation'] = df.apply(
        # If the row Job Est. Completion and Delivery Due Date are not NA
        # and if Job Est. Completion is later than sales _deliveryDueDate then color the timeline in darkred
        lambda row: "Late" if pd.notna(row['Job Est. Completion'])
                              and pd.notna(row['Delivery Due Date'])
                              and row['Job Est. Completion'] > row['Delivery Due Date']

        # Else if the row Job Est. Completion & Delivery Due Date are not NA then color timeline in green
        else "On Time",  # Else color in grey

        axis=1
    )

    df['job_color'] = df.apply(
        # If the row Job Est. Completion and Delivery Due Date are not NA
        # and if Job Est. Completion is later than sales _deliveryDueDate then color the timeline in darkred
        lambda row: "darkred" if pd.notna(row['Job Est. Completion'])
                                 and pd.notna(row['Delivery Due Date'])
                                 and row['Job Est. Completion'] > row['Delivery Due Date']

        # Else if the row Job Est. Completion & Delivery Due Date are not NA then color timeline in green
        else ("green" if pd.notna(row['Job Est. Completion'])
                         and pd.notna(row['Delivery Due Date']) else 'grey'),  # Else color in grey

        axis=1
    )

    # Dataframe to identify the operation status (Running, Paused, complete, Pending & cancelled)
    df['operation_status_color'] = df.apply(
        lambda row: "orange" if row['Operation'] in ['Threading and Surface Coating', 'Threading', 'Surface Coating']
        else "grey" if row['Status'] == 'Paused'
        else "blue" if row['Status'] == 'complete'
        else "purple" if row['Status'] == 'Pending'
        else "cyan" if row['Status'] == 'Ready'
        else "yellowgreen" if row['Status'] == 'Running'
        else 'black',

        axis=1
    )

    # Dataframe to have the line on job status in another opacity
    df['operation_line_color'] = df.apply(
        lambda row: "rgba(255, 87, 51, 0.5)" if row['Operation'] in ['Threading and Surface Coating', 'Threading',
                                                                     'Surface Coating', 'Induction Hardening', 'Xylan']
        else "rgba(61, 61, 61, 0.5)" if row['Status'] == 'Paused'
        else "rgba(89, 136, 255, 0.5)" if row['Status'] == 'complete'
        else "rgba(128, 0, 128, 0.5)" if row['Status'] == 'Pending'
        else "rgba(141, 228, 255, 0.5)" if row['Status'] == 'Ready'
        else "rgba(80, 247, 0, 0.7)" if row['Status'] == 'Running'
        else 'black',

        axis=1
    )

    # Create a new column for the maximum date
    df['Max end date'] = df[['Delivery Due Date', 'Job Est. Completion']].max(axis=1)

    # Convert dataframe column 'Job' to string format
    df['Job'] = df['Job'].astype(str)

    # Save current date and time
    today = datetime.now(timezone)

    # Create a plotly figure
    fig = px.timeline(
        # Identify the dataframe we wish to use
        df,
        # Assign the start on the x-axis
        x_start="Date Created",
        # Assign the end on the x-axis
        x_end="Max end date",
        # Assign the y-axis
        y="SO/WO",
        # Assign the color using dataframe column 'job_color' created earlier
        color='Job Situation',
        color_discrete_map={
            'Late': 'rgba(235, 189, 189, 0.8)',
            'On Time': 'rgba(189, 255, 190, 0.8)'
        },
        # Select the info we want when hovering over the timeline
        hover_data=[
            "Operation"
        ],
        # Create a title
        title=f"Job Progress Timeline - {today}",
        # Assign the custom_colors to color_discrete_sequence
    )

    # Add markers for scheduled start and end dates
    for _, row in df.iterrows():
        job_color = row['job_color']
        operation_status_color = row['operation_status_color']
        operation_line_color = row['operation_line_color']

        # Add marker for Scheduled Due Date
        fig.add_trace(go.Scatter(
            x=[row['Job Est. Completion']],
            y=[row['SO/WO']],
            mode='markers',
            marker=dict(color='orange', size=15, symbol='circle'),
            showlegend=False,
            hoverinfo='text',
            text=f"Scheduled Due Date: {row['Job Est. Completion']}"
        ))
        # Add line+markers for operation details
        fig.add_trace(go.Scatter(
            x=[row['Start'], row['End']],
            y=[row['SO/WO'], row['SO/WO']],
            mode='lines+markers',
            line=dict(color=operation_line_color, width=10),
            marker=dict(color=operation_status_color, size=10),
            showlegend=False,
            hoverinfo='text',
            text=(f"Operation: {row['Operation']}<br>"
                  f"Start:{row['Start']}, End:{row['End']}<br>"
                  f"Status:{row['Status']}<br>"
                  f"Job Number: {row['SO/WO']}")
        ))
        # Add marker for Date Created
        fig.add_trace(go.Scatter(
            x=[row['Date Created']],
            y=[row['SO/WO']],
            mode='markers',
            marker=dict(color='purple', size=15, symbol='circle'),
            showlegend=False,
            hoverinfo='text',
            text=(
                f"Created: {row['Date Created']}<br>"
                f"Job Number: {row['SO/WO']}<br>"
                f"Item Number: {row['Job Item']}<br>"
                f"Item Description: {row['Job Item Description']}"
            )
        ))
        # Add marker for Production Due Date
        fig.add_trace(go.Scatter(
            x=[row['Delivery Due Date']],
            y=[row['SO/WO']],
            mode='markers',
            marker=dict(color=job_color, size=15, symbol='circle'),
            showlegend=False,
            hoverinfo='text',
            text=f"Production Due Date: {row['Production Due Date']},Job Number: {row['SO/WO']}"
        ))

    # Add a logo to the top of the graph
    fig.add_layout_image(
        dict(
            source="https://i.postimg.cc/4yLKW7PF/logo-grey-full.png",  # URL or path to your logo image
            x=1.0,
            y=1.15,
            sizex=0.15,  # Width of the image (relative to the plot)
            sizey=0.15,  # Height of the image (relative to the plot)
            xanchor="center",
            yanchor="top",
            opacity=1,
        )
    )

    # Add name at the very bottom
    fig.update_layout(
        annotations=[
            go.layout.Annotation(
                text="Ahmed Montasser",
                xref="paper",
                yref="paper",
                x=1,
                y=-0.1,
                showarrow=False,
                font=dict(size=12, color="black"),
                align="right"
            )
        ]
    )

    fig.add_vline(x=today, line_width=2, line_dash="dash", line_color='red')

    # Update y-axis order to follow Delivery Due Date order
    fig.update_yaxes(categoryorder="array", categoryarray=sorted_so_wo[::-1])

    # Update the layout to have the title 'Date' on the x-axis and 'Job Number' on y-axis
    fig.update_layout(xaxis_title="Date", yaxis_title="Job Number")

    fig.show()
