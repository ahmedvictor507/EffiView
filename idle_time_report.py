import pandas as pd
import datetime
import plotly.express as px


def extract_idle_data(df_copy):
    df_copy = df_copy[~df_copy['Clock Type'].str.contains('ClockIn')]
    df_copy = df_copy.dropna(subset=['Operation'], inplace=False)
    # Change the 'Started On Date' to datetime format
    df_copy['Started On Date'] = pd.to_datetime(df_copy['Started On Date'], errors='coerce')
    df_copy['Stopped On Date'] = pd.to_datetime(df_copy['Stopped On Date'], errors='coerce')
    if df_copy['Started On Date'].isna().any():
        print("Warning: Some 'Started On Date' values could not be converted to datetime.")

    # Ensure 'Labor Hours' is numeric
    df_copy['Labor Hours'] = pd.to_numeric(df_copy['Labor Hours'], errors='coerce')
    if df_copy['Labor Hours'].isna().any():
        print("Warning: Some 'Labor Hours' values could not be converted to numeric.")

    # Create a new column called 'Total Hours Per Day' group based on 'Team Member' & 'Started On Date' and sum up
    # the 'Labor Hours'
    # Unlike agg() or sum(), which would collapse the DataFrame into a smaller result, transform() returns a new series
    # that has the same size as the original DataFrame.
    # This means that for each row in df_copy, the value of Total Hours Per Day will be the same for all rows within
    # the same group (i.e., for the same Team Member and Started On Date).
    df_copy['Total Hours Per Day'] = df_copy.groupby(['Team Member', 'Started On Date'])['Duration Hours'].transform(
        'sum')

    # Create a new dataframe to group data for each unique combination of 'Started On Date' & 'Team Member'
    # and then calculate the sum for 'Labor Hours'
    df_total_hours = df_copy.groupby(['Started On Date', 'Team Member'])['Duration Hours'].sum().reset_index()

    df_copy.merge(df_total_hours, how='left', on='Started On Date')
    if df_copy is None:
        print("Error: df_copy after merge is None")
        return None

    # Filter out data that have a Labor Time less than 1 minute
    df_copy = df_copy[df_copy['Duration Hours'] >= 0.01]

    # Convert to datetime using (to_datetime) and then time format using (dt.time)
    df_copy.loc[:, 'Started On Time'] = pd.to_datetime(df_copy['Started On Time'], format='%I:%M:%S %p',
                                                       errors='coerce').dt.time
    df_copy.loc[:, 'Stopped On Time'] = pd.to_datetime(df_copy['Stopped On Time'], format='%I:%M:%S %p',
                                                       errors='coerce').dt.time

    # Store working hours during the day and night shifts
    day_start = datetime.datetime.strptime('07:45:00 am', '%I:%M:%S %p').time()
    day_end = datetime.datetime.strptime('06:15:00 pm', '%I:%M:%S %p').time()
    night_start = datetime.datetime.strptime('07:45:00 pm', '%I:%M:%S %p').time()
    night_end = datetime.datetime.strptime('06:15:00 am', '%I:%M:%S %p').time()

    # Store the value of working hours for overtime
    normal_working_hours = 8.5
    overtime_working_hours = 10.5

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
                        'Ahmed Yasser Montasser', 'RABDUL  HAKIM BIN ISMAIL', 'MAICHAEL ANTONIRAJ',
                        'Sawinder Singh', 'Nur Syafiq  Bin Kamal']

    # Remove leading and trailing whitespace from the 'Employee' column
    # If you want to do the opposite, create names_to_include and delete '~' from this line
    df_copy = df_copy[~df_copy['Team Member'].isin(names_to_exclude)]

    # Create a function to check whether the Team Member has worked overtime or not
    def check_overtime(start_time, end_time):

        # If data is NaN then it should be marked as "Faulty Data"
        if pd.isna(start_time):
            return 'FAULTY DATA'
        if pd.isna(end_time):
            return 'FAULTY DATA'

        if night_end < start_time < day_start or night_end <= end_time <= day_start:
            return 'YES'
        elif day_end < start_time < night_start or day_end <= end_time:
            return 'YES'
        if day_start <= start_time <= day_end or day_start <= end_time <= day_end:
            return 'NO'
        elif night_start <= start_time <= night_end or night_start <= end_time <= night_end:
            return 'NO'
        else:
            return 'FAULTY DATA'

    # Function to mark all overtime rows of the same "Team Member" and "Started On Date" as "YES"
    def apply_overtime(group):
        group['Overtime'] = group.apply(lambda row: check_overtime(row['Started On Time'], row['Stopped On Time']),
                                        axis=1)

        # If any row has 'YES' for overtime, set all rows in this group to 'YES'
        if 'YES' in group['Overtime'].values:
            group['Overtime'] = 'YES'

        return group

    # Apply the (apply_overtime) function on df_copy
    df_copy = df_copy.groupby(['Team Member', 'Started On Date']).apply(apply_overtime).reset_index(drop=True)

    # Create a function to calculate the idle time
    def calculate_idle_time(row):
        overtime_status = row['Overtime']
        total_hours = row['Total Hours Per Day']
        if overtime_status == 'YES':
            return overtime_working_hours - total_hours
        elif overtime_status == 'NO':
            return normal_working_hours - total_hours

    # You can store any special days here
    special_days = {
        'MUHAMMAD USAMA  KHAN': {
            pd.to_datetime('2024-09-12 00:00:00'): '1/2 day',
            pd.to_datetime('2024-09-10 00:00:00'): 'Medical Leave',
            pd.to_datetime('2024-09-19 00:00:00'): 'Medical Leave'
        },
        'YEN KONG CHIN': {
            pd.to_datetime('2024-09-11 00:00:00'): 'Medical Leave',
        },
    }

    # Create a function to get Feedback based on efficiency
    def feedback(row):

        if pd.isna(row['Started On Date']):
            return 'Unkown'

        efficiency = row['Idle Time Percentage']
        team_member = row['Team Member']
        started_on = row['Started On Date']

        if isinstance(started_on, str):
            try:
                started_on = pd.to_datetime(started_on)
            except ValueError:
                return 'Unknown'

        if team_member in special_days:
            date_mapping = special_days[team_member]
            if started_on in date_mapping:
                return date_mapping[started_on]

        if efficiency < -10:
            return 'Left Time Running'
        elif -10 < efficiency < 10:
            return 'Good Time'
        elif efficiency > 10:
            return 'Too Much Idle Time'

        return 'Unknown'

    # Create a function to calculate the total time based on the overtime_status
    def total_time(row):
        overtime_status = row['Overtime']
        if overtime_status == 'YES':
            return 10.5
        elif overtime_status == 'NO':
            return 8.5

    # Apply the (calculate_idle_time) function on a new column "Idle Hours"
    df_copy.loc[:, 'Idle Hours'] = df_copy.apply(calculate_idle_time, axis=1)

    # Apply the (total_time) function to a new column "Set Working Hours"
    df_copy.loc[:, 'Set Working Hours'] = df_copy.apply(total_time, axis=1)

    # Calculate the Idle Time Percentage
    df_copy['Idle Time Percentage'] = (df_copy['Idle Hours'] / df_copy['Set Working Hours']) * 100

    # Create a new column and apply (feedback) function
    df_copy.loc[:, 'Feedback'] = df_copy.apply(feedback, axis=1)

    # Round columns to 2 decimal places
    df_copy['Duration Hours'] = df_copy['Duration Hours'].round(decimals=2)
    df_copy['Total Hours Per Day'] = df_copy['Total Hours Per Day'].round(decimals=2)
    df_copy['Idle Hours'] = df_copy['Idle Hours'].round(decimals=2)
    df_copy['Set Working Hours'] = df_copy['Set Working Hours'].round(decimals=2)
    df_copy['Idle Time Percentage'] = df_copy['Idle Time Percentage'].round(decimals=2)
    df_copy['Idle Time Percentage (%)'] = df_copy['Idle Time Percentage'].round(decimals=2).astype(str) + '%'

    # Create desired_data to store only the columns needed
    desired_data = [
        'Job',
        'Operation',
        'Team Member',
        'Started On Date',
        'Started On Time',
        'Started On Week',
        'Stopped On Date',
        'Stopped On Time',
        'Duration Hours',
        'Total Hours Per Day',
        'Overtime',
        'Idle Hours',
        'Set Working Hours',
        'Idle Time Percentage',
        'Feedback'
    ]

    # Sort the values by Started On Date and Team Member name
    df_copy = df_copy.sort_values(by=['Started On Date', 'Team Member'], ascending=False)

    # Assign the desired_data to filter out the columns needed
    df_copy = df_copy[desired_data]
    df_copy.head(100)

    return df_copy


def create_summary(df_copy, file_path):
    try:
        # Extract idle data from the original DataFrame and create a copy
        df_copy_1 = extract_idle_data(df_copy)
        print(df_copy_1)

        # Create a new function to give feedback
        def feedback(row):

            efficiency = row['Idle Time Percentage']

            if efficiency < -10:
                return 'Left Time Running'
            elif -10 < efficiency < 10:
                return 'Good Time'
            elif efficiency > 10:
                return 'Too Much Idle Time'

            return 'Unknown'

        feedback_messages = {
            'Left Time Running': 'Leaves timer running usually',
            'Good Time': 'Great job',
            'Too Much Idle Time': 'Try to reduce idle time',
            'Unknown': 'Unknown issue, needs investigation'
        }

        # Create a function to get the Overall Feedback
        def overall_feedback(group):
            # Count the occurrences of each feedback
            feedback_counts = group['Feedback'].value_counts()

            # Find the most frequent feedback count
            max_count = feedback_counts.max()

            # Identify the dominant feedback(s) (if multiple have the same max count)
            dominant_feedbacks = feedback_counts[feedback_counts == max_count].index.tolist()

            # If there is more than one feedback with the same max count, return 'Needs investigation'
            if len(dominant_feedbacks) > 1:
                return "Needs Investigation"

            # If only one dominant feedback, return the custom message for that feedback
            dominant_feedback = dominant_feedbacks[0]
            return feedback_messages.get(dominant_feedback, 'Mixed Feedback')

        # Daily Summary for each Team Member ##
        # Create a new dataframe to store a summary each Team Members idle time for each day
        df_daily_feedback = df_copy_1.drop(
            columns=['Started On Time', 'Stopped On Time', 'Duration Hours', 'Job', 'Operation', 'Stopped On Date'])

        feedback_data_daily = df_daily_feedback.groupby(['Team Member']).apply(
            overall_feedback).reset_index(name='Daily Overall Feedback')

        df_daily_feedback = pd.merge(df_daily_feedback, feedback_data_daily, on=['Team Member'],
                              how='left')

        # Drop the duplicates from the dataframe
        df_daily_feedback = df_daily_feedback.drop_duplicates()
        #########################

        # Weekly Feedback for each Team Member ##
        df_weekly_feedback = df_copy_1.drop(columns=['Job',
                                                     'Operation',
                                                     'Stopped On Time',
                                                     'Duration Hours',
                                                     'Overtime',
                                                     ])

        # Remove duplicate entries based on 'Team Member' and 'Started On Date' columns
        df_weekly_feedback = df_weekly_feedback.drop_duplicates(subset=['Team Member', 'Started On Date'])

        df_weekly_feedback['Total Idle Hours/Week'] = df_copy_1['Idle Hours']
        df_weekly_feedback['Total Hours/Week'] = df_copy_1['Total Hours Per Day']
        df_weekly_feedback['Set Working Hours/Week'] = df_copy_1['Set Working Hours']

        # Group the DataFrame by 'Team Member' and 'Started On Week', summing specified columns for aggregation
        df_weekly_feedback = df_weekly_feedback.groupby(['Team Member', 'Started On Week']).agg({
            'Total Idle Hours/Week': 'sum',
            'Total Hours/Week': 'sum',
            'Set Working Hours/Week': 'sum'
        }).reset_index()

        # Calculate the percentage of idle time for each team member based on total hours worked in a week
        df_weekly_feedback['Idle Time Percentage'] = (df_weekly_feedback['Total Idle Hours/Week'] / df_weekly_feedback[
            'Total Hours/Week']) * 100

        # Apply the feedback function to each row in the DataFrame to generate feedback based on the data
        df_weekly_feedback.loc[:, 'Feedback'] = df_weekly_feedback.apply(feedback, axis=1)

        # Group by 'Team Member' and 'Started On Week' to calculate overall feedback
        feedback_data_weekly = df_weekly_feedback.groupby(['Team Member']).apply(
            overall_feedback).reset_index(name='Weekly Overall Feedback')

        # Merge the weekly feedback DataFrame with additional feedback data based on 'Team Member'
        df_weekly_feedback = pd.merge(df_weekly_feedback, feedback_data_weekly, on=['Team Member'],
                                      how='left')

        # Define the desired column and their order for the DataFrame
        desired_weekly_columns = [
            'Team Member',
            'Started On Week',
            'Total Hours/Week',
            'Total Idle Hours/Week',
            'Set Working Hours/Week',
            'Feedback',
            'Weekly Overall Feedback'
        ]

        # Select only the columns specified in 'desired_data_em'
        df_weekly_feedback = df_weekly_feedback[desired_weekly_columns]

        print(df_weekly_feedback)
        #########################

        # Summary for each Team Member ##
        # Drop unnecessary columns from the summary DataFrame to focus on relevant data
        df_summary_total = df_daily_feedback.drop(columns=['Feedback', 'Started On Date', 'Overtime', 'Total Hours Per Day'])
        # Create a new column 'Total Hours' by copying the values from 'Total Hours Per Day'
        df_summary_total['Total Hours'] = df_daily_feedback['Total Hours Per Day']

        # Group the DataFrame by 'Team Member' and aggregate total hours, idle hours, and set working hours
        df_summary_total = df_summary_total.groupby(['Team Member']).agg({
            'Total Hours': 'sum',
            'Idle Hours': 'sum',
            'Set Working Hours': 'sum',
        }).reset_index()

        # Calculate the percentage of idle time relative to total hours for each team member
        df_summary_total['Idle Time Percentage'] = (df_summary_total['Idle Hours'] / df_summary_total[
            'Total Hours']) * 100

        # Create a new column and apply (feedback) function
        df_summary_total.loc[:, ' Total Time based Feedback'] = df_summary_total.apply(feedback, axis=1)

        # Merge the summary total DataFrame with feedback data on the 'Team Member' column
        df_summary_total = pd.merge(df_summary_total, feedback_data_weekly, on=['Team Member'])
        df_summary_total = pd.merge(df_summary_total, feedback_data_daily, on=['Team Member'])

        # Remove duplicate rows from the summary total DataFrame to ensure each entry is unique
        df_summary_total = df_summary_total.drop_duplicates()
        print(df_summary_total)
        #########################

        # Create an Excel writer object using XlsxWriter as the engine
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            # Write the extensive details DataFrame to the first sheet
            df_copy_1.to_excel(writer, sheet_name='Extensive Details', index=False)
            # Write the daily summary DataFrame to the second sheet
            df_daily_feedback.to_excel(writer, sheet_name='Daily Summary', index=False)
            # Write the weekly feedback DataFrame to the third sheet
            df_weekly_feedback.to_excel(writer, sheet_name='Weekly Summary', index=False)
            # Write the overall summary DataFrame to the fourth sheet
            df_summary_total.to_excel(writer, sheet_name='Overall Summary', index=False)

        print(f"Summary file created at {file_path}")
    except Exception as e:
        print(e)


def generate_mean_chart(df_copy):
    try:
        df_copy_1 = extract_idle_data(df_copy)

        earliest_date = df_copy_1['Stopped On Date'].min()
        latest_date = df_copy_1['Started On Date'].max()
        # Create a new dataframe to calculate the mean
        # Group by "Team Member" and calculate the mean
        df_mean = df_copy_1.groupby(['Team Member']).agg({
            'Total Hours Per Day': 'mean',
            'Idle Hours': 'mean',
            'Idle Time Percentage': 'mean'
        }).reset_index()

        # Function to categorize idle time efficiency feedback
        def feedback(row):
            efficiency = row['Idle Time Percentage']
            if efficiency < -10:
                return 'Left Time Running'
            elif -10 < efficiency < 10:
                return 'Good Time'
            elif efficiency > 10:
                return 'Too Much Idle Time'

        # Round the means from 2 decimal places
        df_mean['Mean Total Hours Per Day'] = df_mean['Total Hours Per Day'].round(2)
        df_mean['Mean Idle Hours'] = df_mean['Idle Hours'].round(2)
        df_mean['Mean Idle Time (%)'] = df_mean['Idle Time Percentage'].round(2)
        df_mean['Feedback'] = df_mean.apply(feedback, axis=1)

        df_mean['custom_text'] = (
                df_mean['Team Member'] +
                '<br> Average Idle Time Percentage: ' + df_mean['Mean Idle Time (%)'].astype(str) + "%" +
                '<br> Average Idle Hours: ' + df_mean['Mean Idle Hours'].astype(str) +
                '<br> Feedback: ' + df_mean['Feedback']
        )

        fig = px.bar(df_mean,
                     x='Team Member',
                     y='Mean Idle Time (%)',
                     title=f'Mean Idle Time Percentage from {earliest_date.date()} to {latest_date.date()}',
                     color='Feedback',
                     color_discrete_map={
                         'Good Time': '#2CA02C',
                         'Left Time Running': '#FF9900',
                         'Too Much Idle Time': '#B82E2E'},
                     text='Mean Idle Time (%)',
                     custom_data=['custom_text']
                     )
        fig.update_traces(
            hovertemplate='%{customdata[0]}<extra></extra>'
        )
        fig.show()
    except Exception as e:
        print(e)
