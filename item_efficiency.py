import pandas as pd
import plotly.express as px
from tkinter import filedialog, messagebox

def item_extract_efficiency(df):
    # Debug: Print the columns to check if all required columns are present
    print("DataFrame columns:", df.columns)

    # Check if all required columns are in the DataFrame
    required_columns = [
        'Actual Total Hours', 'Estimated Total Hours',
        'Estimated Make Quantity', 'Job', 'Make Item',
        'Make Item Description'
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise ValueError(f"Missing columns: {', '.join(missing_columns)}")

    df = df[df['Actual Total Hours'] != 0].copy()

    df['Estimated Total Hours'] = df['Estimated Total Hours'] / df['Estimated Make Quantity']
    df['Actual Total Hours'] = df['Actual Total Hours'] / df['Estimated Make Quantity']

    categories = {
        'MOE': ['millout', 'moe'],
        'TBR': ['tbr', 'ext', 'tieback',  'extension', 'extensions'],
        'BODY': ['body', 'bodies', ],
        'MANDREL': ['mandrel', 'mandrels', 'mandrel', 'mandrels', 'mandrels', 'mandrels', 'mndrel', 'mdrl', 'MNDRL'],
        'PBR': ['PBR', 'polished', 'bore'],
        'FLOW COUPLING': ['flow coupling'],
        'BLAST JOINT': ['blast joint'],
        'CROSSOVER': ['crossover', 'crossovers', 'crossovers', 'crossovers', 'crossovers', 'crossovers', 'xover',
                      'sub']
    }

    def categorize(description):
        if isinstance(description, str):
          for category, keywords in categories.items():
              for keyword in keywords:
                  if keyword.lower() in description.lower():
                      return category
          return 'OTHER CATEGORY'

    df['Item Category'] = df['Make Item Description'].apply(categorize)

    df['Estimated Time'] = df.groupby(['Job'])['Estimated Total Hours'].transform('sum')
    df['Actual Time'] = df.groupby(['Job'])['Actual Total Hours'].transform('sum')

    df['Mean Estimated Time'] = df.groupby(['Make Item'])['Estimated Time'].transform('mean').round(2)
    df['Mean Actual Time'] = df.groupby(['Job'])['Actual Time'].transform('mean').round(2)

    df['Efficiency'] = (df['Estimated Time'] / df['Actual Time']) * 100
    df = df[df['Efficiency'] < 100].copy()

    df['Mean Efficiency'] = df.groupby(['Make Item'])['Efficiency'].transform('mean').round(2)

    desired_data = [
        'Make Item',
        'Item Category',
        'Make Item Description',
        'Mean Estimated Time',
        'Mean Actual Time',
        'Mean Efficiency'
    ]

    df_item = df[desired_data].drop_duplicates(subset=['Make Item'], keep='first')

    df_item = df_item.sort_values(by='Item Category', ascending=False)

    return df_item


def item_graph_efficiency(df):
    try:
      df_item = item_extract_efficiency(df)

      fig = px.bar(df_item,
                   x='Make Item',
                   y='Mean Efficiency',
                   color='Make Item',
                   title='Average Item Efficiency',
                   text='Mean Efficiency',
                   hover_data={'Make Item': True,
                               'Make Item Description': True,
                               'Item Category': True})
      fig.update_layout(showlegend=False)
      fig.show()
    except Exception as e:
        print(e)

def save_to_csv(df_item, file_path):
    try:
        df_copy_1 = item_extract_efficiency(df_item)
        print(df_copy_1)

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df_copy_1.to_excel(writer, index=False)

        print(f"Summary file created at {file_path}")
    except Exception as e:
        print(e)