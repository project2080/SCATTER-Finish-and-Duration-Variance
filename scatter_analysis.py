"""
SCATTER Finish and Duration Variance Analysis Tool

This module provides functions to analyze and visualize project activities based on their
finish date variance, duration variance, and total float. The main visualization is an
interactive scatter plot that categorizes activities by their total float values and
positions them in quadrants based on their variance characteristics.

Main functions:
- generate_excel_by_categories: Creates an Excel file with activities grouped by total float
- create_interactive_scatter_plot: Generates an interactive scatter plot visualization

The tool helps project managers identify:
1. Critical activities (Total Float ≤ 0) that require immediate attention
2. Activities performing better or worse than baseline in terms of schedule and duration
3. Distribution of activities across different performance quadrants
"""

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

def generate_excel_by_categories(filtered_df, filename='activities_by_category.xlsx'):
    """
    Generates an Excel file with 4 sheets, one for each Total Float category.
    
    Parameters:
    filtered_df (pandas.DataFrame): Filtered DataFrame with activities.
    filename (str): Name of the Excel file to generate.
    
    Returns:
    str: Path to the generated Excel file.
    """
    # Create an Excel Writer to save the DataFrame in multiple sheets
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Filter and save each category in a separate sheet
        
        # Total Float ≤ 0
        df_tf_0 = filtered_df[filtered_df['Total Float'] <= 0]
        df_tf_0.to_excel(writer, sheet_name='Total Float ≤ 0', index=False)
        
        # 0 < Total Float ≤ 10
        df_tf_0_10 = filtered_df[(filtered_df['Total Float'] > 0) & (filtered_df['Total Float'] <= 10)]
        df_tf_0_10.to_excel(writer, sheet_name='0 < Total Float ≤ 10', index=False)
        
        # 10 < Total Float ≤ 20
        df_tf_10_20 = filtered_df[(filtered_df['Total Float'] > 10) & (filtered_df['Total Float'] <= 20)]
        df_tf_10_20.to_excel(writer, sheet_name='10 < Total Float ≤ 20', index=False)
        
        # Total Float > 20
        df_tf_20 = filtered_df[filtered_df['Total Float'] > 20]
        df_tf_20.to_excel(writer, sheet_name='Total Float > 20', index=False)
    
    # Get the absolute path to the file
    absolute_path = os.path.abspath(filename)
    print(f"Excel file generated: {absolute_path}")
    
    return absolute_path

def create_interactive_scatter_plot(file, generate_excel=True, excel_name='activities_by_category.xlsx'):
    """
    Create an interactive scatter plot visualization of project activities.
    
    The plot is divided into four quadrants:
    - Q1 (top-right): Earlier and shorter - Optimal performance
    - Q2 (top-left): Later but shorter - Schedule issues despite efficiency
    - Q3 (bottom-right): Earlier but longer - Schedule improvement despite inefficiency
    - Q4 (bottom-left): Later and longer - Problematic performance
    
    Activities are color-coded by their Total Float values:
    - Red: Total Float ≤ 0 (Critical)
    - Orange: 0 < Total Float ≤ 10 (Near-critical)
    - Gold: 10 < Total Float ≤ 20 (Medium float)
    - Green: Total Float > 20 (High float)
    
    Parameters:
    file (str): Path to the Excel file with the data.
    generate_excel (bool): If True, generates an Excel file with activities divided by category.
    excel_name (str): Name of the Excel file to generate.
    
    Returns:
    plotly.graph_objects.Figure: Plotly Figure with the interactive chart.
    dict: Dictionary with basic statistics of the data and the path to the generated Excel if applicable.
    """
    # Read the Excel file
    df = pd.read_excel(file)
    
    # Show general information about the data
    print(f"Total activities in the file: {len(df)}")
    print(f"Available columns: {', '.join(df.columns)}")
    print(f"Unique values of Activity Status: {df['Activity Status'].unique()}")
    
    # Filter activities that are not completed (making an explicit copy)
    filtered_df = df[df['Activity Status'] != 'Completed'].copy()
    print(f"\nNon-completed activities: {len(filtered_df)}")
    
    # Convert columns to numbers (in case they come as strings)
    filtered_df.loc[:, 'Variance - BL Project Finish Date'] = pd.to_numeric(filtered_df['Variance - BL Project Finish Date'], errors='coerce')
    filtered_df.loc[:, 'Variance - BL Project Duration'] = pd.to_numeric(filtered_df['Variance - BL Project Duration'], errors='coerce')
    filtered_df.loc[:, 'Total Float'] = pd.to_numeric(filtered_df['Total Float'], errors='coerce')
    
    # Remove rows with NaN after conversion
    filtered_df = filtered_df.dropna(subset=['Variance - BL Project Finish Date', 'Variance - BL Project Duration', 'Total Float'])
    
    # Show basic statistics of filtered data
    print("\nBasic statistics:")
    print(f"X-axis range (Variance - BL Project Finish Date): {filtered_df['Variance - BL Project Finish Date'].min()} to {filtered_df['Variance - BL Project Finish Date'].max()}")
    print(f"Y-axis range (Variance - BL Project Duration): {filtered_df['Variance - BL Project Duration'].min()} to {filtered_df['Variance - BL Project Duration'].max()}")
    print(f"Total Float range: {filtered_df['Total Float'].min()} to {filtered_df['Total Float'].max()}")
    
    # Create a column to categorize according to Total Float
    def categorize_float(total_float):
        if total_float <= 0:
            return 'Total Float ≤ 0'
        elif total_float <= 10:
            return '0 < Total Float ≤ 10'
        elif total_float <= 20:
            return '10 < Total Float ≤ 20'
        else:
            return 'Total Float > 20'
    
    filtered_df.loc[:, 'Category'] = filtered_df['Total Float'].apply(categorize_float)
    
    # Assign colors to each category
    color_map = {
        'Total Float ≤ 0': 'red',
        '0 < Total Float ≤ 10': 'orange',
        '10 < Total Float ≤ 20': 'gold',  # Using gold instead of yellow for better visibility
        'Total Float > 20': 'green'
    }
    
    # Sort the categories to appear in order in the legend
    category_order = [
        'Total Float ≤ 0',
        '0 < Total Float ≤ 10',
        '10 < Total Float ≤ 20',
        'Total Float > 20'
    ]
    
    # Show distribution by categories
    category_counts = filtered_df['Category'].value_counts()
    print("\nDistribution by categories (Total Float):")
    for cat in category_order:
        print(f"{cat}: {category_counts.get(cat, 0)}")
    
    # Create custom information for tooltips with specific format
    filtered_df['tooltip_text'] = filtered_df.apply(
        lambda row: f"Activity ID: {row['Activity ID']}<br>" +
                   f"Activity Name: {row['Activity Name']}<br>" +
                   f"Activity Status: {row['Activity Status']}<br><br>" +
                   f"Variance - BL Project Duration: {row['Variance - BL Project Duration']}<br>" +
                   f"Variance - BL Project Finish Date: {row['Variance - BL Project Finish Date']}<br>" +
                   f"Total Float: {row['Total Float']}",
        axis=1
    )

    # Create the interactive chart with Plotly using custom_data to fully control the tooltip
    fig = px.scatter(
        filtered_df, 
        x='Variance - BL Project Finish Date',
        y='Variance - BL Project Duration',
        color='Category',
        color_discrete_map=color_map,
        category_orders={'Category': category_order},
        custom_data=['tooltip_text'],  # Use our custom text
        title='Interactive Scatter Plot: Finish Date Variation vs. Duration Variation',
        labels={
            'Variance - BL Project Finish Date': 'Variation - BL Project Finish Date',
            'Variance - BL Project Duration': 'Variation - BL Project Duration'
        },
        height=700,
        opacity=0.7,
    )
    
    # Customize the chart design
    fig.update_traces(
        marker=dict(
            size=10,
            line=dict(width=1, color='DarkSlateGrey')
        ),
        hovertemplate='%{customdata[0]}<extra></extra>'  # Use custom text
    )
    
    # Add quadrant lines
    fig.add_shape(
        type="line", line=dict(dash="dash", color="gray"),
        x0=filtered_df['Variance - BL Project Finish Date'].min(),
        y0=0,
        x1=filtered_df['Variance - BL Project Finish Date'].max(),
        y1=0
    )
    fig.add_shape(
        type="line", line=dict(dash="dash", color="gray"),
        x0=0,
        y0=filtered_df['Variance - BL Project Duration'].min(),
        x1=0,
        y1=filtered_df['Variance - BL Project Duration'].max(),
    )
    
    # Add annotations for quadrants
    fig.add_annotation(
        x=filtered_df['Variance - BL Project Finish Date'].max() * 0.9,
        y=filtered_df['Variance - BL Project Duration'].max() * 0.9,
        text="Earlier and shorter",
        showarrow=False,
        bgcolor="white",
        opacity=0.7,
        bordercolor="black",
        borderwidth=1
    )
    fig.add_annotation(
        x=filtered_df['Variance - BL Project Finish Date'].min() * 0.9,
        y=filtered_df['Variance - BL Project Duration'].max() * 0.9,
        text="Later but shorter",
        showarrow=False,
        bgcolor="white",
        opacity=0.7,
        bordercolor="black",
        borderwidth=1
    )
    fig.add_annotation(
        x=filtered_df['Variance - BL Project Finish Date'].max() * 0.9,
        y=filtered_df['Variance - BL Project Duration'].min() * 0.9,
        text="Earlier but longer",
        showarrow=False,
        bgcolor="white",
        opacity=0.7,
        bordercolor="black",
        borderwidth=1
    )
    fig.add_annotation(
        x=filtered_df['Variance - BL Project Finish Date'].min() * 0.9,
        y=filtered_df['Variance - BL Project Duration'].min() * 0.9,
        text="Later and longer",
        showarrow=False,
        bgcolor="white",
        opacity=0.7,
        bordercolor="black",
        borderwidth=1
    )
    
    # Update layout to improve aesthetics
    fig.update_layout(
        plot_bgcolor='white',
        legend_title_text='Categories by Total Float',
        xaxis=dict(
            showgrid=True,
            gridcolor='lightgray',
            zeroline=True,
            zerolinecolor='gray',
            zerolinewidth=2,
            title_font=dict(size=14, family='Arial', color='black'),
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='lightgray',
            zeroline=True,
            zerolinecolor='gray',
            zerolinewidth=2,
            title_font=dict(size=14, family='Arial', color='black'),
        ),
        title=dict(
            font=dict(size=16, family='Arial', color='black'),
            x=0.5,
            xanchor='center'
        ),
        legend=dict(
            bordercolor='gray',
            borderwidth=1,
            bgcolor='rgba(255, 255, 255, 0.8)'
        ),
        margin=dict(l=20, r=20, t=50, b=20),
    )
    
    # Add text with information about quantities
    info_text = (
        f"<b>Total activities:</b> {len(filtered_df)}<br>"
        f"<b>Total Float ≤ 0:</b> {len(filtered_df[filtered_df['Total Float'] <= 0])}<br>"
        f"<b>0 < Total Float ≤ 10:</b> {len(filtered_df[(filtered_df['Total Float'] > 0) & (filtered_df['Total Float'] <= 10)])}<br>"
        f"<b>10 < Total Float ≤ 20:</b> {len(filtered_df[(filtered_df['Total Float'] > 10) & (filtered_df['Total Float'] <= 20)])}<br>"
        f"<b>Total Float > 20:</b> {len(filtered_df[filtered_df['Total Float'] > 20])}"
    )
    
    fig.add_annotation(
        x=0.02,
        y=0.02,
        xref="paper",
        yref="paper",
        text=info_text,
        showarrow=False,
        font=dict(size=10),
        align="left",
        bgcolor="white",
        opacity=0.8,
        bordercolor="gray",
        borderwidth=1
    )
    
    # Show additional information for analysis
    print("\n--- Analysis by quadrants ---")
    quadrants = {
        "Earlier and shorter": filtered_df[(filtered_df['Variance - BL Project Finish Date'] < 0) & 
                                         (filtered_df['Variance - BL Project Duration'] < 0)],
        "Earlier but longer": filtered_df[(filtered_df['Variance - BL Project Finish Date'] < 0) & 
                                        (filtered_df['Variance - BL Project Duration'] > 0)],
        "Later but shorter": filtered_df[(filtered_df['Variance - BL Project Finish Date'] > 0) & 
                                       (filtered_df['Variance - BL Project Duration'] < 0)],
        "Later and longer": filtered_df[(filtered_df['Variance - BL Project Finish Date'] > 0) & 
                                      (filtered_df['Variance - BL Project Duration'] > 0)]
    }
    
    for name, data in quadrants.items():
        print(f"{name}: {len(data)} activities")
        if len(data) > 0:
            print(f"  Mean Total Float: {data['Total Float'].mean():.2f}")
    
    # Generate the Excel file if requested
    excel_generated = None
    if generate_excel:
        excel_generated = generate_excel_by_categories(filtered_df, excel_name)
        print(f"\nAn Excel file has been generated with activities by category: {excel_generated}")
        
        # Show information about the content of the generated Excel
        print("\nSummary of the generated Excel content:")
        print(f"Total Float ≤ 0: {len(filtered_df[filtered_df['Total Float'] <= 0])} activities")
        print(f"0 < Total Float ≤ 10: {len(filtered_df[(filtered_df['Total Float'] > 0) & (filtered_df['Total Float'] <= 10)])} activities")
        print(f"10 < Total Float ≤ 20: {len(filtered_df[(filtered_df['Total Float'] > 10) & (filtered_df['Total Float'] <= 20)])} activities")
        print(f"Total Float > 20: {len(filtered_df[filtered_df['Total Float'] > 20])} activities")
    
    # Return the figure and statistics for additional analysis
    return fig, {
        'total_activities': len(filtered_df),
        'min_x': filtered_df['Variance - BL Project Finish Date'].min(),
        'max_x': filtered_df['Variance - BL Project Finish Date'].max(),
        'min_y': filtered_df['Variance - BL Project Duration'].min(),
        'max_y': filtered_df['Variance - BL Project Duration'].max(),
        'category_distribution': category_counts.to_dict(),
        'excel_generated': excel_generated  # Add the path to the generated Excel
    }