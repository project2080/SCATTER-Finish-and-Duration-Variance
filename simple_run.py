"""
Simple script to run the scatter plot visualization with default settings.
"""
from scatter_analysis import create_interactive_scatter_plot

# Run the function with the specific file
excel_file = 'SCATTER Finish and Duration Variance.xlsx'
fig, statistics = create_interactive_scatter_plot(excel_file)

# Display the interactive chart
fig.show()

# Print the returned statistics
print("\nStatistics returned by the function:")
for key, value in statistics.items():
    print(f"{key}: {value}")

# Save the chart as an interactive HTML to share it
fig.write_html('interactive_scatter_plot.html')
