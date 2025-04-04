"""
SCATTER Finish and Duration Variance - Visualization Runner

This script executes the visualization functions from scatter_analysis.py
to analyze project activities based on their finish date variance,
duration variance, and total float.

Usage:
    python run_visualization.py [--excel FILENAME] [--output OUTPUT_NAME] [--no-excel]

Options:
    --excel FILENAME    Specify the Excel file to analyze (default: 'SCATTER Finish and Duration Variance.xlsx')
    --output OUTPUT_NAME    Name for output files (default: 'interactive_scatter_plot')
    --no-excel          Skip generating the Excel file with categorized activities
"""

import os
import argparse
from scatter_analysis import create_interactive_scatter_plot

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Run SCATTER Finish and Duration Variance visualization')
    parser.add_argument('--excel', type=str, default='SCATTER Finish and Duration Variance.xlsx',
                        help='Excel file to analyze')
    parser.add_argument('--output', type=str, default='interactive_scatter_plot',
                        help='Base name for output files')
    parser.add_argument('--no-excel', action='store_true',
                        help='Skip generating Excel output file')
    
    return parser.parse_args()

def main():
    """Main function to run the visualization."""
    # Parse command line arguments
    args = parse_arguments()
    
    # Check if input file exists
    if not os.path.exists(args.excel):
        print(f"Error: File '{args.excel}' not found.")
        return
    
    # Run the visualization
    print(f"Analyzing file: {args.excel}")
    excel_name = f"{os.path.splitext(args.output)[0]}_categories.xlsx" if not args.no_excel else None
    
    # Create the interactive scatter plot
    fig, statistics = create_interactive_scatter_plot(
        args.excel,
        generate_excel=not args.no_excel,
        excel_name=excel_name
    )
    
    # Save the output files
    html_output = f"{args.output}.html"
    fig.write_html(html_output)
    print(f"Interactive HTML saved to: {html_output}")
    
    # Display summary statistics
    print("\nSummary Statistics:")
    print(f"Total Activities: {statistics['total_activities']}")
    print(f"Activities by Category:")
    for category, count in statistics['category_distribution'].items():
        print(f"  - {category}: {count}")
    
    # Display the plot
    fig.show()

if __name__ == "__main__":
    main()
