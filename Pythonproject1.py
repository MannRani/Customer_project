#import all libraries
import os # for having access to the functionality of the operating system
import numpy as np # for numerical operations
import pandas as pd # for data manipulation and analysis
from datetime import datetime # for data manipulation and times
import matplotlib.pyplot as plt # for visualisation



# Excel Insights  This class provides methods to analyze and visualize Excel data, generate insights, and perform various analyses on the loaded data.
class ExcelInsights:
    """
    A class to analyze and visualize Excel data and generate insights.
    """

    def __init__(self, file_path): # Initialize with the path to the Excel file
        self.file_path = file_path # Path to the Excel file
        self.data = None # Default data to be loaded
        self.sheets = {} # Dictionary to store all sheets as DataFrames
        self.insights = {} # Dictionary to  store insights generated from the data

    def load_excel(self): 
        """
        Load the Excel file and store all sheets as DataFrames.

        """
        if not self.file_path: # Check if the file path is provided
            print("Error: No file path provided.")
            return False
        #try
        try:
            excel_file = pd.ExcelFile(self.file_path) # Load the excel file
            sheets = excel_file.sheet_names # Get all sheet names
            print(f"Sheets found: {sheets}") # Print the names of the sheets found in the Excel file

            # Store each sheet as a DataFrame in the sheets dictionary.
            for sheet in sheets:
                self.sheets[sheet] = pd.read_excel(self.file_path, sheet_name=sheet) # Read each sheet into a DataFrame and store it in the sheets dictionary.
                print(f"Loaded sheet: {sheet} with {len(self.sheets[sheet])} rows and {len(self.sheets[sheet].columns)} columns.") # Print the number of rows and columns in each sheet loaded.

            # Set the defaults data to the first sheet loaded.    
            if sheets:  # If there are any sheets, set the first one as the default data.
                self.data = self.sheets[sheets[0]] # Set the default data to the first sheet loaded.
            print(f"Successfully loaded {len(self.sheets)} sheets from {os.path.basename(self.file_path)}.") # Print the number of sheets loaded from the Excel file.
            return True
        
        #except Exception as e:
        # print(f"Error loading Excel file: {e}")
        # return False
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False

    def get_basic_info(self):
        """
        Get basic information about the loaded data.

        """
        if self.data is None:
            print("No data loaded. Please load an Excel file first.")
            return {}

        info = {
            'rows': len(self.data), #Get the length of rows.
            'columns': len(self.data.columns), # Get the length of columns.
            'column_names': list(self.data.columns), # Get the column_names.
            'data_types': {col: str(dtype) for col, dtype in self.data.dtypes.items()}, # Get the data types of the DataFrame.
            'missing_values': self.data.isnull().sum().to_dict(), # To find missing values.
            'shape': self.data.shape, # Get the shape of the DataFrame (rows, columns).
            'columns': self.data.columns.tolist(), # Get the list of column names.
            'dtypes': self.data.dtypes.to_dict(), # Get the data types of each column.
            'head': self.data.head().to_dict(orient = 'records') #Get the first few rows as a list of dict.
        }

        self.insights["basic_info"] = info # Store the basic information in the insights dictionary.
        return info

    def generate_summary_statistics(self, sheet_name=None): 
        """
        Generate summary statistics for numerical columns.

        Args:
            sheet_name (str, optional): Name of the sheet to analyze. If None, uses the default data.

        Returns:
            dict: Dictionary containing summary statistics for numerical columns.

        """
        if sheet_name and sheet_name in self.sheets: # Check if the specified sheet exists in the loaded sheets.
            data = self.sheets[sheet_name] # If the sheet exists, use it for analysis.
        elif self.data is not None: # If no sheet is specified, use the default data.
            data = self.data # Use the default data.
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}

        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist() # Get numerical columns from the DataFrame.

        if not numerical_cols: # Check if there are any numerical columns.
            # If no numerical columns are found, print a message and return an empty dictionary.
            print("No numerical columns found for summary statistics.")
            return {}
        # Get summary statistics for numerical columns.
        summary = { 
            "numerical_columns": numerical_cols, # List of numerical columns.
            "statistics": data[numerical_cols].describe().to_dict() # Get descriptive statistics for numerical columns and convert to dictionary.
        }

        if "summary_statistics" not in self.insights: # Check if summary statistics exist in insights.
            self.insights["summary_statistics"] = {} # Initialize summary statistics in insights if not already present.

        if sheet_name:
            self.insights["summary_statistics"][sheet_name] = summary # Store summary statistics for the specified sheet.
        # If no sheet name is provided, store summary statistics under the default key.
        else:
            self.insights["summary_statistics"]["default"] = summary 

        return summary

    # Find correlations between numerical columns above a given threshold.

    def find_correlations(self, sheet_name=None, threshold=0.5):
        """
        Find correlations between numerical columns above a given threshold.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}

        # Get numerical columns for summary statistics.
        numerical_cols = data.select_dtypes(include = [np.number]).columns.tolist()

        if len(numerical_cols) < 2: # Check if there are at least two numerical columns for correlation analysis.
            # If not enough numerical columns are found, print a message and return an empty dictionary.
            print("Not enough numerical columns for correlation analysis.")
            return {}

        corr_matrix = data[numerical_cols].corr().abs() # Calculate the absolute correlation matrix for numerical columns.

        high_corr = {} # Dictionary to store pairs of columns with high correlations.
        for i in range(len(numerical_cols)): # Iterate through each numerical column.
            for j in range(i + 1, len(numerical_cols)): 
                col1 = numerical_cols[i]
                col2 = numerical_cols[j]
                corr_value = corr_matrix.loc[col1, col2] # Get the correlation value between the two columns.
                if corr_value >= threshold: # Check if the correlation value is above the specified threshold.
                    high_corr[f"{col1} - {col2}"] = corr_value # Store the pair of columns and their correlation value in the high_corr dictionary.

        if "correlations" not in self.insights: # Check if correlations exist in insights.
            self.insights["correlations"] = {} # Initialize correlations in insights if not already present.

        if sheet_name: # If a sheet name is provided, store the high correlations for that sheet.
            self.insights["correlations"][sheet_name] =  high_corr # Store high correlations for the specified sheet.
        else:
            self.insights["correlations"]["default"] = high_corr # If no sheet name is provided, store high correlations under the default key.

        return high_corr

    def identify_outliers(self, sheet_name=None, method="iqr", threshold=0.5): # Identify outliers in numerical columns.
        """
        Identify outliers in numerical columns.
        
        Args:
           Sheet_name(str, optional): Name of the sheet to analyse. If None, users the default data.
           method (str, optional): Method to use for ioutlier detection ('iqr' or 'zscore').
           threshold (float, optional): Threshold for outlier detection.

        Returns:
            dict: Dictionary containing outliers for each numerical column.

        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        #Get numerical columns
        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()

        if not numerical_cols: # Check if there are any numerical columns.
            # If no numerical columns are found, print a message and return an empty dictionary.
            print("No numerical columns found in the data.")
            return {}

        outliers = {} # Dictionary to store outliers for each numerical column.
        for col in numerical_cols: # Iterate through each numerical column.
            col_data = data[col].dropna() # Drop NaN values from the column data.

            if method == "iqr":
                # IQR method
                Q1 = col_data.quantile(0.25) # Calculate the first quartile (25th percentile)
                Q3 = col_data.quantile(0.75) # Calculate the third quartile (75th percentile)
                IQR = Q3 - Q1 # Calculate the interquartile range (IQR)

                lower_bound = Q1 - threshold * IQR # Calculate the lower bound for outliers
                upper_bound = Q3 + threshold * IQR # Calculate the upper bound for outliers

                col_outliers = col_data[(col_data < lower_bound ) | (col_data > upper_bound)] # Identify outliers based on the lower and upper bounds.

            elif method == "zscore": 
                # Z-score method
                mean = col_data.mean()
                std = col_data.std()

                if std == 0:
                    continue

                z_scores = abs((col_data - mean) / std) # Calculate the z-scores for the column data.
                col_outliers = col_data[z_scores >threshold] # Identify outliers based on the z-scores exceeding the threshold.

            if not col_outliers.empty: # Check if there are any outliers in the column.
                outliers[col] = {
                    "count": len(col_outliers), # Count of outliers in the column.
                    "percentage": (len(col_outliers) / len(col_data)) * 100, # Percentage of outliers in the column.
                    "values": col_outliers.tolist() if len(col_outliers) <= 10 else col_outliers.tolist()[:10] # limit to 10 values

                }   

        # Add to insights
        if "outliers" not in self.insights: # If no outliers key exists, create one.
            self.insights["outliers"] = {} # Initialize the outliers dictionary in insights.

        if sheet_name: # If a sheet name is provided, store the outliers for that sheet.
            self.insights["outliers"][sheet_name] = outliers # Store outliers for the specified sheet.
        # If no sheet name is provided, store outliers under the default key.
        else:
            self.insights["outliers"]["default"] = outliers

        return outliers
    
    def analyze_categorical_data(self, sheet_name = None, top_n = 5):
        """Analyse categorical columns in the data.
        
        Args:
           sheet_name (str, optional): None of the sheet to analyse. If none, use the default data.
           top_n(int,optional): Number of top categories to include in the analysis.
           
        Return: 
            dict: Dictionary containing analysis of categorical columns.
            
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]

        elif self.data is not None:
            data = self.data

        else:
            print("No data loaded. plese load an Excel file first.") 
            return {} 
        
        #Get categorical columns (object, string, or category dtype)
        categorical_cols = data.select_dtypes(include = ["object", "string","category"]).columns.tolist()

        if not categorical_cols:
            print("No categorical columns found in the data.")
            return{}
        
        categorical_analysis = {}

        for col in categorical_cols:
            # Count value frequencies
            value_counts = data[col].value_counts()

            #Get top N categories
            top_categories = value_counts.head(top_n).to_dict()

            #Calculate percentagge of total for each category
            total_count = len(data[col].dropna())
            top_categories_pct = {k: (v /total_count) * 100 for k , v in top_categories.items()}

            #Count unique values and missing values
            unique_count = data[col].nunique()
            missing_count = data[col].isnull().sum()

            categorical_analysis[col] = {
                "unique_values": unique_count, # Count unique values in the column
                "missing_values": missing_count, # Count missing values in the column
                "missing_percentage" : (missing_count / len(data)) *100, # Calculate percentage of missing values
                "top_categories": top_categories, # Get top N categories and their counts
                "top_categories_pct": top_categories_pct # Get top N categories and their percentages
            }

        # Add in insights
        if "categorical_analysis" not in self.insights: # If no categorical_analysis key exists, create one.
            self.insights["categorical_analysis"] = {} # Initialize the categorical_analysis dictionary in insights.

        if sheet_name:
            self.insights["categorical_analysis"][sheet_name] = categorical_analysis # Store categorical analysis for the specified sheet.

        # If no sheet name is provided, store categorical analysis under the default key.        
        else:
            self.insights["categorical_analysis"]["default"] =  categorical_analysis

        return categorical_analysis
    
    def analyze_date_columns (self, sheet_name = None):
        """Analyse data columns in the data.
        
        Args:
           sheet_name(str , optional): Name of the sheet to analyse. If None, uses the data.
                        
        Returns:
            dict: Dictionary containing analysis of date columns.
        """
        
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        date_analysis = {}

        #Try to identify date columns
        for col in data.columns:
            date_col = pd.to_datetime(data[col], errors='coerce')  # Convert to datetime, coerce errors to Nan
            # Check if the column is already in datetime type
            if pd.api.types.is_datetime64_any_dtype(data[col]):
                date_cols = data[col]
            else:
                # Try to convert the column to datetime
                try:
                    
                    if date_col.notnull().sum() / len(date_col) < 0.7:
                        continue # If more than 70% values could be converted to dates , comsider it a date column                        
                except:
                    continue

            #Analyze the date column
            date_analysis[col] = {
                "min_date" : date_col.min().strftime('%Y-%m-%d') if not pd.isna(date_col.min()) else None , # Get the minimum date
                "max_date": date_col.max().strftime('%Y-%m-%d') if not pd.isna(date_col.max()) else None, # Get the maximum date
                "date_range": str(date_col.max() - date_col.min()) if not pd.isna(date_col.min()) and not pd.isna(date_col.max()) else None, # Get the date range
                "missing_values": date_col.isnull().sum(), # Count missing values
                "missing_percentage": (date_col.isnull().sum() / len(date_col)) * 100, # Calculate missing percentage
            }

            #Add day of week distribution if there are enough dates
            if date_col.notnull().sum() > 10:
                day_of_week = date_col.dt.day_name().value_counts().to_dict()
                date_analysis[col]["day_of_week_distribution"] = day_of_week

                #Add month distribution
                month_dist = date_col.dt.month_name().value_counts().to_dict()
                date_analysis[col]["month_distribution"] = month_dist

                # Add year distibution
                year_dist = date_col.dt.year.value_counts().to_dict()
                date_analysis[col] ["year_distribution"] = year_dist

            
          
        # Add to insights
        if "date_analysis" not in self.insights:
            self.insights["date_analysis"] = {}

        if sheet_name:
            self.insights["date_analysis"][sheet_name] = date_analysis
        else:
            self.insights["date_analysis"]["default"] = date_analysis

        return date_analysis

    def generate_insights_report(self, output_file= None):
        """Generate a comprehensive insights report.
        
        Args:
           output_file (str,optional): Path to save the report. If None, returns the report as a string.
           
           Returns:
            str: Reprot as a string if output_file is NOne, otherwise None.
            
        """ 
        if not self.insights: # Check if insights have been generated.
            # If no insights have been generated, print a message and return None.
            print("No insights generated. Please analyze the data forst.")
            return None
        
        report = [] # Initialize an empty list to store the report content.
        report.append("Excel Insights Report") # Add the title of the report.
        report.append(f"Generated on: {datetime.now().strftime('%y-%m-%d %H:%M:%S')}\n") # Add the current date and time of report generation.

        # Add file information
        if self.file_path: # Check if the file path is provided.
            # If the file path is provided, add file information to the report.
            report.append(f"## File Information") # Add file information section
            report.append(f"- File: {os.path.basename(self.file_path)}") # Get the base name of the file.
            report.append(f"- Path: {self.file_path}") # Add the full path of the file.
            if "basic_info" in self.insights: # Check if basic information is available in insights.
                info = self.insights["basic_info"] # Get the basic information from insights.
                report.append(f"- Sheets : {',' .join(info['column_names'])}") # Add the names of the sheets in the file.
                report.append(f"- Rows: {info['rows']}") # Add the number of rows in the file.
                report.append(f"- Columns: {info['columns']}") # Add the numbers of columns in the file.
                report.append("") # Add an empty line for better readability

        # Add summary statistics
        if "summary_statistics" in self.insights: # Check if summary statistics are available in insights.
            report.append("## Summary statistics") # Add summary statistics section
            for sheet, stats in self.insights["summary_statistics"].items(): # Check each sheet in the summary statistics.
                if sheet != "default" : # If the sheet is not the default sheet, add a section for it.
                    report.append(f"\n ### sheet: {sheet}") # Add a subheading for the sheet.

                if not stats.get("numerical_columns"): # Check if there are any numerical columns in the statistics.
                    report.append("No numerical columns found for analysis.") # If no numerical columns are found, print a message and continue to the next sheet.
                    continue

                report.append("The following numerical columns were analyzed") # Add a message indicating which columns were analyzed.
                report.append(f"- {' . '.join(stats['numerical_columns'])} \n") # List the numerical columns that were analyzed.

                for col, col_stats in stats["statistics"].items(): # Check each column in the statistics.
                    report.append(f"## {col}") # Add a subheading for the column.
                    report.append(f" - Count: {col_stats.get('count','N/A')}") # Add the count of non-null values in the column.
                    report.append(f" - Mean: {col_stats.get('mean', 'N/A'):.2f}") # Add the mean of the column, formatted to two decimal places.
                    report.append(f" - Std Dev: {col_stats.get('std', 'N/A'):.2f}") # Add the standard deviation of the column, formatted to two decimal places.
                    report.append(f" - Min: {col_stats.get('min', 'N/A'):.2f}") # Add the minimum value of the column, formatted to two decimal places.
                    report.append(f" - 25%: {col_stats.get('25%', 'N/A'):.2f}") # Add the 25th percentile of the column, formatted to two decimal places.
                    report.append(f" - Median: {col_stats.get('median')}") # Add the median of the column.
                    report.append(f" - 75%: {col_stats.get('75%', 'N/A'):.2f}") # Add the 75th percentile of the column, formatted to two decimal places.
                    report.append(f" - Max: {col_stats.get('max', 'N/A'):.2f}") # Add the maximum value of the column, formatted to two decimal places.
                    report.append("") # Add an empty line for better readability

        # Add correlations
        if "correlations" in self.insights: # Check if correlations are available in insights.
            report.append("## Correlations") # Add correlations section
            for sheet, corrs in self.insights["correlations"].items():
                if sheet != "default": # If the sheet is not the default sheet, add a section for it.
                    report.append("No significant correlations found.") # If no correlations are found, print a message and continue to the next sheet.

                if not corrs: # If no correlations are found, print a message and continue to the next sheet.
                    report.append("No significant correlations found.") 
                    continue
                report.append("The following pairs of columns show significant correlations:")
                for pair, corr in corrs.items(): # Check each pair of columns in the correlations.
                    report.append(f" - {pair}: {corr:.2f}") # Add the pair of columns and their correlation value, formatted to two decimal places.
                report.append("") # Add an empty line for better readability

        

        


        # Add outliers
        if "outliers" in self.insights: # Check if outliers are available in insights.
            report.append("## Outliers") # Add outliers section
            for sheet , outs in self.insights ["outliers"].items(): # Check each sheet in the outliers.
                if sheet != "default": # If the sheet is not the default sheet, add a section for it.
                    report.append(f"\n ## Sheet : {sheet}") # Add a subheading for the sheet.

                if not outs: # If no outliers are found, print a message and continue to the next sheet.
                    report.append("No outliers detected.")
                    continue

                for col, out_info in outs.items(): # Check each column in the outliers.
                    report.append(f" ## {col}") # Add a subheading for the column.
                    report.append(f" - Count of outliers: {out_info['count']}") # Add the count of outliers in the column.
                    report.append(f" - Percentage of data: {out_info['percentage']:.2f}%") # Add the percentage of outliers in the column, formatted to two decimal places.
                    if out_info['values']: # If there are outlier values, add them to the report.
                        report.append(f"- Shape outliers: {out_info['values']}") # Add the outlier values in the column.
                    report.append("")

        # Add categorical analysis
        if "categorical_analysis" in self.insights: # Check if categorical analysis is available in insights.
            report.append("## Categorical Analysis") # Add categorical analysis section
            for sheet, cat_analysis in self.insights["categorical_analysis"].items(): # Check each sheet in the categorical analysis.
                if sheet != "default": # If the sheet is not the default sheet, add a section for it.
                    report.append(f"\n### Sheet: {sheet}") # Add a subheading for the sheet.

                if not cat_analysis: # If no categorical columns are found, print a message and continue to the next sheet.
                    report.append("No categorical columns found for analysis.")
                    continue

                for col, col_analysis in cat_analysis.items(): # Check each column in the categorical analysis.
                    report.append(f" ## {col}\n") # Add a subheading for the column.
                    report.append(f" - Unique values: {col_analysis['unique_values']}\n") # Add the count of unique values in the column.
                    report.append(f" - Missing values: {col_analysis['missing_values']} ({col_analysis['missing_percentage']:.2f}%)\n") # Add the count and percentage of missing values in the column, formatted to two decimal places.
                    report.append("\nTop categories:") # Add a subheading for the top categories in the column.

                    for category, count in col_analysis['top_categories'].items(): # Check each top category in the column analysis.
                        category_str = str(category) if category is not None else "Null" # Convert the category to a string, handling None values.
                        pct = col_analysis['top_categories_pct'].get(category, 0) # Get the percentage of the category in the column.
                        report.append(f" - {category_str}: {count} ({pct:.2f}%)") # Add the category, its count, and percentage, formatted to two decimal places.
                    report.append("")

        # Add date analysis
        if "date_analysis" in self.insights: # Check if date analysis is available in insights.
            report.append("## Date Analysis") # Add date analysis section
            for sheet, date_analysis in self.insights["date_analysis"].items(): # Check each sheet in the date analysis.
                if sheet != "default": # If the sheet is not the default sheet, add a section for it.
                    report.append(f"\n### Sheet: {sheet}") # Add a subheading for the sheet.

                if not date_analysis: # If no date columns are found, print a message and continue to the next sheet.
                    report.append("No date columns found for analysis.")
                    continue

                for col, analysis in date_analysis.items(): # Check each date column in the date analysis.
                    report.append(f"## {col}") # Add a subheading for the date column.
                    report.append(f"- Date range: {analysis['date_range']}\n") # Add the date range of the column.
                    report.append(f"- Missing values: {analysis['missing_values']} ({analysis['missing_percentage']:.2f}%)\n") # Add the count and percentage of missing values in the column, formatted to two decimal places.

                    if 'day_of_week_distribution' in analysis: # Check if day of week distribution is available in the analysis.
                        report.append("\nDay of week distribution:") # Add day of week distribution section
                        for day, count in analysis['day_of_week_distribution'].items(): # Check each day in the day of week distribution.
                            report.append(f" - {day}: {count}\n") # Add the count for each day.

                    if 'month_distribution' in analysis: # Check if month distribution is available in the analysis.
                        report.append("\nMonth distribution:") # Add month distribution section
                        for month, count in analysis['month_distribution'].items(): # Check each month in the month distribution.
                            report.append(f" - {month}: {count}\n") # Add the count for each month.

                    if 'year_distribution' in analysis: # Check if year distribution is available in the analysis.
                        report.append("\nYear distribution:") # Add year distribution section
                        for year, count in analysis['year_distribution'].items(): # Check each year in the year distribution.
                            report.append(f" - {year}: {count}\n") # Add the count for each year.
        
        # compile the report
        report_text = "\n".join(report)

        # Save the file if specified
        if output_file: # Check if output file is specified.
            try: # Try to save the report to the specified file.
                with open(output_file, 'w') as f: # Open the output file in write mode.
                    f.write(report_text) # Write the report text to the file.
                print(f"Report saved to {output_file}") # Print a message indicating that the report has been saved.
            except Exception as e: # If an error occurs while saving the report, print an error message.
                print(f"Error saving report: {e}") # Print the error message.
                return None
        else:
            return report_text    

    # visualization of data
    def visualize_data(self, output_dir = None , sheet_name=None):   
        """
        Generate visualizations for the data.
        
        Args:
            output_dir (str, optional): Directory to save the visualizations. If None, displays them inline.
            sheet_name (str, optional): Name of the sheet to visualize. If None, uses the default data.

        Returns:
            list: List of file paths for saved visualizations if output_dir is specified, otherwise None.
            
        """ 
        if sheet_name and sheet_name in self.sheets: # Check if the specified sheet exists in the loaded sheets.
            data = self.sheets[sheet_name] # If the sheet exists, use it for visualization.
        elif self.data is not None: # If no sheet is specified, use the default data.
            data = self.data # Use the default data.
        else: # If no data is available, print an error message.
            print("No data loaded. Please load an Excel file first.")
            return None
        
        if output_dir and not os.path.exists(output_dir): # Check if the output directory exists.
            os.makedirs(output_dir) # If the output directory does not exist, create it.

        saved_files = [] # List to store file paths of saved visualizations.

        # Get numerical and categorical columns
        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist() 
        categorical_cols = data.select_dtypes(include=["object", "string", "category"]).columns.tolist()

        # 1. Histogram for numerical columns
        for col in numerical_cols[:5]:  # Limit to first 5 numerical columns
            plt.figure(figsize=(10, 6)) # Create a new figure for the histogram
            plt.hist(data[col].dropna(), bins=30, color='blue', alpha=0.7 , edgecolor='black') # Plot the histogram for the numerical column
            plt.title(f" Distribution of {col}") # Add title to the histogram
            plt.xlabel(col) # Add x-axis label
            plt.ylabel("Frequency") # Add y-axis label
            plt.grid(True , alpha=0.3) # Add grid lines to the histogram

            if output_dir: # If output directory is specified, save the histogram as a PNG file.
                file_path = os.path.join(output_dir, f"{col}_histogram.png") # Create the file path for the histogram
                plt.savefig(file_path) # Save the histogram to the specified file path
                saved_files.append(file_path) # Append the file path to the list of saved files
                plt.close() # Close the figure to free up memory

            else:
                plt.show()

        # 2. Bar chart for categorical columns
        for col in categorical_cols[:5]: # Limit to first 5 categorical columns
            # Get top 10 categories
            value_counts = data[col].value_counts().head(10)

            plt.figure(figsize=(10, 6)) # Create a new figure for the bar chart
            bars = plt.bar(value_counts.index.astype(str), value_counts.values, color='orange', alpha=0.7, edgecolor='black') # Plot the bar chart for the categorical column
            plt.title(f"Top 10 Categories in {col}") # Add title to the bar chart
            plt.xlabel(col) # Add x-axis label
            plt.ylabel("Count") # Add y-axis label
            plt.xticks(rotation=45, ha='right') # Rotate x-axis labels for better readability
            plt.grid(True,axis='y', alpha=0.3  ) # Add grid lines to the y-axis of the bar chart    

            # Add count labels on top of the bars
            for bar in bars: # Iterate through each bar in the bar chart
                height = bar.get_height() # Get the height of the bar
                plt.text(bar.get_x() + bar.get_width() / 2, height + 0.1, # Add the count label on top of the bar
                         f"{height}", ha='center', va='bottom') # Add the count label on top of the bar

            plt.tight_layout() # Adjust layout to prevent overlap of labels

            if output_dir: # If output directory is specified, save the bar chart as a PNG file.
                file_path = os.path.join(output_dir, f"{col}_bar_chart.png") # Create the file path for the bar chart
                plt.savefig(file_path) # Save the bar chart to the specified file path
                saved_files.append(file_path) # Append the file path to the list of saved files
                plt.close() # Close the figure to free up memory

            else: # If no output directory is specified, display the bar chart inline.
                plt.show()

        # 3. Correlation heatmap for numerical columns
        if len(numerical_cols) > 1: # Check if there are at least two numerical columns for correlation analysis.
            plt.figure(figsize=(12, 8)) # Create a new figure for the correlation heatmap
            corr_matrix = data[numerical_cols].corr() # Calculate the correlation matrix for numerical columns
            plt.imshow(corr_matrix, cmap='coolwarm', interpolation='none' , aspect='auto') # Plot the correlation heatmap
            plt.colorbar(label='Correlation Coefficient') # Add color bar to the heatmap
            plt.title("Correlation Heatmap") # Add title to the heatmap

            # Add correlation values
            for i in range(len(numerical_cols)): # Iterate through each numerical column
                for j in range(len(numerical_cols)): # Iterate through each numerical column again
                    plt.text(j, i, f"{corr_matrix.iloc[i, j]:.2f}", ha='center', va='center', color='white' if abs(corr_matrix.iloc[i, j]) > 0.5 else 'black') # Add correlation values to the heatmap

            plt.xticks(range(len(numerical_cols)), numerical_cols, rotation=45, ha='right') # Set x-axis ticks to numerical column names
            plt.yticks(range(len(numerical_cols)), numerical_cols) # Set y-axis ticks to numerical column names
            plt.tight_layout() # Adjust layout to prevent overlap of labels

            if output_dir: # If output directory is specified, save the correlation heatmap as a PNG file.
                file_path = os.path.join(output_dir, "correlation_heatmap.png") # Create the file path for the correlation heatmap
                plt.savefig(file_path) # Save the correlation heatmap to the specified file path
                saved_files.append(file_path) # Append the file path to the list of saved files
                plt.close() # Close the figure to free up memory

            else: # If no output directory is specified, display the correlation heatmap inline.
                plt.show() 

        # 4. Box plots for numerical columns to visulalize outliers
        if numerical_cols: # Check if there are any numerical columns for box plots.
            plt.figure(figsize=(12, 6)) # Create a new figure for the box plots
            data[numerical_cols[:10]].boxplot() # Plot box plots for the first 10 numerical columns
            plt.xlabel('Numerical Columns') # Add x-axis label
            plt.ylabel('Values') # Add y-axis label 
            plt.title('Box Plots for numrical columns') # Add title to the box plots
            plt.xticks(rotation = 45 , ha = 'right') # Rotate x-axis labels for better readability
            plt.grid(True, alpha =0.3) # Add grid lines to the box plots
            plt.tight_layout() # Adjust layout to prevent overlap of labels

            if output_dir: # If output directory is specified, save the box plots as a PNG file.
                file_path = os.path.join(output_dir , "boxplots.png") # Create the file path for the box plots
                plt.savefig(file_path) # Save the box plots to the specified file path
                saved_files.append(file_path) # Append the file path to the list of saved files
                plt.close() # Close the figure to free up memory

            else: # If no output directory is specified, display the box plots inline.
                plt.show()

        # 5. Pie charts for categorical columns with few unique values
        for col in categorical_cols[:3]: # Limit to first 3 categorical columns
            value_counts = data[col].value_counts()
            # Only create pie chart if there are enough unique values
            if len(value_counts) <= 10: # Check if the number of unique values is less than or equal to 10
                plt.figure(figsize= (8, 8)) 
                plt.pie(value_counts.values, labels=value_counts.index.astype(str), autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors , shadow=True) # Plot the pie chart for the categorical column
                plt.title(f"Pie Chart of {col}") # Add title to the pie chart
                plt.axis('equal') # Equal aspect ratio ensures that pie chart is a circle.

            if output_dir: # If output directory is specified, save the pie chart as a PNG file.
                file_path = os.path.join(output_dir, f"{col}_pie_chart.png") # Create the file path for the pie chart
                plt.savefig(file_path) # Save the pie chart to the specified file path
                saved_files.append(file_path) # Append the file path to the list of saved files
                plt.close()  # Close the figure to free up memory

            else:
                plt.show()

        return saved_files


# Create an instance of the class ExcelInsights
file_path = r"C:\Users\imann\OneDrive\Desktop\customers\sample_data.xlsx"  # Update this path
excel_insights = ExcelInsights(file_path)
# load_excel() method to load the Excel file
if excel_insights.load_excel():
    print("\n\nBasic Information:")
    # get_basic_info() method to get basic information about the loaded data
    basic_info = excel_insights.get_basic_info()
    for key, value in basic_info.items():
        print(f"{key}: {value}")
    # Call generate_summary_statistics() methods to get summary statistics    
    print("\n\nSummary Statistics:")
    summary_statistics = excel_insights.generate_summary_statistics()
    for key, value in summary_statistics.items():
        print(f"{key}: {value}")
    #Print the correlations
    print("\n\nCorrelations:")
    correlations = excel_insights.find_correlations()
    for key, value in correlations.items():
        print(f"{key}: {value}")


    print("\n\noutliers:")  
    outliers = excel_insights.identify_outliers()

    for key, value in outliers.items():
        print(f"{key}:  {value}")    
 
    print("\n\nCategorical Analysis:")
    #Call analyse_categorical_data() method to analyse categorical data
    categorical_analysis = excel_insights.analyze_categorical_data()
     #Print the categorical analysis
    for key, value in categorical_analysis.items():
        print(f"{key}: {value}") 

    print("\n\ndate_Analysis:")
    #Call analyse_date_columns() method to analyse date columns
    date_analysis = excel_insights.analyze_date_columns()
    #Print the date analysis    
    for key, value in date_analysis.items():
        print (f" {key}: {value}")

    # Generate insights report
    print("\n\nGenerating Insights Report:")
    # Call generate_insights_report() method to generate a comprehensive insights report
    insights_report = excel_insights.generate_insights_report(output_file="insights_report.txt")
    # Print the insights report if it was generated as a string
    if insights_report:
        print(insights_report)

    else:
        print("Insights report generated and saved to 'insights_report.txt'.")

    # Visualize data and save visualizations to a directory
    visualizations = excel_insights.visualize_data(output_dir="visualizations") 
    if visualizations: # If visualizations were generated and saved, print the file paths
        print("Visualizations saved to the following files:") 
        for file in visualizations: # Print each file path
            print(file) 
    else:
        print("No visualizations generated.")