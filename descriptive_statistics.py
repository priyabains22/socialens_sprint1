import pandas as pd
import openai

def analyze_file(file_path):
    output = {}
    try:
        # Attempt to read file based on its extension
        if file_path.endswith(('.xlsx', '.xls')):
            data = pd.read_excel(file_path, sheet_name=None)  # None reads all sheets
        elif file_path.endswith('.csv'):
            data = {'CSV': pd.read_csv(file_path)}
        else:
            return {'error': "Unsupported file format"}

        # Analyze each sheet
        for sheet_name, df in data.items():
            # Prepare data dictionary for each sheet
            stats = {}
            numeric_columns = df.select_dtypes(include=['float64', 'int64'])
            for column in numeric_columns:
                stats[column] = {
                    'Mean': numeric_columns[column].mean(),
                    'Median': numeric_columns[column].median(),
                    'Mode': numeric_columns[column].mode().values.tolist(),
                    'Standard Deviation': numeric_columns[column].std(),
                    'Variance': numeric_columns[column].var(),
                    'Skewness': numeric_columns[column].skew(),
                    'Kurtosis': numeric_columns[column].kurt(),
                    'Min': numeric_columns[column].min(),
                    'Max': numeric_columns[column].max(),
                    '25th Percentile': numeric_columns[column].quantile(0.25),
                    '50th Percentile': numeric_columns[column].quantile(0.5),
                    '75th Percentile': numeric_columns[column].quantile(0.75),
                    'Count': numeric_columns[column].count(),
                    'Missing Values': numeric_columns[column].isna().sum()
                }
            output[sheet_name] = stats
    except Exception as e:
        return {'error': f"Failed to process file: {str(e)}"}

    return output


def summarize_files(file_path):
    summaries = {}
    try:
        # Attempt to read file based on its extension
        if file_path.endswith(('.xlsx', '.xls')):
            data = pd.read_excel(file_path, sheet_name=None)  # None reads all sheets
        elif file_path.endswith('.csv'):
            data = {'CSV': pd.read_csv(file_path)}
        else:
            return {'error': "Unsupported file format"}

        # Analyze each sheet
        for sheet_name, df in data.items():
            # Prepare data dictionary for each sheet
            summary_lines =[]
            summary_lines.append("I am going to provide a summary of some statistical values from an excel spreadsheet. Please provide in your response a summary and detailed explanation of what these results are and what they mean.")
            stats = {}
            numeric_columns = df.select_dtypes(include=['float64', 'int64'])
            for column in numeric_columns:
                stats[column] = {
                    'Mean': numeric_columns[column].mean(),
                    'Median': numeric_columns[column].median(),
                    'Mode': numeric_columns[column].mode().values.tolist(),
                    'Standard Deviation': numeric_columns[column].std(),
                    'Variance': numeric_columns[column].var(),
                    'Skewness': numeric_columns[column].skew(),
                    'Kurtosis': numeric_columns[column].kurt(),
                    'Min': numeric_columns[column].min(),
                    'Max': numeric_columns[column].max(),
                    '25th Percentile': numeric_columns[column].quantile(0.25),
                    '50th Percentile': numeric_columns[column].quantile(0.5),
                    '75th Percentile': numeric_columns[column].quantile(0.75),
                    'Count': numeric_columns[column].count(),
                    'Missing Values': numeric_columns[column].isna().sum()
                }
                summary_lines.append(
                f"The column '{column}' has a mean of {stats[column]['Mean']}, median of {stats[column]['Median']}, and standard deviation of {stats[column]['Standard Deviation']}. "
                f"It ranges from a minimum of {stats[column]['Min']} to a maximum of {stats[column]['Max']}. "
                f"The 25th percentile is {stats[column]['25th Percentile']}, 50th percentile (median) is {stats[column]['50th Percentile']}, and 75th percentile is {stats[column]['75th Percentile']}. "
                f"It has {stats[column]['Missing Values']} missing values."
                )
            summaries[sheet_name] = " ".join(summary_lines)
        #OpenAI API Key
        openai.api_key = 'sk-proj-l28vi099mhVpldhAdv1WT3BlbkFJ9A23j9wCGtblQ6SORaSR'
        #Sending summaries to ChatGPT 3.5 Turbo for interpretation
        for sheet_name, summary in summaries.items():
            response = openai.chat.completions.create(
                messages=[
                    {
                        "role": "system", "content": "You are a helpful assistant.",
                        "content": summary,
                    }
                ],
                model="gpt-3.5-turbo",
            )
            summaries[sheet_name] = response.choices[0].message.content
    except Exception as e:
        return {'error': f"Failed to process file: {str(e)}"}

    return summaries
