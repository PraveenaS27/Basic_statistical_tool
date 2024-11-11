from flask import Flask, request, render_template, redirect, url_for
import pandas as pd

app = Flask(__name__)

# Global variable to hold the DataFrame
df = None

# Route for uploading the Excel file
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    global df
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return 'No file uploaded', 400
        
        file = request.files['file']
        
        try:
            # Read the uploaded file into a DataFrame
            sheet_name = request.form.get('sheet_name')
            df = pd.read_excel(file, sheet_name=sheet_name)
            return redirect(url_for('analyze_columns'))
        except Exception as e:
            return f"Error loading file: {str(e)}", 400
    
    return render_template('upload.html')

# Route for analyzing columns
@app.route('/analyze', methods=['GET', 'POST'])
def analyze_columns():
    global df
    if df is None:
        return redirect(url_for('upload_file'))
    
    if request.method == 'POST':
        try:
            # Get column numbers to analyze from the form
            column_input = request.form.get('columns')
            columns_to_analyze = [int(x.strip()) for x in column_input.split(",")]

            # Validate column numbers
            if any(col < 0 or col >= len(df.columns) for col in columns_to_analyze):
                return 'Invalid column number', 400

            # Perform analysis on selected columns
            analysis_results = perform_analysis(columns_to_analyze)
            return render_template('result.html', results=analysis_results)
        except ValueError:
            return 'Please enter valid column numbers', 400

    return render_template('analyze.html', columns=df.columns)

# Function to perform analysis on selected columns
def perform_analysis(columns):
    global df
    results = []
    
    for col in columns:
        col_name = df.columns[col]
        if pd.api.types.is_numeric_dtype(df[col_name]):
            col_data = df[col_name].dropna()

            # Calculate statistics
            max_value = col_data.max()
            min_value = col_data.min()
            mean_value = col_data.mean()
            median_value = col_data.median()
            std_dev = col_data.std()
            variance = col_data.var()
            count = col_data.count()

            results.append({
                'column': col_name,
                'count': count,
                'mean': mean_value,
                'median': median_value,
                'std_dev': std_dev,
                'variance': variance,
                'max_value': max_value,
                'min_value': min_value
            })
        else:
            results.append({
                'column': col_name,
                'error': 'Non-numeric data'
            })

    return results

if __name__ == '__main__':
    app.run(debug=True)
