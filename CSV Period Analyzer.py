import pandas as pd
import numpy as np
from scipy.fft import fft, fftfreq
import os
from tkinter import Tk, filedialog, simpledialog
from openpyxl import Workbook
import logging

# Setup logging
logging.basicConfig(filename='period_calculation.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def calculate_period(file_path, x_col_num=3, y_col_num=4, encoding='shift_jis'):
    try:
        # Try reading the CSV file with the specified encoding
        try:
            df = pd.read_csv(file_path, encoding=encoding)
        except UnicodeDecodeError:
            logging.warning(f"Encoding error with {encoding} for {file_path}. Trying utf-8.")
            df = pd.read_csv(file_path, encoding='utf-8')
        except Exception as e:
            logging.error(f"Failed to read {file_path} with {encoding} encoding: {e}")
            return None

        # Replace '#DIV/0!' and NaN with appropriate NaN values and drop irrelevant columns
        df.replace('#DIV/0!', pd.NA, inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        df.columns = [col.strip() for col in df.columns]
        
        # Extract the specified columns by their index
        try:
            x_column = df.columns[x_col_num]
            y_column = df.columns[y_col_num]
        except IndexError as e:
            logging.error(f"Column index error in {file_path}: {e}")
            return None
        
        # Drop rows where the x or y values are NaN
        df = df.dropna(subset=[x_column, y_column])

        # Extract the y-values for FFT analysis
        y_values = df[y_column].values

        # Number of sample points
        N = len(y_values)

        # Sample spacing
        T = df[x_column].diff().dropna().mean()

        if N < 2 or T == 0:
            raise ValueError("Insufficient data points or invalid sample spacing.")

        # Perform the FFT
        yf = fft(y_values)
        xf = fftfreq(N, T)[:N//2]

        # Identify the peak frequency
        peak_freq = xf[np.argmax(np.abs(yf[:N//2]))]

        # Calculate the period
        period = 1 / peak_freq
        
        return period
    except ValueError as e:
        logging.error(f"Value error in {file_path}: {e}")
        return None
    except Exception as e:
        logging.error(f"Unexpected error processing {file_path}: {e}")
        return None

def main():
    # Create a Tkinter root window
    root = Tk()
    root.withdraw()  # Hide the root window

    # Ask the user to select a directory
    directory = filedialog.askdirectory(title="Select Directory with CSV Files")

    if not directory:
        print("No directory selected.")
        return

    # Ask for encoding
    encoding = simpledialog.askstring("Input", "Enter file encoding (default: shift_jis):", initialvalue="shift_jis")

    # Prepare the Excel workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["File Name", "Period"])

    # Process each CSV file in the directory
    for file_name in os.listdir(directory):
        if file_name.endswith('.csv'):
            file_path = os.path.join(directory, file_name)
            period = calculate_period(file_path, encoding=encoding)
            if period is not None:
                sheet.append([file_name, period])
            else:
                sheet.append([file_name, "Error"])
                logging.error(f"Error processing {file_name}")

    # Save the results to an Excel file
    output_file = os.path.join(directory, "periods.xlsx")
    workbook.save(output_file)
    logging.info(f"Results saved to {output_file}")
    print(f"Results saved to {output_file}")

if __name__ == "__main__":
    main()
