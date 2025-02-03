# Jacques Boucher
# jjrboucher@gmail.com
#
# Written using Google Gemini
#
# Date: 3 February 2025
# Description: Script will prompt you for an Excel sheet via a tkinter menu.
# The script will then have you select the longitude and latitude columns via a pull down in the tkinter menu.
#
# It will then process the Excel file and add a new column to it containing the resolved addresses.
#
# The script will also display a status message indicating which row is being processed.
#

import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import re

def geocode_address(latitude, longitude, retries=3):
    """Geocodes a latitude/longitude pair to an address in English."""
    geolocator = Nominatim(user_agent="geo_app")

    for attempt in range(retries):
        try:
            location = geolocator.reverse((latitude, longitude), timeout=10, language="en")  # Add language="en"
            if location:
                return location.address
            else:
                return "Address not found"

        except (GeocoderTimedOut, GeocoderServiceError) as e:
            print(f"Geocoding failed: {e}. Retrying ({attempt + 1}/{retries})...")
            if attempt < retries - 1:
                import time
                time.sleep(2)
            else:
                print(f"Geocoding failed after {retries} retries.")
                return "Geocoding failed"

        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            return "Error during geocoding"

def convert_coordinates(coord):
    """Converts coordinates from various formats to decimal degrees."""
    if isinstance(coord, str):
        match = re.match(r"(\d+)\s?(°|deg)?\s*(\d+)'?\s*(\d+(?:\.\d+)?)[\"\"]?\s*([NSEW]?)", coord, re.IGNORECASE)
        if match:
            degrees = float(match.group(1))
            minutes = float(match.group(3))
            seconds = float(match.group(4)) if match.group(4) else 0.0
            direction = match.group(5) or ''

            decimal_degrees = degrees + (minutes / 60) + (seconds / 3600)
            if direction.lower() in ('s', 'w'):
                decimal_degrees *= -1
            return decimal_degrees
        else:  # If the regex doesn't match, try decimal
            try:
                return float(coord)
            except ValueError:
                return None
    elif isinstance(coord, (int, float)):
        return float(coord)
    return None


def process_file(file_path, lon_col, lat_col):
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        messagebox.showerror("Error", "File not found.")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        return

    if lon_col not in df.columns or lat_col not in df.columns:
        messagebox.showerror("Error", "Invalid column names.")
        return

    df['Address'] = ""

    total_rows = len(df)

    successful_count = 0
    unsuccessful_count = 0

    for index, row in df.iterrows():

        lat_str = str(row[lat_col])  # Convert to string explicitly
        lon_str = str(row[lon_col])  # Convert to string explicitly


        lat = convert_coordinates(lat_str)  # Convert latitude
        lon = convert_coordinates(lon_str)  # Convert longitude

        if lat is None or lon is None:  # Check for invalid coordinate format
            df.loc[index, 'Address'] = "Invalid Coordinates"
            empty_rows_count += 1
            status_label.config(text=f"Processed row: {index + 1} of {total_rows} (Skipped - Invalid Coordinates)")
            root.update_idletasks()
            continue

        if pd.isna(lat) or pd.isna(lon) or not (lat and lon):
            df.loc[index, 'Address'] = "No Coordinates"
            empty_rows_count += 1
            status_label.config(text=f"Processed row: {index + 1} of {total_rows} (Skipped - No Coordinates)")
            root.update_idletasks()
            continue

        address = geocode_address(lat, lon)  # Assign address here!
        df.loc[index, 'Address'] = address

        if (address == "Address not found" or address == "Geocoding failed" or address == "Error during geocoding" or
                address == "No Coordinates" or address == "Invalid Coordinates"):
            empty_rows_count += 1
            unsuccessful_count += 1
        else:
            empty_rows_count = 0
            successful_count += 1

        update_status_label(successful_count, unsuccessful_count, total_rows, index)  # Call the update function

        root.update_idletasks()

    try:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Addresses added to {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving Excel file: {e}")


def update_status_label(successful, unsuccessful, total, current):
    """Updates the status label with colored counts using ttk."""
    status_text = (
        f"Processed row: {current + 1} of {total} "
        f"(Successful: {successful}, "
        f"Unsuccessful: {unsuccessful})"
    )
    status_label.config(text=status_text)

    root.update_idletasks()

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)
        try:
            df = pd.read_excel(file_path)
            column_names = df.columns.tolist()
            lon_col_dropdown['values'] = column_names
            lat_col_dropdown['values'] = column_names
            lon_col_dropdown.config(state="readonly")
            lat_col_dropdown.config(state="readonly")
        except Exception as e:
            messagebox.showerror("Error", f"Error reading Excel file: {e}")
            file_path_entry.delete(0, tk.END)
            lon_col_dropdown['values'] = []
            lat_col_dropdown['values'] = []


def start_processing():
    file_path = file_path_entry.get()
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return
    lon_col = lon_col_var.get()
    lat_col = lat_col_var.get()
    process_file(file_path, lon_col, lat_col)

root = tk.Tk()
root.title("Geocoding Tool")

file_path_label = tk.Label(root, text="Excel File:")
file_path_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

file_path_entry = tk.Entry(root, width=50)
file_path_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

lon_col_label = tk.Label(root, text="Longitude Column:")
lon_col_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

lon_col_var = tk.StringVar(root)
lon_col_dropdown = ttk.Combobox(root, textvariable=lon_col_var, state="disabled")
lon_col_dropdown.grid(row=1, column=1, padx=5, pady=5)

lat_col_label = tk.Label(root, text="Latitude Column:")
lat_col_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

lat_col_var = tk.StringVar(root)
lat_col_dropdown = ttk.Combobox(root, textvariable=lat_col_var, state="disabled")
lat_col_dropdown.grid(row=2, column=1, padx=5, pady=5)

process_button = tk.Button(root, text="Start Geocoding", command=start_processing)
process_button.grid(row=3, column=0, columnspan=3, pady=10)

status_label = ttk.Label(root, text="Ready")  # ttk label for overall status
status_label.grid(row=4, column=0, columnspan=3, pady=(0, 10))

root.mainloop()
