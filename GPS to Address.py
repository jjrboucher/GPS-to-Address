# Enhanced Geocoding Tool
# Originally created with Gemini, updated with Claude.ai
# Jacques Boucher
# jjrboucher@gmail.com
#
# Last updated 16 May 2025
# Enhanced GUI by Claude 3.7 Sonnet

import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import re
import time
import threading
from PIL import Image, ImageTk
import os
from ttkthemes import ThemedStyle
import base64
import io

def geocode_address(latitude, longitude, retries=3):
    """Geocodes a latitude/longitude pair to an address in English."""
    geolocator = Nominatim(user_agent="geo_app")

    for attempt in range(retries):
        try:
            location = geolocator.reverse((latitude, longitude), timeout=10, language="en")
            if location:
                return location.address
            else:
                return "Address not found"

        except (GeocoderTimedOut, GeocoderServiceError) as e:
            print(f"Geocoding failed: {e}. Retrying ({attempt + 1}/{retries})...")
            if attempt < retries - 1:
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
        # Try to directly convert simple decimal format (including negative values)
        try:
            return float(coord)
        except ValueError:
            # If direct conversion fails, try the more complex format
            match = re.match(r"(\-?\d+)\s?(°|deg)?\s*(\d+)'?\s*(\d+(?:\.\d+)?)[\"\"]?\s*([NSEW]?)", coord, re.IGNORECASE)
            if match:
                degrees = float(match.group(1))
                minutes = float(match.group(3))
                seconds = float(match.group(4)) if match.group(4) else 0.0
                direction = match.group(5) or ''

                decimal_degrees = abs(degrees) + (minutes / 60) + (seconds / 3600)
                # Apply negative sign if original degrees were negative or if direction is S/W
                if degrees < 0 or direction.lower() in ('s', 'w'):
                    decimal_degrees *= -1
                return decimal_degrees
            return None
    elif isinstance(coord, (int, float)):
        return float(coord)
    return None

def process_file_threaded():
    """Run the processing in a separate thread to keep the UI responsive"""
    threading.Thread(target=process_file, daemon=True).start()

def process_file():
    file_path = file_path_entry.get()
    lon_col = lon_col_var.get()
    lat_col = lat_col_var.get()
    
    # Disable buttons during processing
    process_button.config(state=tk.DISABLED)
    browse_button.config(state=tk.DISABLED)
    
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        messagebox.showerror("Error", "File not found.")
        reset_ui()
        return
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        reset_ui()
        return

    if lon_col not in df.columns or lat_col not in df.columns:
        messagebox.showerror("Error", "Invalid column names.")
        reset_ui()
        return

    df['Address'] = ""

    total_rows = len(df)
    progress_bar["maximum"] = total_rows
    
    # Update the status indicators
    status_frame.pack(fill=tk.X, padx=20, pady=10)
    progress_bar.pack(fill=tk.X, padx=20, pady=(0, 10))

    successful_count = 0
    unsuccessful_count = 0
    
    # Show processing panel
    processing_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
    root.update_idletasks()

    for index, row in df.iterrows():
        lat_str = str(row[lat_col])
        lon_str = str(row[lon_col])

        # Update coordinates display
        coords_var.set(f"Coordinates: {lat_str}, {lon_str}")
        
        lat = convert_coordinates(lat_str)
        lon = convert_coordinates(lon_str)

        if lat is None or lon is None:
            df.loc[index, 'Address'] = "Invalid Coordinates"
            unsuccessful_count += 1
            # Update the row status in the current_status variable
            current_status_var.set(f"Status: Processing row {index + 1} of {total_rows}")
            current_status_var.set("Status: Invalid Coordinates")
            root.update_idletasks()
            update_progress(index, successful_count, unsuccessful_count, total_rows)
            time.sleep(0.05)  # Small delay for visual effect
            continue

        if pd.isna(lat) or pd.isna(lon) or not (isinstance(lat, (int, float)) and isinstance(lon, (int, float))):
            df.loc[index, 'Address'] = "No Coordinates"
            unsuccessful_count += 1
            current_status_var.set("Status: No Coordinates")
            update_progress(index, successful_count, unsuccessful_count, total_rows)
            root.update_idletasks()
            time.sleep(0.05)
            continue

        # Show "Geocoding..." message
        current_status_var.set("Status: Geocoding...")
        root.update_idletasks()
        
        address = geocode_address(lat, lon)
        df.loc[index, 'Address'] = address

        # Update address display
        address_var.set(f"Address: {address[:100]}{'...' if len(address) > 100 else ''}")

        if (address == "Address not found" or address == "Geocoding failed" or address == "Error during geocoding" or
                address == "No Coordinates" or address == "Invalid Coordinates"):
            unsuccessful_count += 1
            current_status_var.set(f"Status: {address}")
        else:
            successful_count += 1
            current_status_var.set("Status: Success")

        update_progress(index, successful_count, unsuccessful_count, total_rows)
        root.update_idletasks()
        time.sleep(0.05)  # Small delay for visual effect

    try:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Process complete!\n\nSuccessfully geocoded: {successful_count} addresses\nUnsuccessful: {unsuccessful_count} addresses\n\nResults saved to {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving Excel file: {e}")
    
    reset_ui()

def update_progress(current, successful, unsuccessful, total):
    """Updates progress bar and statistics"""
    progress_bar["value"] = current + 1
    
    # Update statistics labels with colors based on values
    success_label.config(text=f"{successful}")
    failed_label.config(text=f"{unsuccessful}")
    total_processed_label.config(text=f"{current + 1}/{total}")
    
    # Calculate completion percentage
    completion = int(((current + 1) / total) * 100)
    progress_percent.config(text=f"{completion}%")

def reset_ui():
    """Reset UI elements after processing"""
    process_button.config(state=tk.NORMAL)
    browse_button.config(state=tk.NORMAL)
    processing_frame.pack_forget()
    current_status_var.set("Ready")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)
        
        # Extract filename for display
        filename = os.path.basename(file_path)
        file_name_var.set(f"Selected: {filename}")
        
        try:
            df = pd.read_excel(file_path)
            column_names = df.columns.tolist()
            lon_col_dropdown['values'] = column_names
            lat_col_dropdown['values'] = column_names
            lon_col_dropdown.config(state="readonly")
            lat_col_dropdown.config(state="readonly")
            
            # Enable the process button
            process_button.config(state=tk.NORMAL)
            
            # Show column selection frame
            column_frame.pack(fill=tk.X, padx=20, pady=10)
            
            # If there are columns with lat/long in their names, try to auto-select them
            lat_candidates = [col for col in column_names if any(x in col.lower() for x in ['lat', 'latitude'])]
            lon_candidates = [col for col in column_names if any(x in col.lower() for x in ['lon', 'lng', 'long', 'longitude'])]
            
            if lat_candidates:
                lat_col_var.set(lat_candidates[0])
            if lon_candidates:
                lon_col_var.set(lon_candidates[0])
                
        except Exception as e:
            messagebox.showerror("Error", f"Error reading Excel file: {e}")
            file_path_entry.delete(0, tk.END)
            lon_col_dropdown['values'] = []
            lat_col_dropdown['values'] = []

def create_tooltip(widget, text):
    """Create a tooltip for a given widget"""
    def enter(event):
        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 25
        
        # Create a toplevel window
        tooltip = tk.Toplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{x}+{y}")
        
        label = ttk.Label(tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()
        
        widget.tooltip = tooltip
        
    def leave(event):
        if hasattr(widget, "tooltip"):
            widget.tooltip.destroy()
            
    widget.bind("<Enter>", enter)
    widget.bind("<Leave>", leave)

# Set up the main window
root = tk.Tk()
root.title("GeoWizard - Advanced Geocoding Tool")
root.geometry("750x680")
root.resizable(True, True)
root.configure(bg="#f5f5f5")

# Alternative approach: Use a simple text-based icon that doesn't require image processing
# This avoids potential issues with Base64 decoding or PIL
try:
    # Set window icon using Unicode character for location/map pin
    root.iconbitmap("") # Remove any existing icon
    
    # Create a custom icon using a canvas
    icon_size = 32
    icon = tk.Canvas(root, width=icon_size, height=icon_size, bg="#4a6984", highlightthickness=0)
    # Draw a map pin shape
    icon.create_oval(8, 8, 24, 24, fill="white", outline="")
    icon.create_polygon(16, 16, 10, 28, 22, 28, fill="white", outline="")
    
    # Convert canvas to a PhotoImage
    icon.update()
    icon_data = tk.PhotoImage(width=icon_size, height=icon_size)
    icon_data.put(icon.postscript(colormode="color"), to=(0, 0))
    
    # Set the icon
    root.iconphoto(True, icon_data)
except:
    # If icon creation fails, simply continue without an icon
    pass

# Apply a modern themed style
style = ThemedStyle(root)
style.set_theme("arc")  # Modern, clean theme

# Create custom styles
style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=10)
style.configure("TLabel", font=("Segoe UI", 10), background="#f5f5f5")
style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"), background="#f5f5f5")
style.configure("Subtitle.TLabel", font=("Segoe UI", 12), background="#f5f5f5")
style.configure("Success.TLabel", foreground="green", background="#f5f5f5")
style.configure("Error.TLabel", foreground="red", background="#f5f5f5")
style.configure("Info.TLabel", foreground="#0066cc", background="#f5f5f5")
style.configure("Header.TFrame", background="#4a6984")
style.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
style.configure("TEntry", padding=5)

# Header with app title
header_frame = ttk.Frame(root, style="Header.TFrame")
header_frame.pack(fill=tk.X, ipady=15)

# Create a simple icon for the header using a Canvas
header_icon_size = 32
header_icon = tk.Canvas(header_frame, width=header_icon_size, height=header_icon_size, 
                      bg="#4a6984", highlightthickness=0)
# Draw a location pin
header_icon.create_oval(8, 8, 24, 24, fill="white", outline="")
header_icon.create_polygon(16, 16, 10, 28, 22, 28, fill="white", outline="")

# Title and icon layout
title_frame = ttk.Frame(header_frame, style="Header.TFrame")
title_frame.pack(pady=5)

header_icon.pack(side=tk.LEFT, padx=(20, 10))

app_title = ttk.Label(title_frame, text="GeoWizard", font=("Segoe UI", 18, "bold"), foreground="white", background="#4a6984")
app_title.pack(side=tk.LEFT)

app_subtitle = ttk.Label(header_frame, text="Convert Coordinates to Addresses with Ease", foreground="white", background="#4a6984")
app_subtitle.pack()

# Main content area
content_frame = ttk.Frame(root)
content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# File selection section (as a card)
file_frame = ttk.Frame(content_frame, style="Card.TFrame")
file_frame.pack(fill=tk.X, padx=20, pady=20)

file_section_title = ttk.Label(file_frame, text="1. Select Your Excel File", style="Title.TLabel")
file_section_title.pack(anchor="w", padx=15, pady=(15, 5))

file_instructions = ttk.Label(file_frame, text="Choose an Excel file containing coordinate data to geocode", foreground="#555555")
file_instructions.pack(anchor="w", padx=15, pady=(0, 10))

file_browser_frame = ttk.Frame(file_frame)
file_browser_frame.pack(fill=tk.X, padx=15, pady=5)

file_path_entry = ttk.Entry(file_browser_frame, width=50, font=("Segoe UI", 10))
file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

browse_button = ttk.Button(file_browser_frame, text="Browse Files", command=browse_file)
browse_button.pack(side=tk.RIGHT)

file_name_var = tk.StringVar()
file_name_label = ttk.Label(file_frame, textvariable=file_name_var, foreground="#0066cc")
file_name_label.pack(anchor="w", padx=15, pady=(5, 15))

# Column selection section
column_frame = ttk.Frame(content_frame, style="Card.TFrame")
# Not packing yet - will be displayed after file selection

column_section_title = ttk.Label(column_frame, text="2. Select Coordinate Columns", style="Title.TLabel")
column_section_title.pack(anchor="w", padx=15, pady=(15, 5))

column_instructions = ttk.Label(column_frame, text="Select the columns containing latitude and longitude data", foreground="#555555")
column_instructions.pack(anchor="w", padx=15, pady=(0, 10))

lat_frame = ttk.Frame(column_frame)
lat_frame.pack(fill=tk.X, padx=15, pady=5)

lat_col_label = ttk.Label(lat_frame, text="Latitude Column:")
lat_col_label.pack(side=tk.LEFT, padx=(0, 10))

lat_col_var = tk.StringVar(root)
lat_col_dropdown = ttk.Combobox(lat_frame, textvariable=lat_col_var, state="disabled", width=30)
lat_col_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)

lon_frame = ttk.Frame(column_frame)
lon_frame.pack(fill=tk.X, padx=15, pady=5)

lon_col_label = ttk.Label(lon_frame, text="Longitude Column:")
lon_col_label.pack(side=tk.LEFT, padx=(0, 10))

lon_col_var = tk.StringVar(root)
lon_col_dropdown = ttk.Combobox(lon_frame, textvariable=lon_col_var, state="disabled", width=30)
lon_col_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)

button_frame = ttk.Frame(column_frame)
button_frame.pack(fill=tk.X, padx=15, pady=15)

process_button = ttk.Button(button_frame, text="Start Geocoding", command=process_file_threaded, state=tk.DISABLED)
process_button.pack(pady=5)
create_tooltip(process_button, "Begin the geocoding process")

# Status indicators section
status_frame = ttk.Frame(content_frame, style="Card.TFrame")
# Will be packed during processing

status_section_title = ttk.Label(status_frame, text="Processing Status", style="Title.TLabel")
status_section_title.pack(anchor="w", padx=15, pady=(15, 5))

stats_frame = ttk.Frame(status_frame)
stats_frame.pack(fill=tk.X, padx=15, pady=5)

# Status indicators with more visual appeal
ttk.Label(stats_frame, text="Success: ", foreground="green").grid(row=0, column=0, sticky="w", padx=5, pady=2)
success_label = ttk.Label(stats_frame, text="0", style="Success.TLabel")
success_label.grid(row=0, column=1, sticky="w", padx=5, pady=2)

ttk.Label(stats_frame, text="Failed: ", foreground="red").grid(row=0, column=2, sticky="w", padx=(20, 5), pady=2)
failed_label = ttk.Label(stats_frame, text="0", style="Error.TLabel")
failed_label.grid(row=0, column=3, sticky="w", padx=5, pady=2)

ttk.Label(stats_frame, text="Processed: ", foreground="#0066cc").grid(row=0, column=4, sticky="w", padx=(20, 5), pady=2)
total_processed_label = ttk.Label(stats_frame, text="0/0", style="Info.TLabel")
total_processed_label.grid(row=0, column=5, sticky="w", padx=5, pady=2)

# Progress bar with percentage
progress_frame = ttk.Frame(status_frame)
progress_frame.pack(fill=tk.X, padx=15, pady=(5, 15))

progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

progress_percent = ttk.Label(progress_frame, text="0%")
progress_percent.pack(side=tk.RIGHT)

# Processing details panel (shown during geocoding)
processing_frame = ttk.Frame(content_frame, style="Card.TFrame")
# Will be packed during processing

processing_title = ttk.Label(processing_frame, text="Live Geocoding Information", style="Subtitle.TLabel")
processing_title.pack(anchor="w", padx=15, pady=(15, 10))

current_status_var = tk.StringVar(value="Ready")
current_status = ttk.Label(processing_frame, textvariable=current_status_var, foreground="#0066cc")
current_status.pack(anchor="w", padx=15, pady=2)

coords_var = tk.StringVar(value="Coordinates: ")
coords_label = ttk.Label(processing_frame, textvariable=coords_var)
coords_label.pack(anchor="w", padx=15, pady=2)

address_var = tk.StringVar(value="Address: ")
address_label = ttk.Label(processing_frame, textvariable=address_var)
address_label.pack(anchor="w", padx=15, pady=2)

# Footer with attribution
footer_frame = ttk.Frame(root, style="Header.TFrame")
footer_frame.pack(fill=tk.X, side=tk.BOTTOM, ipady=5)

footer_text = ttk.Label(footer_frame, 
                       text="Created by Jacques Boucher (jjrboucher@gmail.com) • Enhanced UI by Claude 3.7 Sonnet • Last updated: 16 May 2025",
                       foreground="white", background="#4a6984", font=("Segoe UI", 8))
footer_text.pack(pady=5)

root.mainloop()
