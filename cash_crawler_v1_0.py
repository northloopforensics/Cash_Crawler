import pandas as pd
import matplotlib.pyplot as plt
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
import os
import zipfile
import datetime
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap import ttk
import base64
import sys
import shutil

# To Do List:


# Determine the base directory
if getattr(sys, 'frozen', False):
    # If the application is frozen (e.g., packaged with PyInstaller)
    script_dir = os.path.dirname(sys.executable)
else:
    # If the application is not frozen
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Path to the template
template_path = os.path.join(script_dir, 'report_template.html')

# Check if the template file exists
if not os.path.exists(template_path):
    raise FileNotFoundError(f"Template file not found at {template_path}")

# Icon images
nlc_icon = "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"
zip_icon = "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"
folder_icon = "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"
def on_closing():
    sys.exit()
def generate_html_id(name):
    # Generate a safe HTML ID by replacing non-word characters with underscores
    return re.sub(r'\W+', '_', name)

# Collect data provided for associated accounts into a list of dictionaries
def collect_associate_acct_info(reportrootdirectory):
    associate_files = []
    associate_data = {}

    for dir in os.listdir(reportrootdirectory):
        if os.path.isdir(os.path.join(reportrootdirectory, dir)):
            if "sender_receiver_accounts" in dir.lower():
                associate_dir = os.path.join(reportrootdirectory, dir)
                for file in os.listdir(associate_dir):
                    if file.endswith('.xlsx'):
                        associate_files.append(os.path.join(associate_dir, file))

    if associate_files:
        for file in associate_files:
            # print(f"Processing {file} data...")
            tokenname = os.path.basename(file).split('-')[4]
            tokenname = tokenname.split('.')[0]
            # print(f"Processing {tokenname} data...")
            try:
                assoc_data = read_excel_datasets(file)
                account_data = collect_account_data(assoc_data)
                associate_data[tokenname] = account_data  # Store the account data in the dictionary
            except IndexError:
                print(f"Error reading {tokenname} data. Index Error related to Excel sheet")
                continue

    return associate_data


def collect_account_data(datasets):
    account_ids = {
        "Active Account Token": [],
        "Identity Verification Names": [],
        "Dates of Birth": [],
        "Social Security Numbers": [],
        "Addresses": [],
        "Aliases": []
    }

    # Process 'Identity History'
    identity_history = datasets.get('Identity History', {})
    if identity_history:
        # regex to get active account token
        active_account_token = re.search(r'C_[a-zA-Z0-9]{9}', str(identity_history))
        if active_account_token:
            account_ids["Active Account Token"].append(active_account_token.group())
        # Process 'Identity Verification Info' within 'Identity History'
        identity_verification_info_df = identity_history.get('Identity Verification Info', pd.DataFrame())
        if not identity_verification_info_df.empty:
            # Set the first row as the header
            identity_verification_info_df.columns = identity_verification_info_df.iloc[0]
            identity_verification_info_df = identity_verification_info_df[1:].reset_index(drop=True)
            # Clean column names
            identity_verification_info_df.columns = identity_verification_info_df.columns.astype(str).str.strip()

            # Collect identity verification data
            if 'Identity Verification Name' in identity_verification_info_df.columns:
                account_ids["Identity Verification Names"].extend(
                    identity_verification_info_df['Identity Verification Name'].dropna().unique()
                )
            else:
                print("Column 'Identity Verification Name' not found in 'Identity Verification Info'")

            if 'Date of Birth' in identity_verification_info_df.columns:
                account_ids["Dates of Birth"].extend(
                    identity_verification_info_df['Date of Birth'].dropna().unique()
                )
            else:
                print("Column 'Date of Birth' not found in 'Identity Verification Info'")


            if 'Full SSN (may be provided by third party)' in identity_verification_info_df.columns:
                account_ids["Social Security Numbers"].extend(
                    identity_verification_info_df['Full SSN (may be provided by third party)'].dropna().unique()
                )
            else:
                print("Column 'Full SSN (may be provided by third party)' not found in 'Identity Verification Info'")
        else:
            print("'Identity Verification Info' DataFrame is empty or missing.")
    else:
        print("'Identity History' dataset is missing.")

    # Process 'Address History'
    address_history_df = datasets.get('Address History', pd.DataFrame())
    if not address_history_df.empty:
        # Clean column names
        address_history_df.columns = address_history_df.columns.astype(str).str.strip()

        # Check if all required columns are present
        required_columns = ['Street1', 'City', 'State', 'Zip']
        if all(col in address_history_df.columns for col in required_columns):
            # Concatenate address components into a single string
            address_history_df['Full Address'] = address_history_df.apply(
                lambda row: f"{row['Street1']}, {row['City']}, {row['State']} {row['Zip']}", axis=1
            )
            account_ids["Addresses"].extend(address_history_df['Full Address'].dropna().unique())
            account_ids["Addresses"] = list(set(account_ids["Addresses"]))  # Remove duplicates
        else:
            missing_columns = [col for col in required_columns if col not in address_history_df.columns]
            print(f"Missing columns in 'Address History': {', '.join(missing_columns)}")
    else:
        print("'Address History' DataFrame is empty or missing.")

    # Process 'Alias History'
    alias_history_df = datasets.get('Alias History', pd.DataFrame())
    if not alias_history_df.empty:
        # Clean column names
        alias_history_df.columns = alias_history_df.columns.astype(str).str.strip()

        # Collect alias data
        if 'Alias' in alias_history_df.columns and 'Source' in alias_history_df.columns:
            alias_data = alias_history_df[['Alias', 'Source']].dropna().to_dict('records')
            formatted_aliases = [f"{item['Alias']} - {item['Source']}" for item in alias_data]
            account_ids["Aliases"].extend(formatted_aliases)
            account_ids["Aliases"] = list(set(account_ids["Aliases"]))  # Remove duplicates
        else:
            print("Columns 'Alias' or 'Source' not found in 'Alias History'")
    else:
        print("'Alias History' DataFrame is empty or missing.")
    return account_ids
# Extract the ZIP file to a temporary directory
def extract_zip(zip_path, password=None, output_dir=None):
    foldername = os.path.splitext(os.path.basename(zip_path))[0]

    with zipfile.ZipFile(zip_path) as zf:
        if password:
            zf.setpassword(password.encode())
        temp_dir = os.path.join(output_dir, foldername)
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        zf.extractall(temp_dir)
        return temp_dir
    
def find_ID_photos(directory):
    photo_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.jpg') or file.endswith('.jpeg') or file.endswith('.png'):
                photo_files.append(os.path.join(root, file))
    return photo_files   

# Find the Excel file in the root of the extracted directory
def find_excel_file(directory):
    for file in os.listdir(directory):
        if file.endswith('.xlsx'):
            return os.path.join(directory, file)
    return None  # Return None if no .xlsx file is found

def read_excel_datasets(file_path):
    import pandas as pd

    # Load the entire Excel file into a DataFrame without headers
    df = pd.read_excel(file_path, header=None)

    datasets = {}
    nrows = df.shape[0]
    i = 0

    while i < nrows:
        # Skip any empty rows to find the dataset name
        while i < nrows and df.iloc[i].isnull().all():
            i += 1
        if i >= nrows:
            break

        # Get the dataset name
        name = df.iloc[i, 0]
        i += 1  # Move to the next row

        # Check for end of file
        if pd.isna(name):
            break

        # Handle 'Identity History' separately due to its structure
        if name == 'Identity History':
            # Extract 'Active Account Token'
            if i + 1 < nrows:
                active_account_token_label = df.iloc[i, 0]
                active_account_token_value = df.iloc[i + 1, 0]
                active_account_token_df = pd.DataFrame({
                    active_account_token_label: [active_account_token_value]
                })
                i += 2  # Move past the 'Active Account Token' label and value
            else:
                active_account_token_df = pd.DataFrame()

            # Extract identity verification data
            if i < nrows:
                headers = df.iloc[i].tolist()
                i += 1  # Move to the first data row
                data_rows = []

                while i < nrows and not df.iloc[i].isnull().all():
                    data_rows.append(df.iloc[i].tolist())
                    i += 1

                # Check if data_rows is not empty before creating DataFrame
                if data_rows:
                    identity_info_df = pd.DataFrame(data_rows, columns=headers)
                    identity_info_df = identity_info_df.fillna('')  # Replace NaN with empty string
                else:
                    identity_info_df = pd.DataFrame(columns=headers)
            else:
                identity_info_df = pd.DataFrame()

            # Store both DataFrames in the datasets dictionary
            datasets[name] = {
                'Active Account Token': active_account_token_df,
                'Identity Verification Info': identity_info_df
            }
        else:
            # Extract other datasets normally
            if i < nrows:
                headers = df.iloc[i].tolist()
                i += 1  # Move to the first data row
                data_rows = []

                while i < nrows and not df.iloc[i].isnull().all():
                    data_rows.append(df.iloc[i].tolist())
                    i += 1

                # Check if data_rows is not empty before creating DataFrame
                if data_rows:
                    data_df = pd.DataFrame(data_rows, columns=headers)
                    data_df = data_df.fillna('')  # Replace NaN with empty string
                else:
                    data_df = pd.DataFrame(columns=headers)

                datasets[name] = data_df

        # Skip any empty rows before the next dataset
        while i < nrows and df.iloc[i].isnull().all():
            i += 1

    return datasets    
def join_dataframes(dataframes):
    import pandas as pd

    # List of datasets to include
    required_datasets = [
        "Attempted P2P Payments",
        "Attempted Transfers",
        "Attempted Transactions",
        "Attempted Cash App Pays",
        "Attempted Bitcoin Transactions"
    ]

    df_list = []

    for dataset_name in required_datasets:
        if dataset_name in dataframes:
            df = dataframes[dataset_name]
            if isinstance(df, pd.DataFrame):
                df = df.copy()

                # Consolidate 'Total' and 'Amount' into 'Amount'
                if 'Total' in df.columns:
                    df.rename(columns={'Total': 'Amount'}, inplace=True)
                # If 'Amount' exists, leave it as is

                df['Source'] = dataset_name  # Add source column

                # Remove duplicate columns
                df = df.loc[:, ~df.columns.duplicated()]

                df_list.append(df)
            else:
                print(f"Dataset '{dataset_name}' is not a DataFrame.")
        else:
            print(f"Dataset '{dataset_name}' not found in the provided data.")

    if df_list:
        # Concatenate the selected DataFrames
        joined_df = pd.concat(df_list, ignore_index=True)

        # Replace NaN entries with empty strings
        joined_df = joined_df.fillna('')

        # Reorder columns to start with 'Date', 'Source', 'Amount'
        cols = joined_df.columns.tolist()
        desired_order = ['Date', 'Source', 'Amount']
        other_cols = [col for col in cols if col not in desired_order]
        joined_df = joined_df[desired_order + other_cols]

        # print("Successfully joined the selected datasets.")
    else:
        print("No datasets were joined; none of the specified datasets were found.")
        joined_df = pd.DataFrame()  # Return an empty DataFrame if none found

    return joined_df

          
def bitcoin_transactions(dataframes):
    # Ensure the DataFrame for attempted Bitcoin transactions exists
    if 'Attempted Bitcoin Transactions' not in dataframes:
        raise KeyError("'Attempted Bitcoin Transactions' DataFrame not found in the provided dataframes.")

    df = dataframes['Attempted Bitcoin Transactions']
    df.columns = map(str, df.columns)
    try:
        # Check if either 'Wallet Address' is part of any column name
        wallet_column = None
        for column in df.columns:
            if 'Wallet Address' in column:
                wallet_column = column
    except KeyError as e:
        print(f"Error processing Bitcoin transactions: {e}")
        return pd.DataFrame
    try:
        bitcoin_df = df.copy()
        # Remove 'BTC ' prefix and commas, then convert to float
        bitcoin_df['Amount'] = bitcoin_df['Amount'].str.replace('BTC ', '').str.replace(',', '', regex=False).astype(float)
                
        # Fill empty cells with a placeholder value
        bitcoin_df[wallet_column] = bitcoin_df[wallet_column].fillna('No Wallet Address - Generally Indicates Purchase or Sale of Bitcoin')
        bitcoin_df[wallet_column] = bitcoin_df[wallet_column].replace('', 'No Wallet Address - Generally Indicates Purchase or Sale of Bitcoin')

        # Group by wallet_column and calculate aggregates
        bitcoin_group = bitcoin_df.groupby(wallet_column)
        bitcoin_stats = bitcoin_group['Amount'].agg(
            Total_Amount='sum', 
            Number_of_Transactions='count',
            Average_Transaction_Amount='mean',
            Maximum_Transaction_Amount='max',
            Minimum_Transaction_Amount='min'
        ).reset_index()

        # Ensure Number_of_Transactions is treated as an integer
        bitcoin_stats['Number_of_Transactions'] = bitcoin_stats['Number_of_Transactions'].astype(int)

        # Sort by 'Total_Amount' in descending order while it's still numeric
        bitcoin_stats = bitcoin_stats.sort_values(by='Total_Amount', ascending=False)

        # Format amounts as currency AFTER sorting
        for col in [
            'Total_Amount', 'Average_Transaction_Amount',
            'Maximum_Transaction_Amount', 'Minimum_Transaction_Amount' 
        ]:
            bitcoin_stats[col] = bitcoin_stats[col].apply(lambda x: f'BTC {x:,.8f}')
        
        print("Bitcoin Stats:", bitcoin_stats)
        return bitcoin_stats
    except KeyError as e:
        print(f"Error processing Bitcoin transactions: {e}")
    
def sender_totals(dataframes):
    if "Attempted P2P Payments" in dataframes:
        sent_payments_df = dataframes["Attempted P2P Payments"]

        # Ensure that 'Total' and 'Sender' columns exist
        if 'Total' in sent_payments_df.columns and 'Sender' in sent_payments_df.columns:
            # Clean and convert the "Total" column to numerical values
            sent_payments_df['Total'] = sent_payments_df['Total'].astype(str)
            sent_payments_df['Total'] = sent_payments_df['Total'].str.replace(
                'USD ', '', regex=False
            ).str.replace(',', '', regex=False).astype(float)

            # Group by 'Sender' and calculate aggregates
            send_group = sent_payments_df.groupby('Sender')
            send_stats = send_group['Total'].agg(
                Total_Amount='sum',
                Number_of_Transactions='count',
                Average_Transaction_Amount='mean',
                Maximum_Transaction_Amount='max',
                Minimum_Transaction_Amount='min'
            ).reset_index()

            # Sort by 'Total_Amount' in descending order while it's still numeric
            send_stats = send_stats.sort_values(by='Total_Amount', ascending=False)

            # Format amounts as currency AFTER sorting
            for col in [
                'Total_Amount', 'Average_Transaction_Amount',
                'Maximum_Transaction_Amount', 'Minimum_Transaction_Amount'
            ]:
                send_stats[col] = send_stats[col].apply(lambda x: f'${x:,.2f}')

            return send_stats
        else:
            print("Columns 'Total' or 'Sender' not found in the dataset.")
            return pd.DataFrame()
    else:
        print("Dataset 'Attempted P2P Payments' not found.")
        return pd.DataFrame()
    
def plot_top_sender_totals(df):
    df_plot = df.copy()
    df_plot['Total_Amount'] = df_plot['Total_Amount'].replace(r'[\$,]', '', regex=True).astype(float)
    top_senders = df_plot.head(10)
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.barh(top_senders['Sender'], top_senders['Total_Amount'], color='purple')
    ax.set_xlabel('Total Amount Sent')
    ax.set_ylabel('Sender')
    ax.set_title('Top 10 Senders by Total Amount')
    ax.invert_yaxis()
    plt.tight_layout()
    return fig

def plot_sender_totals(df, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    # Convert 'Total_Amount' back to float for plotting
    df_plot = df.copy()
    df_plot['Total_Amount'] = df_plot['Total_Amount'].replace(r'[\$,]', '', regex=True).astype(float)

    # Plot the recipient totals as a horizontal bar graph
    senders = df_plot
    plt.figure(figsize=(14, 8))
    plt.barh(senders['Sender'], senders['Total_Amount'], color='skyblue')
    plt.xlabel('Total Amount Received')
    plt.ylabel('Sender')
    plt.title('Senders by Total Amount')
    plt.gca().invert_yaxis()  # Highest values at the top
    plt.tight_layout()
    sender_chart_path = os.path.join(output_folder, 'senders_chart.png')
    plt.savefig(sender_chart_path)
    plt.close()
    return sender_chart_path

def recipient_totals(dataframes):
    if "Attempted P2P Payments" in dataframes:
        received_payments_df = dataframes["Attempted P2P Payments"]

        # Ensure that 'Total' and 'Recipient' columns exist
        if 'Total' in received_payments_df.columns and 'Recipient' in received_payments_df.columns:
            # Clean and convert the "Total" column to numerical values
            received_payments_df['Total'] = received_payments_df['Total'].astype(str)
            received_payments_df['Total'] = received_payments_df['Total'].str.replace(
                'USD ', '', regex=False
            ).str.replace(',', '', regex=False).astype(float)

            # Group by 'Recipient' and calculate aggregates
            recipient_group = received_payments_df.groupby('Recipient')
            recipient_stats = recipient_group['Total'].agg(
                Total_Amount='sum',
                Number_of_Transactions='count',
                Average_Transaction_Amount='mean',
                Maximum_Transaction_Amount='max',
                Minimum_Transaction_Amount='min'
            ).reset_index()

            # Sort by 'Total_Amount' in descending order while it's still numeric
            recipient_stats = recipient_stats.sort_values(by='Total_Amount', ascending=False)

            # Format amounts as currency AFTER sorting
            for col in [
                'Total_Amount', 'Average_Transaction_Amount',
                'Maximum_Transaction_Amount', 'Minimum_Transaction_Amount'
            ]:
                recipient_stats[col] = recipient_stats[col].apply(lambda x: f'${x:,.2f}')

            
            # Print the top recipients
            # print(recipient_stats.head())
            return recipient_stats
        else:
            print("Columns 'Total' or 'Recipient' not found in the dataset.")
            return pd.DataFrame()
    else:
        print("Dataset 'Attempted P2P Payments' not found.")
        return pd.DataFrame()

def plot_recipient_totals(df):
    df_plot = df.copy()
    df_plot['Total_Amount'] = df_plot['Total_Amount'].replace(r'[\$,]', '', regex=True).astype(float)
    top_recipients = df_plot.head(10)
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.barh(top_recipients['Recipient'], top_recipients['Total_Amount'], color='darkorange')
    ax.set_xlabel('Total Amount Received')
    ax.set_ylabel('Recipient')
    ax.set_title('Top 10 Recipients by Total Amount')
    ax.invert_yaxis()
    plt.tight_layout()
    return fig

def read_ip_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheet_name = None
    for name in xls.sheet_names:
        if "IP DATA" in name.upper():
            sheet_name = name
            break

    if sheet_name is None:
        print("No sheet with 'IP DATA' found.")
        return pd.DataFrame()

    # Load the specific sheet into a DataFrame, skipping the first row
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=1)

    # Check if 'Client IP Address' column exists
    if 'Client IP Address' not in df.columns:
        print("'Client IP Address' column not found.")
        return pd.DataFrame()

    # Extract the 'Client IP Address' column
    ip_addresses = df['Client IP Address']

    # Count the occurrences of each IP address
    ip_counts = ip_addresses.value_counts().reset_index()
    ip_counts.columns = ['IP Address', 'Count']

    # Get the 10 most common IP addresses
    top_ips = ip_counts.head(10)

    # print(top_ips)
    return top_ips


def create_html_report(
    account_ids, sender_data, sender_chart, recipient_data,
    recipient_chart, ip_data, id_photos, output_path, joined_df, associate_data, case_info, top_bitcoin_stats
):
    current_year = datetime.datetime.now().year
    if getattr(sys, 'frozen', False):
        # If the application is frozen (e.g., packaged with PyInstaller)
        script_dir = os.path.dirname(sys.executable)
    else:
        # If the application is not frozen
        script_dir = os.path.dirname(os.path.abspath(__file__))

    # Path to the template directory
    template_dir = script_dir

    # Create a Jinja2 environment with the template directory
    env = Environment(loader=FileSystemLoader(template_dir))
    # Load the template
    try:
        template = env.get_template('report_template.html')
    except TemplateNotFound:
        raise FileNotFoundError(f"Template file not found: {os.path.join(template_dir, 'report_template.html')}")

    # Create the Images folder in the same directory as the output HTML file
    images_dir = os.path.join(os.path.dirname(output_path), 'Images')
    os.makedirs(images_dir, exist_ok=True)

    # Copy ID photos to the Images folder
    copied_id_photos = []
    for photo in id_photos:
        photo_filename = os.path.basename(photo)
        destination = os.path.join(images_dir, photo_filename)
        shutil.copy(photo, destination)
        copied_id_photos.append(os.path.join('Images', photo_filename))

    sender_chart_path = os.path.join(images_dir, 'sender_chart.png')
    if sender_chart:
        sender_chart.savefig(sender_chart_path)
    # Save recipient chart as an image file
    recipient_chart_path = os.path.join(images_dir, 'recipient_chart.png')
    if recipient_chart:
        recipient_chart.savefig(recipient_chart_path)
    
    # Determine if data is available
    has_sender_data = sender_data is not None and not sender_data.empty
    has_recipient_data = recipient_data is not None and not recipient_data.empty
    has_ip_data = ip_data is not None and not ip_data.empty
    has_joined_data = joined_df is not None and not joined_df.empty
    has_bitcoin_stats = top_bitcoin_stats is not None and not top_bitcoin_stats.empty
    #format the sender and recipient data
    if has_sender_data:
        sender_data.columns = sender_data.columns.str.replace('_', ' ')
    if has_recipient_data:
        recipient_data.columns = recipient_data.columns.str.replace('_', ' ')
    if has_bitcoin_stats:
        top_bitcoin_stats.columns = top_bitcoin_stats.columns.str.replace('_', ' ')


    html_content = template.render(
        sender_chart=os.path.join('Images', 'sender_chart.png') if sender_chart else None,
        sender_data=sender_data,
        has_sender_data=has_sender_data,
        recipient_chart=os.path.join('Images', 'recipient_chart.png') if recipient_chart else None,
        recipient_data=recipient_data,
        has_recipient_data=has_recipient_data,
        ip_data=ip_data,
        has_ip_data=has_ip_data,
        id_photos=copied_id_photos,
        account_ids=account_ids,
        current_year=current_year,
        joined_data=joined_df,
        has_joined_data=has_joined_data,
        associate_data=associate_data,
        case_info=case_info,
        top_bitcoin_stats=top_bitcoin_stats
    )
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)


def format_associate_data(associate_data):
    formatted_data = []
    for tokenname, record in associate_data.items():
        if isinstance(record, dict):
            flat_data = {'Token Name': tokenname}  # Include the token name
            for key, value in record.items():
                if key == "Active Account Token":
                    continue  # Skip the "Active Account Token"
                if isinstance(value, list):
                    if all(isinstance(item, dict) for item in value):
                        # Handle list of dictionaries
                        if key == "Aliases":
                            # Special formatting for Aliases
                            value_strings = [f"{item['Alias']} ({item['Source']})" for item in value]
                            flat_data[key] = '<br>'.join(value_strings)
                        else:
                            value_strings = []
                            for item in value:
                                item_string = '<br>'.join(f"{k}: {v}" for k, v in item.items())
                                value_strings.append(item_string)
                            flat_data[key] = '<br><br>'.join(value_strings)
                    else:
                        # Handle list of strings or other types
                        flat_data[key] = ', '.join(map(str, value))
                else:
                    flat_data[key] = str(value)
            formatted_data.append(flat_data)
    return formatted_data

def select_zip_file():
    file_path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
    if file_path:
        zip_file_entry.delete(0, tk.END)
        zip_file_entry.insert(0, file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        if page2.winfo_ismapped():
            output_folder_entry_zip.delete(0, tk.END)
            output_folder_entry_zip.insert(0, folder_path)
        if page3.winfo_ismapped():
            output_folder_entry_folder.delete(0, tk.END)
            output_folder_entry_folder.insert(0, folder_path)

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_path)

def process_zip_data():
    zip_path = zip_file_entry.get()
    zip_password = zip_password_entry.get()
    output_folder = output_folder_entry_zip.get()

    if not zip_path or not output_folder:
        messagebox.showerror("Error", "Please select both the ZIP file and the output folder.")
        return

    images_dir = os.path.join(output_folder, 'Images')
    os.makedirs(images_dir, exist_ok=True)

    extracted_folder = extract_zip(zip_path, zip_password, output_folder)
    current_year = datetime.datetime.now().year
    file_path = find_excel_file(extracted_folder)
    if not file_path:
        raise FileNotFoundError("No .xlsx file found in the extracted directory")
    
    id_photos = find_ID_photos(extracted_folder)
    dataframes = read_excel_datasets(file_path)
    account_data = collect_account_data(dataframes)
    joined_df = join_dataframes(dataframes)
    top_bitcoin_stats = bitcoin_transactions(dataframes)
    sender_totals_df = sender_totals(dataframes)
    sender_chart = None
    if not sender_totals_df.empty:
        sender_chart = plot_top_sender_totals(sender_totals_df)
    sender_data = sender_totals_df
    
    recipient_totals_df = recipient_totals(dataframes)
    recipient_chart = None
    if not recipient_totals_df.empty:
        recipient_chart = plot_recipient_totals(recipient_totals_df)

    ip_data_df = read_ip_data(file_path)
    associate_data = collect_associate_acct_info(extracted_folder)
    formatted_associate_data = format_associate_data(associate_data)
    
    output_path = os.path.join(output_folder, 'report.html')
    create_html_report(
        account_ids=account_data,
        sender_chart=sender_chart,
        recipient_data=recipient_totals_df,
        recipient_chart=recipient_chart,
        ip_data=ip_data_df,
        id_photos=id_photos,
        output_path=output_path,
        joined_df=joined_df,
        sender_data=sender_data,
        associate_data=formatted_associate_data,
        case_info=saved_case_info,
        top_bitcoin_stats=top_bitcoin_stats
    )
    messagebox.showinfo("Processing Complete", f"Report generated: {output_path}")

def process_folder_data():
    folder_path = folder_entry.get()
    output_folder = output_folder_entry_folder.get()
    if not folder_path or not output_folder:
        messagebox.showerror("Error", "Please select both the folder and the output folder.")
        return
    images_dir = os.path.join(output_folder, 'Images')
    os.makedirs(images_dir, exist_ok=True)
    file_path = find_excel_file(folder_path)
    if not file_path:
        raise FileNotFoundError("No .xlsx file found in the folder")
    id_photos = find_ID_photos(folder_path)
    dataframes = read_excel_datasets(file_path)
    account_data = collect_account_data(dataframes)
    joined_df = join_dataframes(dataframes)

    top_bitcoin_stats = bitcoin_transactions(dataframes)

    sender_totals_df = sender_totals(dataframes)
    sender_chart = None
    if not sender_totals_df.empty:
        sender_chart = plot_top_sender_totals(sender_totals_df)
    sender_data = sender_totals_df
    recipient_totals_df = recipient_totals(dataframes)
    recipient_chart = None
    if not recipient_totals_df.empty:
        recipient_chart = plot_recipient_totals(recipient_totals_df)
    ip_data_df = read_ip_data(file_path)
    associate_data = collect_associate_acct_info(folder_path)
    formatted_associate_data = format_associate_data(associate_data)
    print(formatted_associate_data)
    output_path = os.path.join(output_folder, 'report.html')
    create_html_report(
        account_ids=account_data,
        sender_chart=sender_chart,
        recipient_data=recipient_totals_df,
        recipient_chart=recipient_chart,
        ip_data=ip_data_df,
        id_photos=id_photos,
        output_path=output_path,
        joined_df=joined_df,
        sender_data=sender_data,
        associate_data=formatted_associate_data,
        case_info=saved_case_info,
        top_bitcoin_stats=top_bitcoin_stats
    )
    messagebox.showinfo("Processing Complete", f"Report generated: {output_path}")

# Function to save case data inputs from page0 and move to the next page
def save_inputs_and_next():
    # Save the inputs (if needed, you can process them here)
    case_info = {
        'case_number': case_number.get(),
        'case_name': case_name.get(),
        'investigator_name': investigator_name.get(),
        'additional_info': additional_info_text.get("1.0", tk.END)
    }
    print("Case Info Saved:", case_info)  # For debugging purposes
    global saved_case_info
    saved_case_info = case_info
    # Move to the next page
    show_page(page1)

def show_page(page):
    page.tkraise()

######################################## GUI STUFF ########################################
# Create the main application window
style = Style(theme='darkly')  # Use the 'darkly' theme from ttkbootstrap
root = style.master
root.title("Cash Crawler v1.0")
# Bind the window close event
root.protocol("WM_DELETE_WINDOW", on_closing)
# Variables to store input data
case_number = tk.StringVar()
case_name = tk.StringVar()
investigator_name = tk.StringVar()
additional_info_text = tk.StringVar()

zip_icon_img = tk.PhotoImage(data=base64.b64decode(zip_icon))
folder_icon_img = tk.PhotoImage(data=base64.b64decode(folder_icon))
nlc_icon_img = tk.PhotoImage(data=base64.b64decode(nlc_icon))
# Create a container frame to hold all pages
container = ttk.Frame(root)
container.grid(row=0, column=0, sticky="nsew")

# Configure the grid to expand with window resizing
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Create the pages
splashpage = ttk.Frame(container)
page0 = ttk.Frame(container)
page1 = ttk.Frame(container)
page2 = ttk.Frame(container)
page3 = ttk.Frame(container)

for page in (splashpage, page0, page1, page2, page3):
    page.grid(row=0, column=0, sticky="nsew")

# Splash Page: Display the application title for three seconds and move to page0
# Center the contents
splashpage.columnconfigure(0, weight=1)

# Add the NLC icon image
nlc_icon_label = ttk.Label(splashpage, image=nlc_icon_img)
nlc_icon_label.grid(row=0, column=0, pady=20)

# Existing Labels placed using grid
ttk.Label(splashpage, text="Cash Crawler", font=('Impact', 36)).grid(row=1, column=0, pady=10)
ttk.Label(splashpage, text="A CashApp Production Parser", font=('Helvetica', 16)).grid(row=2, column=0, pady=10)
ttk.Label(splashpage, text="v1.0  ©2024 North Loop Consulting, LLC", font=('Helvetica', 14)).grid(row=3, column=0, pady=10)
ttk.Label(splashpage, text="Loading...", font=('Helvetica', 16)).grid(row=4, column=0, pady=30)

# Navigate to the next page after 3 seconds
splashpage.after(3000, lambda: show_page(page0))

# Page 0: Input page for case information
ttk.Label(page0, text="Provide your case information.").grid(row=0, column=0, columnspan=2, padx=20, pady=10, sticky='ew')
ttk.Label(page0, text="Case Number:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
ttk.Entry(page0, textvariable=case_number).grid(row=1, column=1, padx=10, pady=10, sticky='w')

ttk.Label(page0, text="Case Name:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
ttk.Entry(page0, textvariable=case_name).grid(row=2, column=1, padx=10, pady=10, sticky='w')

ttk.Label(page0, text="Investigator Name:").grid(row=3, column=0, padx=10, pady=10, sticky='e')
ttk.Entry(page0, textvariable=investigator_name).grid(row=3, column=1, padx=10, pady=10, sticky='w')

ttk.Label(page0, text="Additional Information:").grid(row=4, column=0, padx=10, pady=10, sticky='ne')
additional_info_text = tk.Text(page0, height=5, width=40)
additional_info_text.grid(row=4, column=1, padx=10, pady=10, sticky='w')

# Add a "Next" button to the input page
ttk.Button(page0, text="Next", command=save_inputs_and_next).grid(row=5, column=0, columnspan=2, pady=20)

# Page 1: Ask if the user has a ZIP file or a folder
ttk.Label(page1, text="Select the data source for the CashApp production.").grid(row=0, column=0, columnspan=2, pady=20)
# Center the buttons with icons and text
ttk.Button(page1, text="Encrypted ZIP File", image=zip_icon_img, compound="top", command=lambda: show_page(page2)).grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
ttk.Button(page1, text="Folder", image=folder_icon_img, compound="top", command=lambda: show_page(page3)).grid(row=1, column=1, padx=20, pady=20, sticky="nsew")
ttk.Button(page1, text="Back", command=lambda: show_page(page0)).grid(row=4, column=0, columnspan=3, pady=10)

# Configure the grid to expand with window resizing
page0.grid_rowconfigure(0, weight=1)
page0.grid_rowconfigure(1, weight=1)
page0.grid_rowconfigure(2, weight=1)
page0.grid_rowconfigure(3, weight=1)
page0.grid_rowconfigure(4, weight=1)
page0.grid_rowconfigure(5, weight=1)
page0.grid_columnconfigure(0, weight=1)
page0.grid_columnconfigure(1, weight=1)
page1.grid_rowconfigure(1, weight=1)
page1.grid_columnconfigure(0, weight=1)
page1.grid_columnconfigure(1, weight=1)
page2.grid_rowconfigure(0, weight=1)
page2.grid_rowconfigure(1, weight=1)
page2.grid_rowconfigure(2, weight=1)
page2.grid_rowconfigure(3, weight=1)
page2.grid_rowconfigure(4, weight=1)
page2.grid_columnconfigure(0, weight=1)
page2.grid_columnconfigure(1, weight=1)
page2.grid_columnconfigure(2, weight=1)
page3.grid_rowconfigure(0, weight=1)
page3.grid_rowconfigure(1, weight=1)
page3.grid_rowconfigure(2, weight=1)
page3.grid_rowconfigure(3, weight=1)
page3.grid_columnconfigure(0, weight=1)
page3.grid_columnconfigure(1, weight=1)
page3.grid_columnconfigure(2, weight=1)


ttk.Label(page2, text="ZIP File:").grid(row=0, column=0, padx=20, pady=20, sticky="e")
zip_file_entry = ttk.Entry(page2, width=50)
zip_file_entry.grid(row=0, column=1, padx=20, pady=20)
ttk.Button(page2, text="Browse", command=select_zip_file).grid(row=0, column=2, padx=10, pady=10)

ttk.Label(page2, text="ZIP Password:").grid(row=1, column=0, padx=20, pady=20, sticky="e")
zip_password_entry = ttk.Entry(page2, show='*', width=50)
zip_password_entry.grid(row=1, column=1, padx=20, pady=20)

ttk.Label(page2, text="Output Folder:").grid(row=2, column=0, padx=20, pady=20, sticky="e")
output_folder_entry_zip = ttk.Entry(page2, width=50)
output_folder_entry_zip.grid(row=2, column=1, padx=10, pady=10)
ttk.Button(page2, text="Browse", command=select_output_folder).grid(row=2, column=2, padx=10, pady=10)

ttk.Button(page2, text="Process Data", command=process_zip_data).grid(row=3, column=0, columnspan=3, pady=20)
ttk.Button(page2, text="Back", command=lambda: show_page(page1)).grid(row=4, column=0, columnspan=3, pady=10)

# Page 3: Process folder
ttk.Label(page3, text="Folder:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
folder_entry = ttk.Entry(page3, width=50)
folder_entry.grid(row=0, column=1, padx=10, pady=10)
ttk.Button(page3, text="Browse", command=select_folder).grid(row=0, column=2, padx=10, pady=10)

ttk.Label(page3, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
output_folder_entry_folder = ttk.Entry(page3, width=50)
output_folder_entry_folder.grid(row=1, column=1, padx=10, pady=10)
ttk.Button(page3, text="Browse", command=select_output_folder).grid(row=1, column=2, padx=10, pady=10)

ttk.Button(page3, text="Process Data", command=process_folder_data).grid(row=2, column=0, columnspan=3, pady=20)
ttk.Button(page3, text="Back", command=lambda: show_page(page1)).grid(row=3, column=0, columnspan=3, pady=10)
# Show the initial page
show_page(splashpage)

# Run the application
root.mainloop()
