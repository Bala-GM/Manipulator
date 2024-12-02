import PySimpleGUI as sg
import os
import shutil

# Function to move a folder to a specified destination
def move_folder(src, dst):
    try:
        shutil.move(src, dst)
        sg.popup(f"Folder '{os.path.basename(src)}' moved successfully to '{dst}'")
    except Exception as e:
        sg.popup_error(f"Error occurred while moving folder: {e}")

# Function to create a new customer name
def create_customer_name(line):
    new_customer = sg.popup_get_text(f"Enter the name of the new customer for '{line}':", title="Create New Customer")
    if new_customer:
        line_path = os.path.join(destination_path, line)
        new_customer_path = os.path.join(line_path, new_customer)
        os.makedirs(new_customer_path, exist_ok=True)
        return new_customer
    else:
        return None

# Define the root directory
root_dir = 'D:/NX_BACKWORK/Feeder Setup_PROCESS/#Output'

# Define the destination path
destination_path = 'D:/NX_BACKWORK/Feeder Setup_PROCESS/#_Feeder Loading Line List'

# Get a list of lines from the destination path
lines = [name for name in os.listdir(destination_path) if os.path.isdir(os.path.join(destination_path, name))]

# Define the layout for the PySimpleGUI window
layout = [
    [sg.Text("Select the folder to move:")],
    [sg.Listbox(values=os.listdir(root_dir), size=(50, 6), key='-FOLDER LIST-', enable_events=True)],
    [sg.Text("Select the line:")],
    [sg.Listbox(values=lines, size=(50, 3), key='-LINE-', enable_events=True)],
    [sg.Text("Select or create a customer name:"), sg.Button("New Customer")],
    [sg.Listbox(values=[], size=(50, 6), key='-CUSTOMER-', enable_events=True)],
    [sg.Radio("Rename Folder", "RADIO1", default=True, key='-RENAME-'), sg.Text("New Folder Name:"), sg.InputText(size=(20,1), key='-NEW FOLDER NAME-')],
    [sg.Button("Move"), sg.Button("Cancel")]
]

# Create the PySimpleGUI window
window = sg.Window("Move Folder", layout)

# Event loop for the PySimpleGUI window
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == "Cancel":
        break
    elif event == "-FOLDER LIST-":
        selected_folder = values['-FOLDER LIST-'][0]
        window['-CUSTOMER-'].update(values=[])
    elif event == "-LINE-":
        line = values['-LINE-'][0]
        customers = os.listdir(os.path.join(destination_path, line)) if os.path.exists(os.path.join(destination_path, line)) else []
        window['-CUSTOMER-'].update(values=customers)
    elif event == "New Customer":
        line = values['-LINE-'][0]
        new_customer = create_customer_name(line)
        if new_customer:
            window['-CUSTOMER-'].update(values=[new_customer])
    elif event == "Move":
        selected_folder = values['-FOLDER LIST-'][0]
        line = values['-LINE-'][0]
        customer = values['-CUSTOMER-'][0]
        rename_enabled = values['-RENAME-']
        new_folder_name = values['-NEW FOLDER NAME-']
        if not selected_folder:
            sg.popup_error("Please select a folder to move")
        elif not line:
            sg.popup_error("Please select a line")
        elif not customer:
            sg.popup_error("Please select or create a customer")
        else:
            # Move the folder
            old_folder_path = os.path.join(root_dir, selected_folder)
            if rename_enabled and new_folder_name:
                # Rename the folder
                new_folder_path = os.path.join(os.path.dirname(old_folder_path), new_folder_name)
                os.rename(old_folder_path, new_folder_path)
            else:
                new_folder_path = old_folder_path
            new_folder_path = os.path.join(destination_path, line, os.path.basename(new_folder_path))
            move_folder(old_folder_path, new_folder_path)
            break

# Close the PySimpleGUI window
window.close()







''' using code 
    # Function to move a folder to a specified destination
    def move_folder(src, dst):
        try:
            shutil.move(src, dst)
            sg.popup(f"Folder '{os.path.basename(src)}' moved successfully to '{dst}'")
        except Exception as e:
            sg.popup_error(f"Error occurred while moving folder: {e}")

    # Function to create a new customer name
    def create_customer_name(line):
        new_customer = sg.popup_get_text(f"Enter the name of the new customer for '{line}':", title="Create New Customer")
        if new_customer:
            line_path = os.path.join(destination_path, line)
            new_customer_path = os.path.join(line_path, new_customer)
            os.makedirs(new_customer_path, exist_ok=True)
            return new_customer
        else:
            return None

    # Define the root directory
    root_dir = 'D:/NX_BACKWORK/Feeder Setup_PROCESS/#Output'

    # Define the destination path
    destination_path = 'D:/NX_BACKWORK/Feeder Setup_PROCESS/#_Feeder Loading Line List'

    # Get a list of lines from the destination path
    lines = [name for name in os.listdir(destination_path) if os.path.isdir(os.path.join(destination_path, name))]

    # Define the layout for the PySimpleGUI window
    layout = [
        [sg.Text("Select the folder to move:")],
        [sg.Listbox(values=os.listdir(root_dir), size=(50, 6), key='-FOLDER LIST-', enable_events=True)],
        [sg.Text("Select the line:")],
        [sg.Listbox(values=lines, size=(50, 3), key='-LINE-', enable_events=True)],
        [sg.Text("Select or create a customer name:"), sg.Button("New Customer")],
        [sg.Listbox(values=[], size=(50, 6), key='-CUSTOMER-', enable_events=True)],
        #[sg.Listbox(values=[], size=(50, 6), key='-CUSTOMER-', enable_events=True), sg.Button("New Customer")],
        [sg.Button("Move"), sg.Button("Cancel")]
    ]

    # Create the PySimpleGUI window
    window = sg.Window("Move Folder", layout)

    # Event loop for the PySimpleGUI window
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        elif event == "-FOLDER LIST-":
            selected_folder = values['-FOLDER LIST-'][0]
            window['-CUSTOMER-'].update(values=[])
        elif event == "-LINE-":
            line = values['-LINE-'][0]
            customers = os.listdir(os.path.join(destination_path, line)) if os.path.exists(os.path.join(destination_path, line)) else []
            window['-CUSTOMER-'].update(values=customers)
        elif event == "New Customer":
            line = values['-LINE-'][0]
            new_customer = create_customer_name(line)
            if new_customer:
                window['-CUSTOMER-'].update(values=[new_customer])
        elif event == "Move":
            selected_folder = values['-FOLDER LIST-'][0]
            line = values['-LINE-'][0]
            customer = values['-CUSTOMER-'][0]
            if not selected_folder:
                sg.popup_error("Please select a folder to move")
            elif not line:
                sg.popup_error("Please select a line")
            elif not customer:
                sg.popup_error("Please select or create a customer")
            else:
                # Move the folder
                old_folder_path = os.path.join(root_dir, selected_folder)
                new_folder_path = os.path.join(destination_path, line, customer)
                move_folder(old_folder_path, new_folder_path)
                break

    # Close the PySimpleGUI window
    window.close()'''