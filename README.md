---
# Plant Tracker App

A Python-based application for managing and tracking plant data. Think of it as a "Plant Journal." I use `tkinter` for the user interface and `openpyxl` for interacting with Excel files. This app helps you maintain detailed information about your plant collection, including names, species, purchase dates, and descriptions.

## Features

- **Initial Setup**: When the app runs for the first time, it prompts the user to select a location for storing the `plant_data` file and allows naming it (default: `plant_data.xlsx`). The app creates the file if it doesn't exist and saves the path in a configuration file.
- **Configuration File**: Stores the path to the `plant_data` file for future use, enabling seamless data loading every time the app runs.
- **File Menu Options**:
  - Create a new plant data file.
  - Open an existing plant data file.
  - Add new plant information directly to the "Basic info" sheet of the Excel file.
- **Plant Data Management**: Displays plant information such as name, species, purchase date, and description.
- **Extendable UI**: Built using `tkinter`, allowing for future enhancements to the user interface.

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/cybergardner/plant-tracker.git
   cd plant-tracker
   ```

2. **Install dependencies**:
   Make sure you have Python 3.x installed, then install the required Python packages:
   ```bash
   sudo apt install python3-tk
   python3 -m venv venv
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python main.py
   ```

## Project Structure

```
project_folder/
│
├── main.py                  # Main application script
├── plant.py                 # Defines the Plant class
├── plant_collection.py      # Defines the PlantCollection class
├── plant_info_retriever.py  # Handles data retrieval from Excel
├── plant_info_updater.py    # Handles data updates to Excel
├── config.ini               # Configuration file for storing the path to the plant_data file
└── README.md                # Project documentation
```

## How It Works

### Initial Run
- When the app is launched for the first time, it prompts the user to select a directory and name for the `plant_data` file.
- The selected path is saved in `config.ini` for future reference.
- A new Excel file with a "Basic info" sheet is created if it doesn’t already exist.

### Subsequent Runs
- The app reads the file path from `config.ini` and loads the `plant_data` file automatically.
- Users can create a new file or open an existing one if desired.

### Adding New Plant Data
- The "File" menu provides an option to add new plant entries directly to the "Basic info" sheet of the Excel file.

## Dependencies

- `tkinter`: Built-in Python module for creating graphical user interfaces.
- `openpyxl`: Library for reading and writing Excel (XLSX) files.
- `tkcalendar`: Calendar widget for `tkinter`.

## Configuration File

**`config.ini`**:
- This file is used to store the path to the `plant_data` file.
- It is automatically created and updated by the app.

## Future Enhancements

- Add more comprehensive plant management features, such as tracking watering schedules and plant health.
- Implement advanced data visualization for plant growth and environmental conditions.
- Integrate machine learning algorithms for predicting plant care needs.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.

## Contributing

Contributions are not welcome at this time.

## Contact

For any inquiries or issues, please contact Bola at contant@cybergardner.cc

