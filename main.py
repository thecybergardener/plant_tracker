import tkinter as tk
from tkcalendar import Calendar
from datetime import datetime
from plant_info_retriever import PlantInfoRetriever
from plant_info_updater import PlantInfoUpdater
import config

class PlantTrackerApp:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title("Plant Journal")
        self.root.geometry("800x600")
        
        # Initialize Excel interaction classes
        self.plant_collection = PlantInfoRetriever(config.EXCEL_FILE_PATH)
        # self.updater = PlantInfoUpdater(config.EXCEL_FILE_PATH)
        
        # Get the list of plants from Excel
        # self.plant_collection = self.retriever.get_all_plants()
        
        # Create UI
        self.create_widgets()

    def create_widgets(self) -> None:
        # Row 0: Name Frame
        name_frame = tk.Frame(self.root)
        name_frame.grid(row=0, column=0, columnspan=4, pady=10)

        plant_label = tk.Label(name_frame, text="Plant Name:")
        plant_label.grid(row=0, column=0, sticky="W", padx=10)

        # Dropdown for plant name (loaded from Excel sheet)
        self.plant_name = tk.StringVar(self.root)
        plant_names = [plant.plant_name for plant in self.plant_collection]
        plant_menu = tk.OptionMenu(name_frame, self.plant_name, *plant_names, command=self.on_plant_select)
        plant_menu.grid(row=0, column=1, columnspan=2, padx=10, pady=5)

        # Row 1: Detail Frame
        detail_frame = tk.Frame(self.root)
        detail_frame.grid(row=1, column=0, columnspan=2, pady=10)

        self.ownership_data = tk.Label(detail_frame, text="Ownership:")
        self.ownership_data.grid(row=0, column=1, sticky="W", padx=10)

        self.species_data = tk.Label(detail_frame, text="Species:")
        self.species_data.grid(row=1, column=1, sticky="W", padx=10)

        # Date Picker
        date_label = tk.Label(self.root, text="Select Date:")
        date_label.grid(row=4, column=0, pady=10)

        self.cal = Calendar(self.root, selectmode="day", date_pattern="mm/dd/yyyy")
        self.cal.grid(row=4, column=1, pady=10)

        # Submit Button
        submit_button = tk.Button(self.root, text="Submit", width=20, command=self.submit)
        submit_button.grid(row=5, column=0, columnspan=4, pady=20)

    def on_plant_select(self, event):
        selected_plant_name = self.plant_name.get()
        plant = self.plant_collection.get_plant_by_name(selected_plant_name)
        
        if plant:
            self.ownership_data.config(text=f"Ownership: {self.plant_collection.calculate_ownership(plant.purchase_date)}")
            self.species_data.config(text=f"Species: {plant.species}")

    def submit(self):
        plant_name = self.plant_name.get()
        selected_date = self.cal.get_date()
        # Assume some watering and other values (to be expanded based on input)
        watering = "Y"
        humidity = 45
        temperature = 70
        notes = "Looks healthy."

        # Update Excel
        self.updater.update_plant_info(plant_name, datetime.strptime(selected_date, "%m/%d/%Y"), watering, humidity, temperature, notes)
        print(f"Updated plant {plant_name} with data on {selected_date}.")

def main():
    # Instantiate the PlantInfoRetriever and get the plant data
    plants = PlantInfoRetriever(config.EXCEL_FILE_PATH)
    # plants = retriever.get_all_plants()

    # Display the information
    print("Displaying all plants in the collection:")
    print(plants)

    # __repr__
    print(repr(plants))

    # __getitem__
    print(plants[4])

    # __contains__
    print("Snake Plant" in plants)

    # __eq__
    print(plants[1] == plants[4])

    # __call__
    print(plants[1]())

    # description
    print(plants[5].description)

if __name__ == "__main__":
    main()
    root = tk.Tk()
    app = PlantTrackerApp(root)
    root.mainloop()
