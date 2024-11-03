class Plant:
    def __init__(self, plant_name, species, purchase_date, description) -> None:
        self.species = species
        self.plant_name = plant_name
        self.purchase_date = purchase_date
        self.description = description
    
    def __str__(self) -> str:
        return f"{self.plant_name} ({self.species}) - Purchased on {self.purchase_date}"
    
    def __repr__(self) -> str:
        return f"Plant(plant_name='{self.plant_name}', species='{self.species}', purchase_date='{self.purchase_date}')"

    def __call__(self) -> str:
        return f"{self.plant_name} is thriving!"

    # After some time when program is run, suggest when a plant should be watered by using this function
    # if suggested days is lt days since last water do print "watering not suggested in green"
    # def __lt__(self, other):
    #     return self.purchase_date < other.purchase_date
    # 
    # if suggested days is ge days since last watering print, suggested water
    # def __ge__(self, other):
    #     return self.purchase_date < other.purchase_date
    
class PlantCollection:
    def __init__(self, plants=None):
        self.plant_collection = plants if plants is not None else []

    def __getitem__(self, index):
        return self.plant_collection[index]

    def __contains__(self, plant_name):
        return any(plant.plant_name == plant_name for plant in self.plant_collection)

    def __str__(self) -> str:
        return "\n".join([str(plant) for plant in self.plant_collection])
    
    def __repr__(self) -> str:
        return "\n".join([repr(plant) for plant in self.plant_collection])
    
    def get_plant_by_name(self, plant_name):
        return next((plant for plant in self.plant_collection if plant.plant_name == plant_name), None)