# import required libraries
import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename

# initialize a class
class HMIDDQ:
    def __init__(self, file):
        self.data = pd.read_excel(file)
        self.shelter_short = {
            "Transition Projects (TPI) - VA Grant Per Diem (inc. Doreen's Place GPD) - SP(3189)": "Doreen's Place",
            "Transition Projects (TPI) - Doreen's Place - SP(28)": "Doreen's Place",
            "Transition Projects (TPI) - Clark Center - SP(25)": "Clark Center",
            "Transition Projects (TPI) - Jean's Place L1 - SP(29)": "Jean's Place",
            "Transition Projects (TPI) - Willamette Center(5764)": "Willamette Shelter",
            "Transition Projects (TPI) - WyEast Emergency Shelter(6612)": "WyEast Shelter",
            "Transition Projects (TPI) - SOS Shelter(2712)": "SOS Shelter",
            "Transition Projects (TPI) - Columbia Shelter(6527)": "Columbia Shelter"
        }

    def id_errors_at_entry(self):
        pass

    def id_errors_at_exit(self):
        pass

    def id_errors_at_placement(self):
        pass

    def id_errors_at_follow_up(self):
        pass

    def process(self):
        pass

if __name__ == "__main__":
    HMIDQ(askopenfilename("Open the QA Shelter Exit HMID Mismatch Report")).process()
