# import required libraries
import pandas as pd

from numpy import select
from tkinter.filedialog import askopenfilename, asksaveasfilename

# initialize a class
class HMIDDQ:
    def __init__(self, file):
        # convert the file parameter to a data frame
        self.data = pd.read_excel(file)
        self.data["Entry Date"] = self.data["Entry Exit Entry Date"].dt.date
        self.data["Exit Date"] = self.data["Entry Exit Exit Date"].dt.date
        self.data["HMID"] = self.data["Housing Move-in Date(9160)"].dt.date

        # a dictionary of full provider names, specifically shelters, and short
        # names that should be used for them
        self.shelter_short = {
            "Transition Projects (TPI) - VA Grant Per Diem (inc. Doreen's Place GPD) - SP(3189)": "Doreen's Place",
            "Transition Projects (TPI) - Doreen's Place - SP(28)": "Doreen's Place",
            "Transition Projects (TPI) - Clark Center - SP(25)": "Clark Center",
            "Transition Projects (TPI) - Jean's Place L1 - SP(29)": "Jean's Place",
            "Transition Projects (TPI) - Willamette Center(5764)": "Willamette Shelter",
            "Transition Projects (TPI) - WyEast Emergency Shelter(6612)": "Wy'East Shelter",
            "Transition Projects (TPI) - SOS Shelter(2712)": "SOS Shelter",
            "Transition Projects (TPI) - Columbia Shelter(6527)": "Columbia Shelter"
        }

        self.perm_dest = [
            "Owned by client, no ongoing housing subsidy (HUD)",
            "Owned by client, with ongoing housing subsidy (HUD)",
            "Permanent housing for formerly homeless persons (HUD)",
            "Rental by client, no ongoing housing subsidy (HUD)",
            "Rental by client, with other ongoing housing subsidy (HUD)",
            "Rental by client, with VASH subsidy (HUD)",
            "Staying or living with family, permanent tenure (HUD)",
            "Staying or living with friends, permanent tenure (HUD)",
            "Foster care home or foster care group home (HUD)",
            "Rental by client, with GPD TIP subsidy (HUD)",
            "Permanent housing (other than RRH) for formerly homeless persons (HUD)",
            "Moved from one HOPWA funded project to HOPWA PH (HUD)",
            "Long-term care facility or nursing home (HUD)",
            "Residential project or halfway house with no homeless criteria (HUD)"
        ]
        self.temp_dest = [
            "Hospital or other residential non-psychiatric medical facility (HUD)",
            "Hotel or motel paid for without emergency shelter voucher (HUD)",
            "Jail, prison or juvenile detention facility (HUD)",
            "Staying or living with family, temporary tenure (e.g., room, apartment or house)(HUD)",
            "Staying or living with friends, temporary tenure (e.g., room apartment or house)(HUD)",
            "Transitional housing for homeless persons (including homeless youth) (HUD)",
            "Moved from one HOPWA funded project to HOPWA TH (HUD)",
            "Substance abuse treatment facility or detox center (HUD)",
            "Psychiatric hospital or other psychiatric facility (HUD)"
        ]
        self.other_dest = [
            "Deceased (HUD)",
            "Emergency shelter, including hotel or motel paid for with emergency shelter voucher (HUD)",
            "Place not meant for habitation (HUD)",
            "Other (HUD)",
            "No exit interview completed (HUD)",
            "Client doesn't know (HUD)",
            "Client refused (HUD)"
        ]

    def make_errors_df(self):
        # Make a local copy of the dataframe
        df = self.data.drop(
            columns=[
                "Entry Exit Exit Date",
                "Entry Exit Entry Date",
                "Housing Move-in Date(9160)"
            ]
        )

        # Use the select method to create a column showing which HMIDs are
        # erroneous.
        conditions = [
            (df["HMID"] < df["Entry Date"]),
            (df["HMID"] == df["Entry Date"]),
            (
                (df["HMID"] > df["Entry Date"]) &
                ((df["HMID"] < df["Exit Date"]) | (df["HMID"] == df["Exit Date"]))
            ),
            (
                df["Entry Exit Destination"].isin(self.perm_dest) &
                df["HMID"].isna()
            ),
            (df["Entry Exit Destination"] == "Data not collected (HUD)"),
            (
                (
                    df["Entry Exit Destination"].isin(self.temp_dest) |
                    df["Entry Exit Destination"].isin(self.other_dest)
                ) &
                (df["HMID"] > df["Exit Date"])
            )
        ]
        choices = [
            "HMID Prior to Entry Date",
            "HMID Matching Shelter Entry Date",
            "HMID During Shelter Stay",
            "Missing HMID",
            "Invalid Exit Destination Selected",
            "HMID Set When Leaving to Non-Perm Destination"
        ]
        df["HMID Error Type"] = select(conditions, choices, "")

        # Use the select method to create a column showing appropriate
        # corrective actions
        choices_c = [
            "Delete the HMID in the entry assessment",
            "Delete the HMID in the entry assessment",
            "Change the HMID to the day after their shelter exit date",
            "Confirm that pt exited to perm, and either change exit destination or enter HMID in exit as exit date plus one day",
            "Change Exit Destination to a valid option",
            "Confirm Exit Destination and HMID.  Changing either, both, or none of these fields may be valid depending on the situation"
        ]
        df["Suggested Correction"] = select(conditions, choices_c, "")


        return df

    def process(self):
        processed = self.make_errors_df()
        # initialize the writer object
        writer = pd.ExcelWriter(
            asksaveasfilename(title="Save the QA HMID Report"),
            engine="xlsxwriter"
        )

        # write the pivot tables to excel
        processed[~(processed["HMID Error Type"] == "")].to_excel(writer, sheet_name="Processed Data", index=False)
        self.data.to_excel(writer, sheet_name="Raw Data", index=False)
        writer.save()

if __name__ == "__main__":
    run = HMIDDQ(askopenfilename(title="Open the QA Shelter Exit HMID Mismatch Report"))
    run.process()
