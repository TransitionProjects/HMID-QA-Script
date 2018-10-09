# import required libraries
import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename

# initialize a class
class HMIDDQ:
    def __init__(self, file):
        # convert the file parameter to a data frame
        self.data = pd.read_excel(file)

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
        self.entry_error_pivot = self.pivot_errors_at_entry()
        self.exit_error_pivot = self.pivot_errors_at_exit()
        self.entry_error_data = self.errors_at_entry()
        self.exit_error_data = self.errors_at_exit()

    def pivot_errors_at_entry(self):
        # create a local copy of the data frame with rows that have a nan in the
        # hmid column droped
        w_hmid = self.data.dropna(subset=["Housing Move-in Date(9160)"])
        w_hmid["Provider"] = [
            self.shelter_short[provider] for provider in w_hmid["Entry Exit Provider Id"]
        ]

        # create days between entry and HMID column
        w_hmid["Days Between Entry and HMID"] = (
            w_hmid["Entry Exit Entry Date"] - w_hmid["Housing Move-in Date(9160)"]
        ).dt.days

        # sort by the days between entry and HMID then drop rows retaining then
        # rows with the lower values and duplicate Client Uid
        smallest_delta = w_hmid.sort_values(
            by=["Client Uid", "Days Between Entry and HMID"],
            ascending=True
        ).drop_duplicates(
            subset=["Client Uid"],
            keep="first"
        )

        # create a groupby object showing counts of participants with hmid
        # errors at entry
        pt_with_entry_error = smallest_delta[
            smallest_delta["Days Between Entry and HMID"] > 0
        ][["Client Uid", "Provider"]].rename(
            {"Client Uid": "Count of Participants with HMID Error at Entry"},
            axis=1
        ).groupby(
            by="Provider",
            as_index=False
        ).count()

        # return the groupby object with the Client Uid renamed to make it
        # more meaningfull
        return pt_with_entry_error


    def errors_at_entry(self):
        # create a local copy of the data frame with rows that have a nan in the
        # hmid column droped
        w_hmid = self.data.dropna(subset=["Housing Move-in Date(9160)"])
        w_hmid["Provider"] = [
            self.shelter_short[provider] for provider in w_hmid["Entry Exit Provider Id"]
        ]

        # create days between entry and HMID column
        w_hmid["Days Between Entry and HMID"] = (
            w_hmid["Entry Exit Entry Date"] - w_hmid["Housing Move-in Date(9160)"]
        ).dt.days

        # sort by the days between entry and HMID then drop rows retaining then
        # rows with the lower values and duplicate Client Uid
        smallest_delta = w_hmid.sort_values(
            by=["Client Uid", "Days Between Entry and HMID"],
            ascending=True
        ).drop_duplicates(
            subset=["Client Uid"],
            keep="first"
        )

        # create a groupby object showing counts of participants with hmid
        # errors at entry
        pt_with_entry_error = smallest_delta[
            smallest_delta["Days Between Entry and HMID"] > 0
        ]

        # create a column to help end users identify participants with this error
        pt_with_entry_error["HMID Error at Entry"] = "Yes"

        # return the groupby object with the Client Uid renamed to make it
        # more meaningfull
        return pt_with_entry_error[[
            "Client Uid",
            "Provider",
            "Entry Exit Entry Date",
            "HMID Error at Entry"
        ]]


    def pivot_errors_at_exit(self):
        # create a local copy of the data frame with rows that have a nan in the
        # hmid or entry exit exit date column droped
        w_hmid = self.data.dropna(
            subset=["Housing Move-in Date(9160)", "Entry Exit Exit Date"]
        )
        w_hmid["Provider"] = [
            self.shelter_short[provider] for provider in w_hmid["Entry Exit Provider Id"]
        ]

        # create days between exit and HMID column
        w_hmid["Days Between Exit and HMID"] = (
            w_hmid["Entry Exit Exit Date"] - w_hmid["Housing Move-in Date(9160)"]
        ).dt.days

        # create a column showing the absolute value of the Days Between Exit
        # and HMID columns
        w_hmid["Days Between Exit and HMID ABS"] = w_hmid["Days Between Exit and HMID"].abs()

        # sort by the days between exit and HMID abs then drop rows, retaining
        # those rows with the lower absolute values when they have duplicate
        # Client Uid values
        smallest_delta = w_hmid.sort_values(
            by=["Client Uid", "Days Between Exit and HMID ABS"],
            ascending=True
        ).drop_duplicates(
            subset=["Client Uid"],
            keep="first"
        )

        # create a groupby object showing counts of participants with hmid
        # errors at entry
        pt_with_exit_error = smallest_delta[
            (smallest_delta["Days Between Exit and HMID"] < 32) &
            (smallest_delta["Days Between Exit and HMID"] > -15)
        ][["Client Uid", "Provider"]].rename(
            {"Client Uid": "Count of Participants with HMID Error at Exit"},
            axis=1
        ).groupby(
            by="Provider",
            as_index=False
        ).count()

        # return the groupby object with the Client Uid renamed to make it
        # more meaningfull
        return pt_with_exit_error


    def errors_at_exit(self):
        # create a local copy of the data frame with rows that have a nan in the
        # hmid or entry exit exit date column droped
        w_hmid = self.data.dropna(
            subset=["Housing Move-in Date(9160)", "Entry Exit Exit Date"]
        )
        w_hmid["Provider"] = [
            self.shelter_short[provider] for provider in w_hmid["Entry Exit Provider Id"]
        ]

        # create days between exit and HMID column
        w_hmid["Days Between Exit and HMID"] = (
            w_hmid["Entry Exit Exit Date"] - w_hmid["Housing Move-in Date(9160)"]
        ).dt.days

        # create a column showing the absolute value of the Days Between Exit
        # and HMID columns
        w_hmid["Days Between Exit and HMID ABS"] = w_hmid["Days Between Exit and HMID"].abs()

        # sort by the days between exit and HMID abs then drop rows, retaining
        # those rows with the lower absolute values when they have duplicate
        # Client Uid values
        smallest_delta = w_hmid.sort_values(
            by=["Client Uid", "Days Between Exit and HMID ABS"],
            ascending=True
        ).drop_duplicates(
            subset=["Client Uid"],
            keep="first"
        )

        # create a groupby object showing counts of participants with hmid
        # errors at entry
        pt_with_exit_error = smallest_delta[
            (smallest_delta["Days Between Exit and HMID"] < 32) &
            (smallest_delta["Days Between Exit and HMID"] > -15)
        ]

        # create a column to help end users identify participants with this error
        pt_with_exit_error["HMID Error at Exit"] = "Yes"

        # return the groupby object with the Client Uid renamed to make it
        # more meaningful
        return pt_with_exit_error[[
            "Client Uid",
            "Provider",
            "Entry Exit Entry Date",
            "HMID Error at Exit"
        ]]


    def id_errors_at_placement(self):
        pass

    def id_errors_at_follow_up(self):
        pass

    def process(self):
        # merge the self.entry_error_pivot object and the self.exit_error_pivot
        # object
        merged_pivots = self.entry_error_pivot.merge(
            self.exit_error_pivot,
            how="outer"
        )

        # merge the processed data frames which will be used by the end user
        # to identify which participant files need correction
        merged_data = self.data.merge(
            self.exit_error_data,
            on=["Client Uid", "Entry Exit Entry Date"],
            how="left"
        ).merge(
            self.entry_error_data,
            on=["Client Uid", "Entry Exit Entry Date"],
            how="left"
        ).drop(["Provider_x", "Provider_y"], axis=1)


        # initialize the writer object
        writer = pd.ExcelWriter(
            asksaveasfilename(title="Save the QA HMID Report"),
            engine="xlsxwriter"
        )

        # write the pivot tables to excel
        merged_pivots.to_excel(writer, sheet_name="Summary", index=False)
        merged_data.to_excel(writer, sheet_name="Processed Data", index=False)
        self.data.to_excel(writer, sheet_name="Raw Data", index=False)
        writer.save()

if __name__ == "__main__":
    run = HMIDDQ(askopenfilename(title="Open the QA Shelter Exit HMID Mismatch Report"))
    run.process()
