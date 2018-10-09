# import required libraries
import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename

# import the HMID Shelter Data Report and convert it to a pandas data frame
data = pd.read_excel(askopenfilename())

# drop rows where the housing move in data field is nan
cleaned = data.dropna(subset=["Housing Move-in Date(9160)"])

# create a Days Between Exit and HMID ABS column
cleaned["Days Between Exit and HMID"] = (
    cleaned["Entry Exit Exit Date"] - cleaned["Housing Move-in Date(9160)"]
).dt.days

# create a column reflecting the absolute value of the Days Between Exit and
# HMID column values
cleaned["Days Between Exit and HMID ABS"] = cleaned["Days Between Exit and HMID"].abs()

# sort the cleaned dataframe and drop duplicate rows
likely = cleaned.sort_values(
    by=["Client Uid", "Days Between Exit and HMID ABS"]
).drop_duplicates(subset="Client Uid", keep="first").drop(
    "Days Between Exit and HMID ABS",
    axis=1, inplace=True
)

# flag an exit error when the the days between exit and hmid are less than 32 or
# greater than -15
likely_hmid_exit_error = likely[
    (likely["Days Between Exit and HMID"] < 32) &
    (likely["Days Between Exit and HMID"] > -15)
]
hmid_exit_error = likely_hmid_exit_error[["Client Uid"]]
hmid_exit_error["Exit Date HMID Mismatch"] = "Yes"

# create a Days Between Entry and HMID column
cleaned["Days Between Entry and HMID"] = (
    cleaned["Entry Exit Entry Date"] - cleaned["Housing Move-in Date(9160)"]
).dt.days

#
hmid_e_errors = cleaned.sort_values(
    by=["Client Uid", "Days Between Entry and HMID"]
).drop_duplicates(
    subset=["Client Uid"],
    keep="first"
).drop(
    ["Days Between Exit and HMID ABS", "Days Between Exit and HMID"],
    axis=1
)

# flag an entry error when the days between entry minus the hmid is greater than
# zero
pt_w_e_hmid_error = hmid_e_errors[
    hmid_e_errors["Days Between Entry and HMID"] > 0
][["Client Uid"]]
pt_w_e_hmid_error["Error HMID at Entry"] = "Yes"

# merge the two dataframes
merged = data.dropna(
    subset=["Housing Move-in Date(9160)"]
).merge(
    hmid_exit_error, how="left",
    on="Client Uid"
).merge(pt_w_e_hmid_error, how="left", on="Client Uid").fillna("")

# create a pivot tables for the entry and exit errors that will then be merged
entry_pivot = pd.pivot_table(
    merged[merged["Error HMID at Entry"] == "Yes"],
    index="Entry Exit Provider Id",
    columns="Error HMID at Entry",
    values="Client Uid",
    aggfunc=len
)
exit_pivot = pd.pivot_table(
    merged[merged["Exit Date HMID Mismatch"] == "Yes"],
    index="Entry Exit Provider Id",
    columns="Exit Date HMID Mismatch",
    values="Client Uid",
    aggfunc=len
)
merged_pivots = entry_pivot.merge(
    exit_pivot,
    on="Entry Exit Provider Id",
    how="outer"
)

# add a total errors column
merged_pivots["Total Errors"] = sum(merged_pivots)

# write the data and final pivot to a spreadsheet
writer.ExcelWriter(asksaveasfilename, engine="xlsxwriter")
merged_pivots.to_excel(writer, sheet_name="Summary")
merged.to_excel(writer, sheet_name="Data", index=False)
writer.save()
