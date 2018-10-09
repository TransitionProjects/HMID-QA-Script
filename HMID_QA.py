# import required libraries
import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename

# import the HMID Shelter Data Report and convert it to a pandas data frame
data = pd.read_excel(askopenfilename())

#
cleaned = data.dropna(subset=["Housing Move-in Date(9160)"])

#
cleaned["Days Between Exit and HMID"] = (
    cleaned["Entry Exit Exit Date"] - cleaned["Housing Move-in Date(9160)"]
).dt.days

#
cleaned["Days Between Exit and HMID ABS"] = cleaned["Days Between Exit and HMID"].abs()

#
likely = cleaned.sort_values(
    by=["Client Uid", "Days Between Exit and HMID ABS"]
).drop_duplicates(subset="Client Uid", keep="first").drop(
    "Days Between Exit and HMID ABS",
    axis=1, inplace=True
)

#
likely_hmid_exit_error = likely[
    (likely["Days Between Exit and HMID"] < 32) &
    (likely["Days Between Exit and HMID"] > -15)
]
hmid_exit_error = likely_hmid_exit_error[["Client Uid"]]
hmid_exit_error["Exit Date HMID Mismatch"] = "Yes"

#
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

# 
pt_w_e_hmid_error = hmid_e_errors[
    hmid_e_errors["Days Between Entry and HMID"] > 0
][["Client Uid"]]
pt_w_e_hmid_error["Error HMID at Entry"] = "Yes"

#
merged = data.dropna(
    subset=["Housing Move-in Date(9160)"]
).merge(
    hmid_exit_error, how="left",
    on="Client Uid"
).merge(pt_w_e_hmid_error, how="left", on="Client Uid").fillna("")
