#!/usr/bin/env python
import pandas as pd
import datetime
import os

file_name = input("Drag and drop file\n").replace('"', "")

df = pd.read_csv(file_name, parse_dates=["Start Date", "Renewal Date", "End Date"])

todaysDate = datetime.datetime.now()

# Only look at records with same month as 2 months previous
df = df[pd.DatetimeIndex(df["Start Date"]).month == todaysDate.month - 2]

# Remove anyone who has an End Date of previous year
df = df[pd.DatetimeIndex(df["End Date"]).year >= todaysDate.year]

# Remove Timezone indicator from date columns
df["Start Date"] = df["Start Date"].dt.tz_localize(None)
df["Renewal Date"] = df["Renewal Date"].dt.tz_localize(None)
df["End Date"] = df["End Date"].dt.tz_localize(None)

# Change status to active or cancelled
df_grouped = df.copy()
df_grouped.loc[df["Status"] != "Active", "Status"] = "Cancelled"

# Get counts of Active/Cancelled
df_grouped = df_grouped.groupby("Status").size()

with pd.ExcelWriter(f"{file_name.split('.')[0]}_Results.xlsx", mode="w") as writer:
    df[
        [
            "Email",
            "FirstName",
            "LastName",
            "Status",
            "Start Date",
            "Renewal Date",
            "End Date",
        ]
    ].to_excel(writer, sheet_name="Monthly Members", index=False)
    df_grouped.to_excel(writer, sheet_name="Counts")
    df.loc[df["Status"] != "Active"][
        ["Email", "FirstName", "LastName", "Status"]
    ].to_excel(writer, sheet_name="Cancelled Memberships", index=False)
