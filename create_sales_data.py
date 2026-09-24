import pandas as pd

data = [
    ["2026-01-05", "C001", "Laptop", "Electronics", 2, 650, "John"],
    ["2026-01-08", "C002", "Mouse", "Accessories", 5, 25, "Emma"],
    ["2026-01-12", "C003", "Keyboard", "Accessories", 3, 45, "Michael"],
    ["2026-01-18", "C004", "Monitor", "Electronics", 2, 220, "Sophia"],
    ["2026-01-25", "C005", "Laptop", "Electronics", 1, 720, "Daniel"],

    ["2026-02-03", "C006", "Headphones", "Accessories", 4, 80, "John"],
    ["2026-02-09", "C007", "Laptop", "Electronics", 3, 680, "Emma"],
    ["2026-02-14", "C008", "Mouse", "Accessories", 8, 22, "Michael"],
    ["2026-02-21", "C009", "Monitor", "Electronics", 2, 250, "Sophia"],
    ["2026-02-27", "C010", "Keyboard", "Accessories", 5, 40, "Daniel"],

    ["2026-03-04", "C011", "Laptop", "Electronics", 2, 700, "John"],
    ["2026-03-10", "C012", "Monitor", "Electronics", 3, 230, "Emma"],
    ["2026-03-15", "C013", "Headphones", "Accessories", 6, 75, "Michael"],
    ["2026-03-20", "C014", "Mouse", "Accessories", 10, 20, "Sophia"],
    ["2026-03-28", "C015", "Keyboard", "Accessories", 4, 42, "Daniel"],

    ["2026-04-05", "C016", "Laptop", "Electronics", 2, 690, "John"],
    ["2026-04-11", "C017", "Monitor", "Electronics", 4, 210, "Emma"],
    ["2026-04-17", "C018", "Headphones", "Accessories", 5, 85, "Michael"],
    ["2026-04-23", "C019", "Mouse", "Accessories", 7, 24, "Sophia"],
    ["2026-04-29", "C020", "Keyboard", "Accessories", 6, 43, "Daniel"],

    ["2026-05-03", "C021", "Laptop", "Electronics", 3, 710, "John"],
    ["2026-05-08", "C022", "Monitor", "Electronics", 2, 240, "Emma"],
    ["2026-05-14", "C023", "Headphones", "Accessories", 7, 78, "Michael"],
    ["2026-05-20", "C024", "Mouse", "Accessories", 9, 23, "Sophia"],
    ["2026-05-27", "C025", "Keyboard", "Accessories", 5, 44, "Daniel"],

    ["2026-06-04", "C026", "Laptop", "Electronics", 2, 730, "John"],
    ["2026-06-10", "C027", "Monitor", "Electronics", 3, 225, "Emma"],
    ["2026-06-16", "C028", "Headphones", "Accessories", 4, 82, "Michael"],
    ["2026-06-22", "C029", "Mouse", "Accessories", 8, 21, "Sophia"],
    ["2026-06-28", "C030", "Keyboard", "Accessories", 7, 41, "Daniel"],
]

columns = [
    "Date",
    "Customer ID",
    "Product",
    "Category",
    "Quantity",
    "Unit Price",
    "Salesperson"
]

df = pd.DataFrame(data, columns=columns)

df["Date"] = pd.to_datetime(df["Date"])

df["Total Sales"] = df["Quantity"] * df["Unit Price"]

df.to_excel("sales_data.xlsx", index=False)

print("Sales dataset created successfully!")
print("Total Records:", len(df))
print()
print(df.head())