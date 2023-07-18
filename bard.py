import pandas as pd

# Import the two tables
table1 = pd.read_csv("Export_Livraisons - 20230713.csv", skiprows=[187], error_bad_lines=False, delimiter=",")
table2 = pd.read_excel("Export vidanges Descartes 01-06 au 13-07.xlsx")

# Compare the two tables
diff = table1.compare(table2, indicator='difference')

# Print the differences
for difference in diff:
  print(difference)
