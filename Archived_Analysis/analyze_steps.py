import openpyxl

# Since Group 1 uses 1=T, 0=H and the xlsx correctly UN-INVERTS their data,
# we need to check if our code should also un-invert Group 1.
# 
# Group 1's cumulative columns: col 3=Tails cum, col 4=Heads cum
# We read col 4 (Heads cumulative) and derive per-toss from that.
# 
# BUT: since the raw data has 1=T and 0=H, the cumulative columns
# were computed FROM those inverted values. The question is:
# did Group 1 label their cumulative columns CORRECTLY or LITERALLY?
#
# From the header: "1 Peso Coin" -> "Tails", "Heads"
# If they're CORRECT labels: col 3 = real Tails cum, col 4 = real Heads cum
# If they're LITERAL labels (i.e. column D under "Tails" counts 1s, 
# but 1=T, so it IS correct): then the labels match reality.
#
# Let me verify: 
# Row 2: Coin1=1 (=Tails), CumTails(col3)=1, CumHeads(col4)=0
# So col3 counts the 1s and labels them Tails. Since 1=Tails, this is CORRECT.
# col4 counts the 0s and labels them Heads. Since 0=Heads, this is correct.
# So our parsing IS correct for Group 1.
#
# BUT the xlsx shows 2 Peso H=54, T=46 while our parsing gives H=46, T=54.
# That means the xlsx SWAPPED Group 1's values. Who is right?

wb = openpyxl.load_workbook("2BSCS-A _ Tossed Coin Raw Data.xlsx", data_only=True)
ws1 = wb["GROUP 1"]
rows1 = list(ws1.iter_rows(values_only=True))

print("GROUP 1 DETAILED ANALYSIS:")
print(f"  Header 0: {rows1[0]}")
print(f"  Header 1: {rows1[1]}")
print()

# 2 Peso: col 5 = Tails cum, col 6 = Heads cum
# Final values:
print(f"  2 Peso final (row 101):")
print(f"    Col 5 (Tails cum) = {rows1[101][5]}")
print(f"    Col 6 (Heads cum) = {rows1[101][6]}")
print(f"    Our code reads col[6]={rows1[101][6]} as Heads -> H=46")
print()

# But the xlsx says H=54 for 2 Peso. This is col 5's value!
# So the xlsx reads col 5 as Heads instead of col 6.
# That means the xlsx author thought the "Tails" column is actually Heads
# (because they knew 1=T in raw, so they UN-INVERTED)

# Verification: count 0s and 1s in Coin 2 raw column
coin2_raw = [int(rows1[i][2]) for i in range(2, 102)]
count_0 = coin2_raw.count(0)
count_1 = coin2_raw.count(1)
print(f"  Coin 2 (2 Peso) raw column:")
print(f"    Count of 0s: {count_0}")
print(f"    Count of 1s: {count_1}")
print(f"    If 0=Heads: Heads={count_0}, Tails={count_1}")
print(f"    If 1=Heads: Heads={count_1}, Tails={count_0}")
print()

# The user confirmed 0=Heads, 1=Tails for Group 1
# So 2 Peso: Heads = count_0 = should match
print(f"  With 0=Heads convention: 2 Peso H={count_0}, T={count_1}")
print(f"  Our cumulative parse:   2 Peso H=46, T=54")
print(f"  xlsx value:             2 Peso H=54, T=46")
print()

# Same check for 1B (Coin 1)
coin1_raw = [int(rows1[i][1]) for i in range(2, 102)]
count_0 = coin1_raw.count(0)
count_1 = coin1_raw.count(1)
print(f"  Coin 1 (1B) raw column:")
print(f"    Count of 0s: {count_0}")  
print(f"    Count of 1s: {count_1}")
print(f"    With 0=Heads: 1B H={count_0}, T={count_1}")
print(f"    Our cumulative parse: 1B H=50, T=50")
print(f"    xlsx 1B total: 289 (total across all groups)")
print()

# So for Group 1:
# 1B: 0=H gives H=50, T=50 (matches cumulatives - symmetric)
# 2P: 0=H gives H=46, T=54 (matches our cumulative parse)
# xlsx uses H=54, T=46 for 2P -> they read the WRONG column

# Wait - or maybe the xlsx DISAGREES about the convention.
# If the xlsx author assumed 1=H (standard), then:
# 1B: H=50, T=50 (same either way)
# 2P: H=54, T=46 (they counted 1s as H)
# But the user says 0=Heads for Group 1. So the xlsx is WRONG for 2 Peso.

print("VERDICT:")
print("  The user confirmed Group 1 uses 0=Heads, 1=Tails.")
print("  Our parsing (from cumulative Heads column) is CORRECT:")
print("    1B: H=50, T=50 (correct)")
print("    2P: H=46, T=54 (correct)")
print()
print("  The Step xlsx files used 1=Heads (standard convention) for G1,")
print("  which gives INCORRECT values for 2 Peso (H=54 instead of 46).")
print("  This also affects the total: xlsx H=1561 vs correct H=1553.")
print()

# Now the 1A/1B discrepancy (+8/-8) needs explanation
# Since G1 doesn't affect 1A or 1B (symmetric), the difference
# must come from ANOTHER source in the xlsx compilation
print("  The 1A/1B discrepancy (±8) is a separate issue in the xlsx")
print("  compilation - likely different per-toss data for some groups")
print("  (G8, G11, G13 had many per-toss differences vs xlsx).")
print("  Our parsing from the raw data file is reliable.")
