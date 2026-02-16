import openpyxl
import matplotlib.pyplot as plt
import numpy as np

FILENAME = "2BSCS-A _ Tossed Coin Raw Data.xlsx"
wb = openpyxl.load_workbook(FILENAME, data_only=True)
ws = wb["GROUP 14"]
rows = list(ws.iter_rows(values_only=True))

def safe_int(v, default=0):
    if v is None: return default
    try: return int(v)
    except: return default

# Parse Group 14 Data
tosses_1a = [safe_int(rows[i][1]) for i in range(2, 102)]
tosses_20 = [safe_int(rows[i][3]) for i in range(2, 102)]

BLUE = "#2563EB"
RED = "#DC2626"
DARK_BG = "#1E293B"
CARD_BG = "#334155"
TEXT_CLR = "white"

def cumulative_ht(tosses):
    cum_h, cum_t = [], []
    h, t = 0, 0
    for v in tosses:
        if v == 1: h += 1
        else: t += 1
        cum_h.append(h)
        cum_t.append(t)
    return cum_h, cum_t

# Plot Step 1: Group 14 Coin Class Cumulative
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
fig.patch.set_facecolor(DARK_BG)
fig.suptitle("Step 1 — Group 14 H&T (Coin Class)", fontsize=16, fontweight="bold", color=TEXT_CLR)

coins = [("1A", tosses_1a), ("20 Peso", tosses_20)]

for idx, (name, t_list) in enumerate(coins):
    ax = axes[idx]
    ax.set_facecolor(CARD_BG)
    ch, ct = cumulative_ht(t_list)
    x = np.arange(1, len(t_list)+1)
    
    ax.plot(x, ch, color=BLUE, linewidth=2, label="Heads")
    ax.plot(x, ct, color=RED, linewidth=2, label="Tails")
    
    ax.set_title(f"Coin: {name}", color=TEXT_CLR, fontweight="bold")
    ax.set_xlabel("Toss #", color=TEXT_CLR)
    ax.set_ylabel("Cumulative Count", color=TEXT_CLR)
    ax.tick_params(colors=TEXT_CLR)
    ax.legend()
    ax.grid(alpha=0.15)
    
    h_tot = ch[-1]
    t_tot = ct[-1]
    ax.annotate(f"H={h_tot}", xy=(len(t_list), h_tot), color=BLUE, fontweight="bold", xytext=(5,0), textcoords="offset points")
    ax.annotate(f"T={t_tot}", xy=(len(t_list), t_tot), color=RED, fontweight="bold", xytext=(5,-12), textcoords="offset points")

plt.tight_layout()
plt.savefig("group14_step1.png", facecolor=fig.get_facecolor())
print("Saved group14_step1.png")
