
#### What packages we need to download .....

import os
import sys
import time

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")     # save figures to files (no pop-up windows)
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


# STEP 1: Settings, here we must introduce our project name and study case defined in PF. 

PF_ROOT = r"C:\Program Files\DIgSILENT\PowerFactory 2026 SP3"
PF_PYTHON = PF_ROOT + r"\Python\3.12"           # Be careful:  must match your Python version

PROJECT_NAME = "NordEnergi_grid"  ### Our project named defined in PF digsilent earlier (Data manager). 
STUDY_CASE = "Case0"   ### Our study case name. 

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_EXCEL = os.path.join(SCRIPT_DIR, "data", "INPUT EXCEL FILE NAME")  ### Excel file that we use to update power consumption of each load.
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "CASE[INSERT]RESULTS")  ### Excel file that we defined to save all the outputs we receive from PF Digsilent. 

# Operating limits used to highlight problems in plots and Excel (VERIFY UPON TESTING START)
V_MIN, V_MAX = 0.95, 1.05               # pu
LOADING_MAX = 100.0                     # %

# STEP 2: Connect Python to PF Digsilent

# Python must know where the "powerfactory" module is, then add PF's Python folder.
if os.path.isdir(PF_PYTHON):
    sys.path.insert(0, PF_PYTHON)
    os.environ["PATH"] = PF_ROOT + ";" + PF_PYTHON + ";" + os.environ["PATH"]
import powerfactory as pf

app = pf.GetApplication()               # starts PowerFactory in engine mode
if app is None:
    raise RuntimeError("Could not start PowerFactory (is the GUI still open?).")
app.ClearOutputWindow()
print("STEP 2  Connected to PowerFactory.")


# STEP 3: Activate the project and the use case

app.ActivateProject(PROJECT_NAME)
if app.GetActiveProject() is None:
    raise RuntimeError(f"Project '{PROJECT_NAME}' not found.")
study_case = app.GetProjectFolder("study").GetContents(STUDY_CASE)[0]
study_case.Activate()
print(f"STEP 3  Project '{PROJECT_NAME}' and study case '{study_case.loc_name}' activated.")

# STEP 4: Read the excel file with the load profiles

# other column headers = EXACT names of the loads in PowerFactory): ### Be careful !!!! ####

# parse_dates=True: turn those time texts into real date/time values
sheets = pd.read_excel(INPUT_EXCEL, sheet_name=None, index_col=0, parse_dates=True)
loadinput = sheets["Aug_26_consumption_all"]                 # rows = time steps, columns = loads
time_steps = loadinput.index
step_minutes = (time_steps[1] - time_steps[0]).total_seconds() / 60 if len(time_steps) > 1 else 0
print(f"STEP 4  Read {len(time_steps)} time steps of {step_minutes:g} min "
      f"for loads {list(loadinput.columns)}")

# STEP 5: Find the grid elements in PF Digsilent 

# GetCalcRelevantObjects("*.ElmXxx") returns all objects of a class in the active grid.
#   ElmLod = load   ElmLne = line   ElmTr2 = 2-winding transformer
#   ElmTerm = bus/terminal          ElmSym = synchronous generator
def get_sorted(pattern):
    return sorted(app.GetCalcRelevantObjects(pattern), key=lambda o: o.loc_name)

all_loads = {o.loc_name: o for o in get_sorted("*.ElmLod")}
all_gens = {o.loc_name: o for o in get_sorted("*.ElmSym")}
lines = get_sorted("*.ElmLne")
trafos = get_sorted("*.ElmTr2")
buses = get_sorted("*.ElmTerm")

# Each Excel column must match a load name in PowerFactory,  check again it now.
missing = [name for name in loadinput.columns if name not in all_loads]
if missing:
    raise KeyError(f"Excel columns {missing} are not loads in PowerFactory. "
                   f"Loads in the model: {list(all_loads)}")
loads = {name: all_loads[name] for name in loadinput.columns}
print(f"STEP 5  Found {len(loads)} loads"
      f"{len(lines)} lines, {len(trafos)} transformers, {len(buses)} buses.")

# 
# STEP 6: Remember the original values 

# Writing from Python changes the project itself, so we restore it at the end.
original_loads = {name: (obj.plini) for name, obj in loads.items()}
print("STEP 6  Original load values stored.")


# STEP 7: The time series loop 

ldf = app.GetFromStudyCase("ComLdf")      # the load-flow command of the study case
ldf.iopt_net = 0                          # 0 = balanced AC load flow 1 = unbalanced AC 2 = DC

line_loading, trafo_loading, bus_voltage, status = [], [], [], []
app.EchoOff()                             # no messages in the output window = faster
t_start = time.perf_counter()

try:
    for k, t in enumerate(time_steps):

        # 7a  UPDATE: write this time step's values into PowerFactory 
        for name, load in loads.items():
            load.plini = float(loadinput.loc[t, name])        # active power  [MW]

        # 7b  SOLVE: run the load flow (Execute() returns 0 if converged) -
        converged = ldf.Execute() == 0

        # 7c  COLLECT: read the results of this time step 
        #   c:loading = loading in %       m:u = voltage magnitude in pu
        if converged:
            line_loading.append({ln.loc_name: ln.GetAttribute("c:loading") for ln in lines})
            trafo_loading.append({tr.loc_name: tr.GetAttribute("c:loading") for tr in trafos})
            bus_voltage.append({b.loc_name: b.GetAttribute("m:u") for b in buses})
        else:                                               # keep the row, mark it empty
            line_loading.append({ln.loc_name: np.nan for ln in lines})
            trafo_loading.append({tr.loc_name: np.nan for tr in trafos})
            bus_voltage.append({b.loc_name: np.nan for b in buses})
        status.append({"Converged": converged, "Total_Load_MW": loadinput.loc[t].sum()})

        # progress message every ~10 % of the steps
        if k % max(1, len(time_steps) // 10) == 0 or not converged:
            print(f"STEP 7  {k + 1:5d}/{len(time_steps)}  {t}  "
                  f"{'converged' if converged else 'NOT CONVERGED'}  "
                  f"max line loading = {max(line_loading[-1].values()):5.1f} %  "
                  f"min voltage = {min(bus_voltage[-1].values()):.4f} pu")

finally:
    

    # STEP 8: Restore the original value  
    # # Switch PowerFactory's messages back on.
    app.EchoOn()
    for name, (p0) in original_loads.items():
        loads[name].plini = p0
    print("STEP 8  Original values restored in PowerFactory.")

runtime = time.perf_counter() - t_start
print(f"        {len(time_steps)} load flows in {runtime:.1f} s "
      f"({1000 * runtime / len(time_steps):.0f} ms per time step)")


# STEP 9: Save all results to one excel file 
# Lists of dictionaries:  tables with rows = time steps, columns = elements

line_df = pd.DataFrame(line_loading, index=time_steps)
trafo_df = pd.DataFrame(trafo_loading, index=time_steps)
volt_df = pd.DataFrame(bus_voltage, index=time_steps)

summary = pd.DataFrame(status, index=time_steps)
summary["Max_Line_Loading_pct"] = line_df.max(axis=1)
summary["Most_Loaded_Line"] = line_df.idxmax(axis=1)
summary["Max_Trafo_Loading_pct"] = trafo_df.max(axis=1)
summary["V_min_pu"] = volt_df.min(axis=1)
summary["V_min_Bus"] = volt_df.idxmin(axis=1)
summary["V_max_pu"] = volt_df.max(axis=1)
summary["V_max_Bus"] = volt_df.idxmax(axis=1)
summary["Problem"] = ((summary["Max_Line_Loading_pct"] > LOADING_MAX)
                      | (summary["Max_Trafo_Loading_pct"] > LOADING_MAX)
                      | (summary["V_min_pu"] < V_MIN) | (summary["V_max_pu"] > V_MAX))

statistics = pd.concat({
    "Line loading (%)": line_df.agg(["min", "mean", "max"]).T,
    "Trafo loading (%)": trafo_df.agg(["min", "mean", "max"]).T,
    "Voltage (pu)": volt_df.agg(["min", "mean", "max"]).T})

os.makedirs(OUTPUT_DIR, exist_ok=True)
input_name = os.path.splitext(os.path.basename(INPUT_EXCEL))[0]
out_excel = os.path.join(OUTPUT_DIR, f"results_{input_name}.xlsx")
with pd.ExcelWriter(out_excel, engine="openpyxl") as xl:
    summary.to_excel(xl, sheet_name="Summary")
    line_df.to_excel(xl, sheet_name="Line_Loading_pct")
    trafo_df.to_excel(xl, sheet_name="Trafo_Loading_pct")
    volt_df.to_excel(xl, sheet_name="Bus_Voltage_pu")
    statistics.to_excel(xl, sheet_name="Statistics")
    loadinput.to_excel(xl, sheet_name="Input_Load_P_MW")
print(f"STEP 9  Results saved: {out_excel}")


# STEP 10: Plot the results, save them, and add them to the excel file 

COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948",
          "#52514e", "#8b5a2b"]                       # one fixed colour per element
LIMIT_COLOR = "#d03b3b"
plt.rcParams.update({"axes.spines.top": False, "axes.spines.right": False, "axes.grid": True,
                     "grid.color": "#e1e0d9", "font.size": 9.5, "axes.titleweight": "bold",
                     "axes.titlelocation": "left", "legend.frameon": False})


def time_axis(ax):
    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))


def plot_lines(ax, df, title, ylabel, limits=(), legend_outside=False):
    """One line per element + dashed red limit lines."""
    for i, col in enumerate(df.columns):
        ax.plot(df.index, df[col], color=COLORS[i % len(COLORS)], lw=1.6, label=col)
    for y in limits:
        ax.axhline(y, color=LIMIT_COLOR, ls="--", lw=1.1)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    if legend_outside:                                 # many elements -> legend to the right
        ax.legend(loc="upper left", bbox_to_anchor=(1.01, 1.0), fontsize=8)
    else:
        ax.legend(ncol=min(len(df.columns), 5), loc="upper left", fontsize=8)
    time_axis(ax)


figures = []

# Figure 1: inputs (what we wrote into PowerFactory)
fig, ax = plt.subplots(figsize=(11, 3.6))
plot_lines(ax, loadinput, "Input: active power of each load (from Excel)", "MW")
figures.append(("01_input_loads.png", fig))

# Figure 2: line loading
fig, ax = plt.subplots(figsize=(11, 4.2))
plot_lines(ax, line_df, f"Line loading per time step (red dashed = {LOADING_MAX:g} % limit)",
           "Loading (%)", limits=[LOADING_MAX])
ax.set_ylim(0, 1.25 * max(LOADING_MAX, np.nanmax(line_df.values)))
figures.append(("02_line_loading.png", fig))

# Figure 3: transformer loading
fig, ax = plt.subplots(figsize=(11, 3.6))
plot_lines(ax, trafo_df, "Transformer loading per time step", "Loading (%)", limits=[LOADING_MAX])
ax.set_ylim(0, 1.25 * max(LOADING_MAX, np.nanmax(trafo_df.values)))
figures.append(("03_trafo_loading.png", fig))

# Figure 4: bus voltages (curves + heat map)
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 7), sharex=True)
plot_lines(ax1, volt_df, f"Bus voltage magnitude (red dashed = {V_MIN}-{V_MAX} pu band)",
           "Voltage (pu)", limits=[V_MIN, V_MAX], legend_outside=True)
t_num = mdates.date2num(volt_df.index.to_pydatetime())
dt = t_num[1] - t_num[0] if len(t_num) > 1 else 1 / 24
im = ax2.imshow(volt_df.T.values, aspect="auto", cmap="RdBu_r", vmin=V_MIN, vmax=V_MAX,
                extent=[t_num[0], t_num[-1] + dt, len(buses) - 0.5, -0.5], interpolation="nearest")
ax2.set_yticks(range(len(volt_df.columns)))
ax2.set_yticklabels(volt_df.columns)
ax2.grid(False)
ax2.set_title("Same voltages as a heat map (dark red / dark blue = at the limit)")
ax2.xaxis_date()
time_axis(ax2)
cax = ax2.inset_axes([1.01, 0.0, 0.015, 1.0])        # colour bar beside the heat map only
fig.colorbar(im, cax=cax, label="Voltage (pu)", extend="both")
figures.append(("04_bus_voltages.png", fig))

# Figure 5: overview of the worst values per time step
fig, axes = plt.subplots(3, 1, figsize=(11, 7.5), sharex=True)
axes[0].plot(summary.index, summary["Total_Load_MW"], color=COLORS[0])
axes[0].set_title("Total load"); axes[0].set_ylabel("MW")
axes[1].plot(summary.index, summary["Max_Line_Loading_pct"], color=COLORS[1], label="highest line")
axes[1].plot(summary.index, summary["Max_Trafo_Loading_pct"], color=COLORS[2], label="highest transformer")
axes[1].axhline(LOADING_MAX, color=LIMIT_COLOR, ls="--", lw=1.1)
axes[1].set_title("Highest loading in the grid"); axes[1].set_ylabel("%"); axes[1].legend(loc="upper left")
axes[2].plot(summary.index, summary["V_min_pu"], color="#184f95", label="lowest bus voltage")
axes[2].plot(summary.index, summary["V_max_pu"], color="#3987e5", label="highest bus voltage")
for y in (V_MIN, V_MAX):
    axes[2].axhline(y, color=LIMIT_COLOR, ls="--", lw=1.1)
axes[2].set_title("Voltage extremes"); axes[2].set_ylabel("pu"); axes[2].legend(loc="center left")
for ax in axes:
    time_axis(ax)
figures.append(("05_overview.png", fig))

# Save every figure as PNG
for file_name, fig in figures:
    fig.tight_layout()
    fig.savefig(os.path.join(OUTPUT_DIR, file_name), dpi=130, bbox_inches="tight")
    plt.close(fig)
print(f"STEP 10 {len(figures)} figures saved in {OUTPUT_DIR}")

# Put the figures into an extra "Figures" sheet of the results workbook
try:
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    wb = load_workbook(out_excel)
    ws = wb.create_sheet("Figures", 0)                 # first sheet of the workbook
    row = 1
    for file_name, _ in figures:
        img = XLImage(os.path.join(OUTPUT_DIR, file_name))
        img.width, img.height = img.width * 0.6, img.height * 0.6
        ws.add_image(img, f"A{row}")
        row += int(img.height / 20) + 2                # ~20 px per Excel row
    wb.save(out_excel)
    print("        Figures also inserted into the 'Figures' sheet of the Excel file.")
except ImportError:
    print("        (install 'pillow' to also insert the figures into the Excel file)")

# Short report for the class
print("\nRESULT HIGHLIGHTS")
print(f"  highest line loading : {summary['Max_Line_Loading_pct'].max():6.1f} % "
      f"({summary.loc[summary['Max_Line_Loading_pct'].idxmax(), 'Most_Loaded_Line']}, "
      f"{summary['Max_Line_Loading_pct'].idxmax()})")
print(f"  highest trafo loading: {summary['Max_Trafo_Loading_pct'].max():6.1f} %")
print(f"  lowest voltage       : {summary['V_min_pu'].min():6.4f} pu "
      f"({summary.loc[summary['V_min_pu'].idxmin(), 'V_min_Bus']}, {summary['V_min_pu'].idxmin()})")
print(f"  highest voltage      : {summary['V_max_pu'].max():6.4f} pu "
      f"({summary.loc[summary['V_max_pu'].idxmax(), 'V_max_Bus']})")
print(f"  time steps with a limit violation: {int(summary['Problem'].sum())} of {len(summary)}")
print(f"  non-converged load flows         : {int((~summary['Converged']).sum())}")
