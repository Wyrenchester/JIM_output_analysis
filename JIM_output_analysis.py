# -*- coding: utf-8 -*-
"""
Created on Mon Oct 20 15:15:10 2025

@author: yli355

Code for analysing output from JIM

TO DO:
    Excel file for each category, with tabs per channel instead of per particle
"""
import os
os.environ["OMP_NUM_THREADS"] = "1"
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
import scipy as scipy
from hmmlearn import hmm
import xlsxwriter

# -------------------------------
# GUI TOOLTIP
# -------------------------------
class ToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.text = None

    def showtip(self, text):
        if self.tipwindow or not text:
            return
        self.text = text
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 57
        y += cy + self.widget.winfo_rooty() + 27
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "9", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None

def CreateToolTip(widget, text):
    tooltip = ToolTip(widget)
    widget.bind('<Enter>', lambda e: tooltip.showtip(text))
    widget.bind('<Leave>', lambda e: tooltip.hidetip())

# -------------------------------
# DEFAULT GLOBAL VARIABLES
# -------------------------------
RedFile = ""
GreenFile = ""
Fit_targ = "red"
saveplot = "y"
sec_per_f = 5
out_file_name = ""
tracks = pd.DataFrame()
Align = "n"
Align_channel = "red"
FramesBefore = 5

# -------------------------------
# GUI SETUP
# -------------------------------
root = tk.Tk()
root.title("Tracking Particles GUI")
entry_widgets = {}
current_row = 0

def browse_file(entry_widget):
    filename = filedialog.askopenfilename(
        filetypes=(("Excel Files", ("*.xls", "*.csv", "*.xlsx")), ("All files", "*.*")))
    if filename:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filename)

# --- Red and Green Channel Files ---
for label_text, tip in zip(["Red Channel File:", "Green Channel File:"],
                           ["Select Red channel data", "Select Green channel data"]):
    label = tk.Label(root, text=label_text)
    label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
    CreateToolTip(label, tip)

    entry = tk.Entry(root, width=50)
    entry.grid(row=current_row, column=1, padx=5, pady=5)
    entry_widgets[label_text] = entry

    button = tk.Button(root, text="Browse", command=lambda e=entry: browse_file(e))
    button.grid(row=current_row, column=2, padx=5, pady=5)
    current_row += 1

# --- HMM fit radio buttons ---
fit_label = tk.Label(root, text="HMM Fit Channel:")
fit_label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
CreateToolTip(fit_label, "Select which channel to perform HMM fitting on")
fit_var = tk.StringVar(value="red")
tk.Radiobutton(root, text="Red", variable=fit_var, value="red").grid(row=current_row, column=1, padx=5, pady=5)
tk.Radiobutton(root, text="Green", variable=fit_var, value="green").grid(row=current_row, column=2, padx=5, pady=5)
current_row += 1

# --- Save plot radio buttons ---
save_label = tk.Label(root, text="Save plot:")
save_label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
CreateToolTip(save_label, "Select whether to save plots")
save_var = tk.StringVar(value="y")
tk.Radiobutton(root, text="Yes", variable=save_var, value="y").grid(row=current_row, column=1, padx=5, pady=5)
tk.Radiobutton(root, text="No", variable=save_var, value="n").grid(row=current_row, column=2, padx=5, pady=5)
current_row += 1

# Align traces radio buttons
align_label = tk.Label(root, text="Align traces at 2nd drop:")
align_label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
CreateToolTip(align_label, "Align traces to 2nd stage change of HMM Fit Channel")
align_var = tk.StringVar(value="n")
tk.Radiobutton(root, text="Yes", variable=align_var, value="y").grid(row=current_row, column=1, padx=5, pady=5)
tk.Radiobutton(root, text="No", variable=align_var, value="n").grid(row=current_row, column=2, padx=5, pady=5)
current_row += 1

# Alignment channel radio buttons
align_chan_label = tk.Label(root, text="Alignment channel:")
align_chan_label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
CreateToolTip(align_chan_label, "Select which channel to output aligned intensities")
align_chan_var = tk.StringVar(value="red")
tk.Radiobutton(root, text="Red", variable=align_chan_var, value="red").grid(row=current_row, column=1, padx=5, pady=5)
tk.Radiobutton(root, text="Green", variable=align_chan_var, value="green").grid(row=current_row, column=2, padx=5, pady=5)
current_row += 1

# Frames before drop entry
frames_label = tk.Label(root, text="Frames before 2nd drop:")
frames_label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
CreateToolTip(frames_label, "Number of frames before 2nd stage change to include in aligned data")
frames_entry = tk.Entry(root, width=10)
frames_entry.insert(0, "5")
frames_entry.grid(row=current_row, column=1, padx=5, pady=5)
current_row += 1

# --- Other parameters ---
other_labels = ["Seconds per frame:", "New filename:"]
other_tips = ["Seconds per frame", "Folder name where saved data will go"]
default_values = ["5",""]
for label_text, tip, default in zip(other_labels, other_tips, default_values):
    label = tk.Label(root, text=label_text)
    label.grid(row=current_row, column=0, padx=5, pady=5, sticky='w')
    CreateToolTip(label, tip)

    entry = tk.Entry(root, width=20)
    entry.insert(0, default)
    entry.grid(row=current_row, column=1, padx=5, pady=5)
    entry_widgets[label_text] = entry
    current_row += 1

# --- Message box ---
message_entry = tk.Entry(root, width=80)
message_entry.grid(row=current_row, column=0, columnspan=3, padx=5, pady=5)
current_row += 1

# -------------------------------
# FUNCTIONS
# -------------------------------
def submit():
    global RedFile, GreenFile, Fit_targ, saveplot, sec_per_f, out_file_name
    global Align, Align_channel, FramesBefore

    RedFile = entry_widgets["Red Channel File:"].get()
    GreenFile = entry_widgets["Green Channel File:"].get()
    Fit_targ = fit_var.get()
    saveplot = save_var.get()
    sec_per_f = int(entry_widgets["Seconds per frame:"].get())
    out_file_name = entry_widgets["New filename:"].get() or os.path.splitext(os.path.basename(RedFile or GreenFile))[0]

    # New parameters
    Align = align_var.get()
    Align_channel = align_chan_var.get()
    try:
        FramesBefore = int(frames_entry.get())
    except:
        FramesBefore = 5

    message_entry.delete(0, tk.END)
    message_entry.insert(0, f"Parameters updated: HMM={Fit_targ}, SavePlot={saveplot}, Align={Align}, AlignChan={Align_channel}, FramesBefore={FramesBefore}, Output={out_file_name}")

def finish():
    root.destroy()
    root.quit()

def makeplot(x, y_G, y_R, y_HMM, particle_id, folder_name):
    plt.figure(dpi=100)
    plt.scatter(x*sec_per_f/60, y_R, color='r', label='Red')
    plt.scatter(x*sec_per_f/60, y_G, color='g', label='Green')
    plt.plot(x*sec_per_f/60, y_HMM, color='b', label='HMM fit')
    plt.xlabel('Time (min)')
    plt.ylabel('Normalized Intensity (%)')
    plt.legend()
    plt.tight_layout()
    if saveplot=='y':
        plt.savefig(os.path.join(folder_name, f"Particle_{particle_id}.png"))
    plt.close()

def HMM_fit(trace, min_state_length=2):
    rrr = np.array(trace)
    models, bic = [], []
    for n_states in range(1,5):
        scores, candidates = [], []
        for seed in range(20):
            model = hmm.GaussianHMM(n_components=n_states,covariance_type='diag',n_iter=10,random_state=seed)
            model.fit(rrr.reshape(-1,1))
            candidates.append(model)
            scores.append(model.score(rrr.reshape(-1,1)))
        best_model = candidates[np.argmax(scores)]
        models.append(best_model)
        bic.append(best_model.bic(rrr.reshape(-1,1)))
    evalu_b = models[np.argmin(bic)]
    states = evalu_b.predict(rrr.reshape(-1,1))

    cleaned = states.copy()
    start=0
    while start<len(states):
        current = cleaned[start]
        end = start
        while end+1<len(states) and cleaned[end+1]==current:
            end+=1
        if end-start+1<min_state_length and start>0:
            cleaned[start:end+1] = cleaned[start-1]
        start = end+1

    hmm_result = evalu_b.means_[cleaned].reshape(len(rrr))
    tran_num = np.sum(np.diff(hmm_result)!=0)
    res = rrr - hmm_result
    return hmm_result, res, tran_num, evalu_b

def load_channel_files(red_file=None, green_file=None):
    if not red_file and not green_file:
        raise ValueError("At least one file must be provided.")

    RedData = None
    GreenData = None

    if red_file:
        df = pd.read_excel(red_file, header=None).T  # transpose so rows = frames
        df.index.name = 'Frame'
        df.columns = range(df.shape[1])  # columns = Particle
        RedData = df.reset_index().melt(id_vars='Frame', var_name='Particle', value_name='Mean Red Intensity')
        RedData['Particle'] = RedData['Particle'].astype(int)
        RedData = RedData[RedData['Particle'] != 0]  # remove particle 0

    if green_file:
        df = pd.read_excel(green_file, header=None).T  # transpose
        df.index.name = 'Frame'
        df.columns = range(df.shape[1])
        GreenData = df.reset_index().melt(id_vars='Frame', var_name='Particle', value_name='Mean Green Intensity')
        GreenData['Particle'] = GreenData['Particle'].astype(int)
        GreenData = GreenData[GreenData['Particle'] != 0]  # remove particle 0

    return RedData, GreenData

def process():
    submit()  # update GUI parameters
    global tracks  # merged DataFrame containing both channels with Particle, Frame, and intensities

    # Load Red and Green channel files
    RedData, GreenData = load_channel_files(red_file=RedFile, green_file=GreenFile)

    # Merge red and green into one DataFrame
    if RedData is not None and GreenData is not None:
        tracks = pd.merge(RedData, GreenData, on=['Frame','Particle'], how='outer')
    elif RedData is not None:
        tracks = RedData.copy()
        tracks['Mean Green Intensity'] = np.nan
    else:
        tracks = GreenData.copy()
        tracks['Mean Red Intensity'] = np.nan

    tracks = tracks[tracks['Particle'] != 0]
    tracks = tracks.sort_values(['Frame','Particle']).reset_index(drop=True)

    effective, lysis, uncoating, bad_particles, nonlytic = [], [], [], [], []
    lysis_times, uncoating_env_times, uncoating_capsid_times = [], [], []

    # Create folders if saving plots
    if saveplot == 'y':
        folder = filedialog.askdirectory(title="Select folder to save plots")
        new_folder = os.path.join(folder, f"{out_file_name}_plots")
        os.makedirs(new_folder, exist_ok=True)
        for sub in ['bad','lysis','uncoating','Nonlytic']:
            os.makedirs(os.path.join(new_folder, sub), exist_ok=True)
    else:
        new_folder = '.'

    # Loop over particles by Particle ID
    particle_list = tracks['Particle'].unique()
    total_particles = len(particle_list)

    for i, particle in enumerate(particle_list, 1):
        print(f"Processing particle {i}/{total_particles}...")  # progress indicator

        pdata = tracks[tracks['Particle'] == particle].sort_values('Frame').reset_index(drop=True)
        temp_X = pdata['Frame'] * sec_per_f / 60  # convert frame to minutes
        temp_R = pdata['Mean Red Intensity']
        temp_G = pdata['Mean Green Intensity']

        # Normalize traces
        Norm_R = temp_R / temp_R.iloc[0] * 100
        Norm_G = temp_G / temp_G.iloc[0] * 100

        # HMM fitting with error suppression
        trace_to_fit = Norm_R if Fit_targ == 'red' else Norm_G
        try:
            hmm_result, res, tran_num, evalu_b = HMM_fit(trace_to_fit)
        except Exception as e:
            print(f"  Particle {particle} HMM fit failed: {e}")
            bad_particles.append(particle)
            folder_bad = os.path.join(new_folder, 'bad') if saveplot=='y' else '.'
            makeplot(temp_X, Norm_G, Norm_R, np.zeros_like(temp_X), particle, folder_bad)
            continue

        folder_bad = os.path.join(new_folder, 'bad') if saveplot=='y' else '.'

        # Classify particle
        if tran_num >= 5 or hmm_result[-1] > hmm_result[0]:
            makeplot(temp_X, Norm_G, Norm_R, hmm_result, particle, folder_bad)
            bad_particles.append(particle)
            continue

        effective.append(particle)

        # Lysis detection
        if tran_num == 1 and hmm_result[0] > hmm_result[-1]:
            Env = temp_X[np.argmax(np.sort(evalu_b.means_.reshape(evalu_b.n_components)))]
            lysis.append(particle)
            lysis_times.append(Env)
            folder_lysis = os.path.join(new_folder, 'lysis') if saveplot=='y' else '.'
            makeplot(temp_X, Norm_G, Norm_R, hmm_result, particle, folder_lysis)
            continue

        # Uncoating detection
        if evalu_b.n_components == 3 and 1 < tran_num < 5 and max(hmm_result[:2]) == max(hmm_result) and hmm_result[-1] == min(hmm_result):
            sorted_means = np.sort(evalu_b.means_.reshape(evalu_b.n_components))
            Env = temp_X[np.where(hmm_result == sorted_means[-1])[0][0]]
            t_unco = temp_X[np.where(hmm_result == sorted_means[-2])[0][0]]
            Cap = t_unco - Env
            if Cap > 0:
                uncoating.append(particle)
                uncoating_env_times.append(Env)
                uncoating_capsid_times.append(Cap)
                folder_unco = os.path.join(new_folder, 'uncoating') if saveplot=='y' else '.'
                makeplot(temp_X, Norm_G, Norm_R, hmm_result, particle, folder_unco)
            else:
                # Treat as lysis
                lysis.append(particle)
                lysis_times.append(Env)
                folder_lysis = os.path.join(new_folder, 'lysis') if saveplot=='y' else '.'
                makeplot(temp_X, Norm_G, Norm_R, hmm_result, particle, folder_lysis)
            continue

        # Nonlytic particles
        folder_non = os.path.join(new_folder, 'Nonlytic') if saveplot == 'y' else '.'
        makeplot(temp_X, Norm_G, Norm_R, hmm_result, particle, folder_non)
        nonlytic.append(particle)

    print("\nHMM complete. Writing Excel outputs...")

    # === Save per-particle file ===
    results_folder = os.path.join(new_folder, "results")
    os.makedirs(results_folder, exist_ok=True)
    particle_xlsx = os.path.join(results_folder, f"{out_file_name}_particles.xlsx")

    writer = pd.ExcelWriter(particle_xlsx, engine='xlsxwriter')
    subset_tracks = tracks[tracks['Particle'].isin(effective)]
    for a in subset_tracks['Particle'].unique():
        sub = subset_tracks[subset_tracks['Particle'] == a].copy()
        sub['Time (min)'] = sub['Frame'] * sec_per_f / 60
        sub = sub[['Particle', 'Frame', 'Time (min)', 'Mean Red Intensity', 'Mean Green Intensity']]
        if len(sub) > 30:
            sub.to_excel(writer, sheet_name=f'Particle{a}', index=False)
    writer.close()

    # === Build category summary dataframes ===
    categories = {
        "Lysis": lysis,
        "Uncoating": uncoating,
        "Nonlytic": nonlytic
    }
    
    for cat_name, plist in categories.items():
        if not plist:
            continue
    
        cat_df = tracks[tracks['Particle'].isin(plist)].copy()
        cat_df['Time (min)'] = cat_df['Frame'] * sec_per_f / 60
    
        red_pivot = cat_df.pivot(index='Frame', columns='Particle', values='Mean Red Intensity')
        green_pivot = cat_df.pivot(index='Frame', columns='Particle', values='Mean Green Intensity')
    
        cat_xlsx = os.path.join(results_folder, f"{out_file_name}_{cat_name}.xlsx")
        with pd.ExcelWriter(cat_xlsx, engine='xlsxwriter') as cat_writer:
            red_pivot.to_excel(cat_writer, sheet_name='Red')
            green_pivot.to_excel(cat_writer, sheet_name='Green')

    # === Build aligned normalized Excel for uncoating particles ===
    aligned_raw_R, aligned_raw_G = {}, {}
    aligned_norm_R, aligned_norm_G = {}, {}

    for particle in uncoating:
        pdata = tracks[tracks['Particle'] == particle].sort_values('Frame').reset_index(drop=True)
        drop_channel = pdata['Mean Red Intensity'] if Fit_targ == 'red' else pdata['Mean Green Intensity']
        align_channel_R = pdata['Mean Red Intensity']
        align_channel_G = pdata['Mean Green Intensity']

        # HMM fit for drop detection
        hmm_result, _, _, evalu_b = HMM_fit(drop_channel)
        stage_changes = np.where(np.diff(hmm_result) != 0)[0]
        if len(stage_changes) < 2:
            continue

        second_drop_idx = stage_changes[1] + 1
        start_idx = max(second_drop_idx - FramesBefore, 0)

        # Slice frames: from start_idx to end
        aligned_slice_R = align_channel_R.iloc[start_idx:].copy()
        aligned_slice_G = align_channel_G.iloc[start_idx:].copy()

        # Normalize based on mean of frames before 2nd drop
        pre_mean_R = aligned_slice_R.iloc[:FramesBefore].mean()
        pre_mean_G = aligned_slice_G.iloc[:FramesBefore].mean()
        normalized_R = aligned_slice_R / pre_mean_R * 100
        normalized_G = aligned_slice_G / pre_mean_G * 100

        aligned_raw_R[particle] = aligned_slice_R.values
        aligned_raw_G[particle] = aligned_slice_G.values
        aligned_norm_R[particle] = normalized_R.values
        aligned_norm_G[particle] = normalized_G.values

    # Helper to build padded DataFrame with Frame & Time
    def build_aligned_tab(aligned_dict):
        if not aligned_dict:
            return pd.DataFrame()
        max_len = max(len(v) for v in aligned_dict.values())
        padded = {k: np.pad(np.array(v, dtype=float), (0, max_len-len(v)), constant_values=np.nan) 
                  for k,v in aligned_dict.items()}
        frame_col = np.arange(-FramesBefore, -FramesBefore + max_len)
        time_col = frame_col * sec_per_f / 60
        df = pd.DataFrame(padded)
        df.insert(0, 'Time (min)', time_col)
        df.insert(0, 'Frame', frame_col)
        return df

    # --- Save aligned normalized Excel file ---
    aligned_xlsx = os.path.join(results_folder, f"{out_file_name}_Aligned.xlsx")
    with pd.ExcelWriter(aligned_xlsx, engine='xlsxwriter') as writer:
        if aligned_raw_R:
            build_aligned_tab(aligned_raw_R).to_excel(writer, sheet_name='Aligned_Raw_Red', index=False)
        if aligned_raw_G:
            build_aligned_tab(aligned_raw_G).to_excel(writer, sheet_name='Aligned_Raw_Green', index=False)
        if aligned_norm_R:
            build_aligned_tab(aligned_norm_R).to_excel(writer, sheet_name='Aligned_Normalized_Red', index=False)
        if aligned_norm_G:
            build_aligned_tab(aligned_norm_G).to_excel(writer, sheet_name='Aligned_Normalized_Green', index=False)

    print(f"\nAnalysis complete. Files saved in: {results_folder}")


# -------------------------------
# BUTTONS
# -------------------------------
tk.Button(root,text="Update Parameters",command=submit).grid(row=current_row,column=0,padx=5,pady=10)
tk.Button(root,text="Process",command=process).grid(row=current_row,column=1,padx=5,pady=10)
tk.Button(root,text="Finish",command=finish).grid(row=current_row,column=2,padx=5,pady=10)

root.mainloop()
