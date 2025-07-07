import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import os
from matplotlib.ticker import FuncFormatter
import matplotlib.font_manager as fm
import numpy as np
import requests
import matplotlib as mpl
from matplotlib import transforms
import matplotlib.patches as mpatches



# Use Agg backend for better Unicode rendering
mpl.use("agg")

# Load Khmer font
font_path = "FONT/KhmerOSsiemreap.ttf"
font_prop = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()

# Constants
DATA_FILE = "road_maintenance_updated.xlsx"
GOOGLE_SHEET_FILE_ID = "1ESWPe49WlQ1608IH8bACw_XiTiRx4EFTCLvBBW44K_E"
EXCEL_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_FILE_ID}/export?format=xlsx"

# Refresh button to download from Google Sheets
if st.sidebar.button("🔄 Refresh Data"):
    with st.spinner("Downloading latest Excel data from Google Sheet..."):
        r = requests.get(EXCEL_EXPORT_URL)
        if r.status_code == 200:
            with open(DATA_FILE, 'wb') as f:
                f.write(r.content)
            st.success("✅ Excel downloaded from Google Sheet")
        else:
            st.error("❌ Failed to download Excel. Please check sharing permissions.")

# === Load data ===
DATA_FILE = "road_maintenance_updated.xlsx"

if os.path.exists(DATA_FILE):
    df = pd.read_excel(DATA_FILE)

    # Clean Year and Chapter fields
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce").dropna().astype(int).astype(str)
    df["Chapter"] = pd.to_numeric(df["Chapter"], errors="coerce").dropna().astype(int).astype(str)
else:
    st.error("❌ Excel data file not found.")
    st.stop()

# === Sidebar Filters ===
road_ids = sorted(df["Road_ID"].dropna().astype(str).unique())
selected_road = st.sidebar.selectbox("Select Road ID", road_ids)

# PK Range Setup
road_df = df[df["Road_ID"].astype(str) == selected_road]
pk_min = float(road_df["PK_Start"].min())
pk_max = float(road_df["PK_End"].max())

start_input = st.sidebar.number_input("PK Start", min_value=0, max_value=int(pk_max), value=int(pk_min), step=100)
end_input = st.sidebar.number_input("PK End", min_value=pk_min, max_value=pk_max, value=pk_max)

# Year, Type, Chapter Filters
years = sorted(df["Year"].dropna().unique(), reverse=True)
selected_years = st.sidebar.multiselect("Select Year(s) to Show", years, default=years)

types = df["Maintenance_Type"].dropna().astype(str).unique()
selected_types = st.sidebar.multiselect("Select Maintenance Type(s)", types, default=types)

chapters = df["Chapter"].dropna().astype(str).unique()
selected_chapter = st.sidebar.multiselect("Select Chapter(s)", chapters, default=chapters)

# Layer Order Control
request_years = sorted(df[df["Type"] == "Request"]["Year"].dropna().unique(), reverse=True)
available_layers = [f"Approval {y}" for y in selected_years] + [f"Request {y}" for y in request_years if y in selected_years]
layer_order = st.sidebar.multiselect("Set Layer Order (Bottom to Top)", available_layers, default=available_layers[::-1])

# Color Map Setup
st.sidebar.markdown("### Customize Colors")
color_map = {}
default_colors = ["#e74c3c", "#3498db", "#2ecc71", "#9b59b6", "#f39c12", "#1abc9c", "#7f8c8d"]
for i, t in enumerate(types):
    color = st.sidebar.color_picker(f"{t} Color", default_colors[i % len(default_colors)])
    color_map[t] = color

# Font Size Controls
st.sidebar.markdown("### Customize Font Size")
font_size = st.sidebar.slider("Font Size", min_value=6, max_value=20, value=15)
title_font_size = st.sidebar.slider("Title Font Size", min_value=10, max_value=30, value=20)

# Manual Label Input
st.sidebar.markdown("---")
st.sidebar.markdown("### Add Manual Location Label")

manual_labels = []
num_labels = st.sidebar.number_input("How many labels to add?", min_value=0, max_value=10, step=1, value=0)

for i in range(num_labels):
    with st.sidebar.expander(f"Label #{i+1}"):
        label_text = st.text_input(f"Label Text {i+1}", key=f"label_text_{i}")
        label_pk_start = st.number_input(f"Label PK Start {i+1}", min_value=pk_min, max_value=pk_max, key=f"label_pk_start_{i}")
        label_pk_end = st.number_input(f"Label PK End {i+1}", min_value=pk_min, max_value=pk_max, key=f"label_pk_end_{i}")
        label_color = st.color_picker(f"Label Color {i+1}", "#000000", key=f"label_color_{i}")
        manual_labels.append((label_text, label_pk_start, label_pk_end, label_color))

# === Filter Data Based on Inputs ===
filtered = df[
    (df["Road_ID"].astype(str) == str(selected_road)) &
    (df["PK_End"] >= start_input) &
    (df["PK_Start"] <= end_input) &
    (df["Maintenance_Type"].notna()) &
    (df["Maintenance_Type"].astype(str).str.strip() != "") &
    (df["Maintenance_Type"].isin(selected_types)) &
    (df["Chapter"].isin(selected_chapter)) &
    (df["Year"].astype(str).isin([str(y) for y in selected_years])) &
    (
        (df["Type"] != "Approval") |
        ((df["Type"] == "Approval") & df["Year"].notna() & (df["Year"].astype(str).str.strip() != ""))
    )
]

if filtered.empty:
    st.warning("⚠ No data matches your filters. Please adjust selections.")
    st.stop()


# 📊 Preview Chart
st.markdown("### 📊 Preview Chart")

from collections import defaultdict
from matplotlib.ticker import FuncFormatter
import matplotlib.patches as mpatches

# Step 1: Build y_map, label list, and groupings
y_map, y_labels, y_label_colors = {}, [], {}
subrow_tracker = defaultdict(list)
group_titles = []

row_index = 0
for label in layer_order:
    is_request = label.startswith("Request")
    key_prefix = label.replace(" ", "_")

    if is_request:
        year = label.split()[-1]
        filtered_sub = filtered[
            (filtered["Type"] == "Request") & (filtered["Year"].astype(str) == year)
        ]
        unique_types = sorted(filtered_sub["Maintenance_Type"].dropna().map(str.strip).unique())
        if not unique_types:
            continue

        for m_type in unique_types:
            norm_type = m_type.replace(" ", "_")
            sub_key = f"{key_prefix}_{norm_type}"
            y_map[sub_key] = row_index
            y_labels.append(m_type)
            y_label_colors[row_index] = "green"
            subrow_tracker[label].append(row_index)
            row_index += 1

        group_y = sum(subrow_tracker[label]) / len(subrow_tracker[label])
        group_titles.append((label, group_y))
    else:
        y_key = key_prefix
        y_map[y_key] = row_index
        y_labels.append(label)
        y_label_colors[row_index] = "black"
        subrow_tracker[label].append(row_index)
        row_index += 1

# Step 2: Setup plot
fig_preview, ax = plt.subplots(figsize=(16, 6))
label_positions = {}
pk_label_positions = []

# Step 3: Plot segments with sequence
legend_handles = {}
request_segments = filtered[filtered["Type"] == "Request"].sort_values(by="PK_Start").reset_index()
request_seq_map = {row["index"]: i+1 for i, row in request_segments.iterrows()}
sequence_number = 1

for idx, seg in filtered.iterrows():
    mtype = str(seg['Maintenance_Type']).strip()
    color = color_map.get(mtype, "gray")
    mtype_key = mtype.replace(" ", "_")

    if seg["Type"] == "Request":
        key = f"Request_{seg['Year']}_{mtype_key}"
    else:
        key = f"Approval_{seg['Year']}".replace(" ", "_")

    if key not in y_map:
        continue

    y = y_map[key]
    clipped_start = max(seg["PK_Start"], start_input)
    clipped_end = min(seg["PK_End"], end_input)
    if clipped_start >= clipped_end:
        continue

    # Draw bar with black border for requests
    edgecolor = "black" if seg["Type"] == "Request" else "none"
    bar = ax.barh(y, width=clipped_end - clipped_start, left=clipped_start,
                  color=color, edgecolor=edgecolor, height=0.6)
    if mtype not in legend_handles:
        legend_handles[mtype] = bar[0]

    # Add sequence number in circle on bar for request
    if seg["Type"] == "Request":
        seq = request_seq_map.get(idx, "")
        center_x = (clipped_start + clipped_end) / 2
        ax.text(center_x, y, str(seq),
            ha='center', va='center',
            fontsize=10, fontweight='bold',
            fontproperties=font_prop, zorder=5)
        seg_sequence = sequence_number
        sequence_number += 1
    else:
        seg_sequence = ""
# Step 4: Draw color-coded PK labels ON TOP of bars with offset reset per Maintenance Type row, with arrows
base_offset = 0.35     # base offset above the bar
vertical_spacing = 0.15  # spacing between labels

# Group requests by their Y-axis row (maintenance type layer)
grouped_by_y = defaultdict(list)
for i, row in request_segments.iterrows():
    mtype = str(row["Maintenance_Type"]).strip()
    key = f"Request_{row['Year']}_{mtype.replace(' ', '_')}"
    y = y_map.get(key, 0)
    grouped_by_y[y].append((i, row))  # global seq + row

for y, segments in grouped_by_y.items():
    for local_idx, (global_seq, row) in enumerate(segments):
        pk_start, pk_end = row["PK_Start"], row["PK_End"]
        label_x = (pk_start + pk_end) / 2
        mtype = str(row["Maintenance_Type"]).strip()
        label_color = color_map.get(mtype, "gray")

        seq = global_seq + 1  # label index
        pk_label = f"({seq}). {int(pk_start//1000)}+{int(pk_start%1000):03d} to {int(pk_end//1000)}+{int(pk_end%1000):03d}"

        # Vertical offset reset per row
        y_offset = y + base_offset + local_idx * vertical_spacing

        # Draw label with white background
        ax.text(label_x, y_offset, pk_label,
                ha='center', va='bottom', fontsize=6, color=label_color,
                fontproperties=font_prop,
                bbox=dict(boxstyle="round,pad=0.2", facecolor="white", edgecolor="none", alpha=0.9),
                clip_on=False)

        # Draw arrow from label to bar center
        ax.annotate("", xy=(label_x, y + 0.3), xytext=(label_x, y_offset - 0.02),
                    arrowprops=dict(arrowstyle="->", color=label_color, linewidth=0.8))

        # Draw dashed boundary lines
        ax.axvline(pk_start, linestyle="dashed", color=label_color, alpha=0.6)
        ax.axvline(pk_end, linestyle="dashed", color=label_color, alpha=0.6)

# Step 5: Manual Labels
for label_text, pk_start, pk_end, label_color in manual_labels:
    if start_input > pk_end or end_input < pk_start:
        continue
    clipped_start = max(pk_start, start_input)
    clipped_end = min(pk_end, end_input)
    if clipped_start >= clipped_end:
        continue
    label_x = (clipped_start + clipped_end) / 2
    ax.axvline(pk_start, linestyle="dashed", color=label_color, linewidth=1.2, alpha=0.8)
    ax.axvline(pk_end, linestyle="dashed", color=label_color, linewidth=1.2, alpha=0.8)
    ax.text(label_x, max(y_map.values()) + 0.5, label_text,
            ha='center', va='bottom', fontsize=12,
            fontweight='bold', color=label_color, backgroundcolor='white', clip_on=False)

# Step 6: Axes formatting
def format_pk(x, pos):
    return f"{int(x // 1000)}+{int(x % 1000):03d}"
ax.xaxis.set_major_formatter(FuncFormatter(format_pk))
ax.set_xlim(start_input, end_input)
ax.set_ylim(-1, max(y_map.values()) + 1)
ax.set_title(f"\n\nRoad: {selected_road}", fontproperties=font_prop, fontsize=title_font_size)
ax.grid(True)

# Step 7: Y labels
x_offset = (end_input - start_input) * 0.015  # ~1.5% of chart width
ax.set_yticks(list(y_map.values()))
for y_val, label in zip(y_map.values(), y_labels):
    ax.text(start_input - x_offset, y_val, label,
            va='center', ha='right', fontsize=font_size,
            fontproperties=font_prop, color=y_label_colors[y_val])

# Step 8: Group vertical labels
x_offset = (end_input - start_input) * 0.01  # 5% of total width
label_x = start_input + x_offset
for group_label, y_pos in group_titles:
    ax.text(label_x, y_pos, group_label,
            fontsize=font_size + 1, fontproperties=font_prop,
            color="green", ha='center', va='center', rotation=90,
            bbox=dict(boxstyle="round,pad=0.3", edgecolor="green", facecolor="none"),
            clip_on=False)

# Step 9: Request bounding boxes
for label, rows in subrow_tracker.items():
    if not label.startswith("Request") or not rows:
        continue
    top = max(rows) + 0.5
    bottom = min(rows) - 0.5
    ax.hlines([bottom, top], xmin=start_input, xmax=end_input, color='green', linewidth=1.5)
    ax.vlines([start_input, end_input], ymin=bottom, ymax=top, color='green', linewidth=1.5)

# Step 10: Adjust layout and add legend
plt.subplots_adjust(left=0.15, right=0.88)
if legend_handles:
    ax.legend(legend_handles.values(), legend_handles.keys(),
              title="Maintenance Type", loc="upper right")

# Show chart
st.pyplot(fig_preview)

# Summary table
st.markdown("### 📊 Maintenance Summary by Section")

sum_df = filtered.copy()

# Assign sequence number based on PK_Start order within Request type
sum_df["Seq"] = 0
is_request = sum_df["Type"] == "Request"
sum_df.loc[is_request, "Seq"] = sum_df[is_request].sort_values(by="PK_Start").reset_index(drop=True).index + 1

# Format PK label with sequence for Requests
def format_pk_label(row):
    label = f"{int(row['PK_Start']//1000)}+{int(row['PK_Start']%1000):03d} to {int(row['PK_End']//1000)}+{int(row['PK_End']%1000):03d}"
    if row["Type"] == "Request":
        return f"({int(row['Seq'])}). {label}"
    else:
        return label

sum_df["PK_Label"] = sum_df.apply(format_pk_label, axis=1)

# Calculate distance
sum_df["Distance_km"] = ((sum_df["PK_End"] - sum_df["PK_Start"]) / 1000).round(2)
sum_df["Group"] = sum_df["Type"]
sum_df["Maintenance"] = sum_df.apply(lambda row: f"{row['Maintenance_Type']} [{row['Type']}]", axis=1)

# Group and join PKs by maintenance type
sum_grouped = sum_df.groupby(["Group", "Maintenance"]).agg({
    "PK_Label": lambda x: ", ".join(x),
    "Distance_km": "sum"
}).reset_index()

# Rename and format columns
sum_grouped = sum_grouped.rename(columns={
    "Maintenance": "Maintenance [Type]",
    "PK_Label": "PK Range",
    "Distance_km": "Total Distance (km)"
})
sum_grouped["Total Distance (km)"] = sum_grouped["Total Distance (km)"].round(2)

# Show in Streamlit
st.dataframe(sum_grouped)


# === Auto PDF Export Section ===
import io
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import FuncFormatter
from matplotlib.gridspec import GridSpec

export_base = f"{selected_road}_{int(start_input)}_{int(end_input)}"

# Function: Generate Chart Figure
def generate_chart_figure():
    fig, ax1 = plt.subplots(figsize=(16.5, 8))
    ax1.set_title(f"\nRoad: {selected_road}", fontproperties=font_prop, fontsize=title_font_size)

    request_segments = filtered[filtered["Type"] == "Request"].sort_values(by="PK_Start").reset_index()
    request_seq_map = {row["index"]: i + 1 for i, row in request_segments.iterrows()}

    for idx, seg in filtered.iterrows():
        mtype = str(seg['Maintenance_Type']).strip()
        color = color_map.get(mtype, "gray")
        mtype_key = mtype.replace(" ", "_")
        key = f"{seg['Type']}_{seg['Year']}_{mtype_key}" if seg["Type"] == "Request" else f"Approval_{seg['Year']}".replace(" ", "_")
        if key not in y_map:
            continue
        y = y_map[key]
        clipped_start = max(seg["PK_Start"], start_input)
        clipped_end = min(seg["PK_End"], end_input)
        if clipped_start >= clipped_end:
            continue
        edgecolor = "black" if seg["Type"] == "Request" else "none"
        ax1.barh(y, clipped_end - clipped_start, left=clipped_start,
                 color=color, edgecolor=edgecolor, height=0.6)
        if seg["Type"] == "Request":
            seq = request_seq_map.get(idx, "")
            center_x = (clipped_start + clipped_end) / 2
            ax1.text(center_x, y, str(seq), ha='center', va='center',
                     fontsize=10, fontweight='bold', fontproperties=font_prop, zorder=5)

    ax1.set_yticks(list(y_map.values()))
    x_offset = (end_input - start_input) * 0.015
    for y_val, label in zip(y_map.values(), y_labels):
        ax1.text(start_input - x_offset, y_val, label,
                 va='center', ha='right', fontsize=font_size,
                 fontproperties=font_prop, color=y_label_colors[y_val])

    label_x = start_input + (end_input - start_input) * 0.01
    for group_label, y_pos in group_titles:
        ax1.text(label_x, y_pos, group_label,
                 fontsize=font_size + 1, fontproperties=font_prop,
                 color="green", ha='center', va='center', rotation=90,
                 bbox=dict(boxstyle="round,pad=0.3", edgecolor="green", facecolor="none"))

    ax1.xaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{int(x // 1000)}+{int(x % 1000):03d}"))
    ax1.set_xlim(start_input, end_input)
    ax1.set_ylim(-0.5, max(y_map.values()) + 1.5)
    ax1.grid(True)

    if legend_handles:
        ax1.legend(legend_handles.values(), legend_handles.keys(),
                   title="Maintenance Type", loc="upper right")

    return fig

# Function: Generate Summary Table Figure
def generate_summary_figure():
    fig, ax = plt.subplots(figsize=(16.5, 6))
    ax.axis("off")

    table_data = sum_grouped.copy()
    table_data["PK Range"] = table_data["PK Range"].apply(lambda x: "\n".join(x[i:i+60] for i in range(0, len(x), 60)))
    data_matrix = table_data.values.tolist()
    col_labels = list(table_data.columns)

    table = ax.table(cellText=data_matrix, colLabels=col_labels, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(13)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(color='white', weight='bold')
            cell.set_facecolor('#003366')
        elif "Request" in str(data_matrix[row - 1][0]):
            cell.set_facecolor('#e5f5e5')
        else:
            cell.set_facecolor('white')

    table.scale(1, 2.0)
    return fig

# Function: Export to PDF in memory
def export_pdf(figs: list):
    buffer = io.BytesIO()
    with PdfPages(buffer) as pdf:
        for fig in figs:
            pdf.savefig(fig, bbox_inches='tight')
    return buffer.getvalue()

# === EXPORT 1: Chart + Summary
full_pdf = export_pdf([generate_chart_figure(), generate_summary_figure()])
st.download_button(
    label="📥 Download Chart + Summary",
    data=full_pdf,
    file_name=f"{export_base}_full.pdf",
    mime="application/pdf"
)

# === EXPORT 2: Chart Only
chart_pdf = export_pdf([generate_chart_figure()])
st.download_button(
    label="📈 Download Chart Only",
    data=chart_pdf,
    file_name=f"{export_base}_chart.pdf",
    mime="application/pdf"
)

# === EXPORT 3: Summary Only
summary_pdf = export_pdf([generate_summary_figure()])
st.download_button(
    label="📋 Download Summary Only",
    data=summary_pdf,
    file_name=f"{export_base}_summary.pdf",
    mime="application/pdf"
)
