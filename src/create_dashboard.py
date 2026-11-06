"""
create_dashboard.py - Generates the professional 4-sheet Excel dashboard:
1. Dashboard (16:9 layout, 7 KPI cards, 7 charts, summary tables)
2. Predictions (30,000 test predictions, table formatting, conditional formatting)
3. Model Performance (metrics, confusion matrix, feature importance table & chart)
4. Data (150,000 cleaned records, table formatting, filters)
"""

import os
import sys
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
import pandas as pd
import numpy as np
import joblib
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, confusion_matrix

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SRC_DIR)
DATA_PATH = os.path.join(BASE_DIR, "data", "ncr_ride_bookings.csv")
MODEL_PATH = os.path.join(BASE_DIR, "models", "decision_tree_model.pkl")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "Uber_Ride_ML_Dashboard_Final.xlsx")

from train_model import (
    load_and_preprocess_data,
    TEST_SIZE,
    RANDOM_STATE,
    DT_MAX_DEPTH,
    DT_MIN_SAMPLES_LEAF,
    INPUT_FEATURES,
    FEATURE_LABELS
)

NAVY_HEADER = "1B365D"
STEEL_BLUE  = "2C5E8A"
LIGHT_BLUE  = "EBF1F5"
WHITE       = "FFFFFF"
BORDER_GRAY = "D1D5DB"
TEXT_DARK   = "1F2937"
TEXT_MUTED  = "4B5563"
SUCCESS_GRN = "155724"
SUCCESS_BG  = "D4EDDA"
DANGER_RED  = "721C24"
DANGER_BG   = "F8D7DA"
GOLD_TEXT   = "856404"
GOLD_BG     = "FFF3CD"

def thin_border():
    thin = Side(border_style="thin", color=BORDER_GRAY)
    return Border(top=thin, left=thin, right=thin, bottom=thin)

def auto_fit_cols(ws, max_cols=None, sample_rows=50, min_width=12):
    for col in ws.columns:
        col_idx = col[0].column
        if max_cols and col_idx > max_cols:
            break
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in col[:sample_rows]:
            val = str(cell.value or '')
            if '\n' in val:
                val = max(val.split('\n'), key=len)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 3, min_width)

def build_dashboard(wb, df, metrics, test_df):
    ws = wb.create_sheet(title="Dashboard", index=0)
    ws.views.sheetView[0].showGridLines = True

    ws.column_dimensions["A"].width = 3
    for c in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[c].width = 11.5
    ws.column_dimensions["I"].width = 3.5
    for c in ["J", "K", "L", "M", "N", "O"]:
        ws.column_dimensions[c].width = 11.5
    ws.column_dimensions["P"].width = 3

    # Title
    ws.merge_cells("B2:O2")
    ws["B2"] = "UBER RIDE PREDICTION & OPERATIONS DASHBOARD"
    ws["B2"].font = Font(name="Segoe UI", size=18, bold=True, color=WHITE)
    ws["B2"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 34

    ws.merge_cells("B3:O3")
    ws["B3"] = "Machine Learning Predictive System | Decision Tree Classification | NCR Ride Dataset"
    ws["B3"].font = Font(name="Segoe UI", size=10, italic=True, color=WHITE)
    ws["B3"].fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 8

    # KPI Cards
    total_bookings = len(df)
    successful_bookings = int((df['Actual_YES_NO'] == 'YES').sum())
    unsuccessful_bookings = int((df['Actual_YES_NO'] == 'NO').sum())
    success_rate = successful_bookings / total_bookings
    accuracy = metrics['accuracy']
    avg_dist = df['Ride_Distance'].mean()
    avg_val = df['Booking_Value'].mean()

    kpis = [
        ("TOTAL BOOKINGS", f"{total_bookings:,}", "All fleet bookings", "B", "C"),
        ("SUCCESSFUL BOOKINGS", f"{successful_bookings:,}", "Status = Completed", "D", "E"),
        ("UNSUCCESSFUL BOOKINGS", f"{unsuccessful_bookings:,}", "Cancelled / Incomplete", "F", "G"),
        ("SUCCESS RATE", f"{success_rate:.1%}", "Completed / Total", "H", "I"),
        ("MODEL ACCURACY", f"{accuracy:.2%}", "Test set accuracy", "J", "K"),
        ("AVG RIDE DISTANCE", f"{avg_dist:.2f} km", "Average trip length", "L", "M"),
        ("AVG BOOKING VALUE", f"INR {avg_val:.2f}", "Average fare amount", "N", "O")
    ]

    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 28
    ws.row_dimensions[7].height = 16

    for title, val_str, sub_str, c_start, c_end in kpis:
        ws.merge_cells(f"{c_start}5:{c_end}5")
        ws.merge_cells(f"{c_start}6:{c_end}6")
        ws.merge_cells(f"{c_start}7:{c_end}7")

        ws[f"{c_start}5"].value = title
        ws[f"{c_start}5"].font = Font(name="Segoe UI", size=8, bold=True, color=TEXT_MUTED)
        ws[f"{c_start}5"].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
        ws[f"{c_start}5"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"{c_start}6"].value = val_str
        ws[f"{c_start}6"].font = Font(name="Segoe UI", size=13, bold=True, color=NAVY_HEADER)
        ws[f"{c_start}6"].fill = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
        ws[f"{c_start}6"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"{c_start}7"].value = sub_str
        ws[f"{c_start}7"].font = Font(name="Segoe UI", size=7.5, italic=True, color=TEXT_MUTED)
        ws[f"{c_start}7"].fill = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
        ws[f"{c_start}7"].alignment = Alignment(horizontal="center", vertical="center")

        for c in [c_start, c_end]:
            ws[f"{c}5"].border = Border(top=Side(style="medium", color=STEEL_BLUE), left=Side(style="thin", color=BORDER_GRAY), right=Side(style="thin", color=BORDER_GRAY))
            ws[f"{c}6"].border = Border(left=Side(style="thin", color=BORDER_GRAY), right=Side(style="thin", color=BORDER_GRAY))
            ws[f"{c}7"].border = Border(bottom=Side(style="thin", color=BORDER_GRAY), left=Side(style="thin", color=BORDER_GRAY), right=Side(style="thin", color=BORDER_GRAY))

    ws.row_dimensions[8].height = 12

    # Reference Tables at Row 81+
    build_ref_tables(ws, df, test_df)

    # 7 Charts
    CHART_W, CHART_H = 15.6, 8.5

    # Chart 1
    d1 = DoughnutChart()
    d1.title = "Booking Success Status"
    d1.style = 10
    d1.width, d1.height = CHART_W, CHART_H
    d1.add_data(Reference(ws, min_col=3, min_row=84, max_row=86), titles_from_data=True)
    d1.set_categories(Reference(ws, min_col=2, min_row=85, max_row=86))
    ws.add_chart(d1, "B9")

    # Chart 2
    b2 = BarChart()
    b2.type = "col"
    b2.style = 10
    b2.title = "Bookings by Vehicle Type"
    b2.y_axis.title = "Total Bookings"
    b2.x_axis.title = "Vehicle Type"
    b2.width, b2.height = CHART_W, CHART_H
    b2.legend = None
    b2.add_data(Reference(ws, min_col=6, min_row=84, max_row=91), titles_from_data=True)
    b2.set_categories(Reference(ws, min_col=5, min_row=85, max_row=91))
    ws.add_chart(b2, "J9")

    ws.row_dimensions[25].height = 14

    # Chart 3
    b3 = BarChart()
    b3.type = "col"
    b3.style = 11
    b3.title = "Actual vs Predicted Results (Test Set)"
    b3.y_axis.title = "Number of Bookings"
    b3.x_axis.title = "Success Outcome"
    b3.width, b3.height = CHART_W, CHART_H
    b3.add_data(Reference(ws, min_col=9, min_row=84, max_col=10, max_row=86), titles_from_data=True)
    b3.set_categories(Reference(ws, min_col=8, min_row=85, max_row=86))
    ws.add_chart(b3, "B26")

    # Chart 4
    l4 = LineChart()
    l4.title = "Hourly Booking Demand (24-Hour Cycle)"
    l4.style = 13
    l4.y_axis.title = "Booking Volume"
    l4.x_axis.title = "Hour of Day (00:00 - 23:00)"
    l4.width, l4.height = CHART_W, CHART_H
    l4.legend = None
    l4.add_data(Reference(ws, min_col=3, min_row=96, max_row=120), titles_from_data=True)
    l4.set_categories(Reference(ws, min_col=2, min_row=97, max_row=120))
    ws.add_chart(l4, "J26")

    ws.row_dimensions[42].height = 14

    # Chart 5
    b5 = BarChart()
    b5.type = "col"
    b5.style = 12
    b5.title = "Average Booking Value by Vehicle (INR)"
    b5.y_axis.title = "Average Value (INR)"
    b5.x_axis.title = "Vehicle Type"
    b5.width, b5.height = CHART_W, CHART_H
    b5.legend = None
    b5.add_data(Reference(ws, min_col=6, min_row=96, max_row=103), titles_from_data=True)
    b5.set_categories(Reference(ws, min_col=5, min_row=97, max_row=103))
    ws.add_chart(b5, "B43")

    # Chart 6
    b6 = BarChart()
    b6.type = "col"
    b6.style = 14
    b6.title = "Average Ride Distance by Vehicle (km)"
    b6.y_axis.title = "Average Distance (km)"
    b6.x_axis.title = "Vehicle Type"
    b6.width, b6.height = CHART_W, CHART_H
    b6.legend = None
    b6.add_data(Reference(ws, min_col=9, min_row=96, max_row=103), titles_from_data=True)
    b6.set_categories(Reference(ws, min_col=8, min_row=97, max_row=103))
    ws.add_chart(b6, "J43")

    ws.row_dimensions[59].height = 14

    # Chart 7
    b7 = BarChart()
    b7.type = "bar"
    b7.style = 10
    b7.title = "Top 10 Pickup Locations by Volume"
    b7.x_axis.title = "Total Bookings"
    b7.y_axis.title = "Pickup Location"
    b7.width, b7.height = CHART_W, 9.2
    b7.legend = None
    b7.add_data(Reference(ws, min_col=13, min_row=84, max_row=94), titles_from_data=True)
    b7.set_categories(Reference(ws, min_col=12, min_row=85, max_row=94))
    ws.add_chart(b7, "B60")

    # Executive Tables beside Chart 7
    build_exec_tables(ws, df, test_df)

def build_exec_tables(ws, df, test_df):
    def sub_hdr(cell, text):
        cell.value = text
        cell.font = Font(name="Segoe UI", size=8.5, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def cell_fmt(cell, val, fmt=None, align="center", bold=False):
        cell.value = val
        cell.font = Font(name="Segoe UI", size=8.5, bold=bold, color=TEXT_DARK)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border = thin_border()
        if fmt:
            cell.number_format = fmt

    # Table 1: Actual vs Predicted
    ws.merge_cells("J60:L60")
    ws["J60"] = "SUMMARY: ACTUAL VS PREDICTED"
    ws["J60"].font = Font(name="Segoe UI", size=9, bold=True, color=NAVY_HEADER)

    sub_hdr(ws["J61"], "Outcome")
    sub_hdr(ws["K61"], "Actual")
    sub_hdr(ws["L61"], "Predicted")

    act_yes = int((test_df['Actual_YES_NO'] == 'YES').sum())
    act_no  = int((test_df['Actual_YES_NO'] == 'NO').sum())
    pred_yes = int((test_df['Predicted_YES_NO'] == 'YES').sum())
    pred_no  = int((test_df['Predicted_YES_NO'] == 'NO').sum())

    cell_fmt(ws["J62"], "Successful (YES)", align="left")
    cell_fmt(ws["K62"], act_yes, fmt="#,##0")
    cell_fmt(ws["L62"], pred_yes, fmt="#,##0")

    cell_fmt(ws["J63"], "Unsuccessful (NO)", align="left")
    cell_fmt(ws["K63"], act_no, fmt="#,##0")
    cell_fmt(ws["L63"], pred_no, fmt="#,##0")

    cell_fmt(ws["J64"], "Total Test Rides", align="left", bold=True)
    cell_fmt(ws["K64"], act_yes + act_no, fmt="#,##0", bold=True)
    cell_fmt(ws["L64"], pred_yes + pred_no, fmt="#,##0", bold=True)

    # Table 2: Vehicle Performance
    ws.merge_cells("M60:O60")
    ws["M60"] = "VEHICLE PERFORMANCE METRICS"
    ws["M60"].font = Font(name="Segoe UI", size=9, bold=True, color=NAVY_HEADER)

    sub_hdr(ws["M61"], "Vehicle Type")
    sub_hdr(ws["N61"], "Avg Dist (km)")
    sub_hdr(ws["O61"], "Avg Value (INR)")

    v_sum = df.groupby('Vehicle Type').agg({'Ride_Distance': 'mean', 'Booking_Value': 'mean'}).reset_index()
    for idx, r in v_sum.iterrows():
        row_n = 62 + idx
        cell_fmt(ws[f"M{row_n}"], str(r['Vehicle Type']), align="left")
        cell_fmt(ws[f"N{row_n}"], round(float(r['Ride_Distance']), 2), fmt="0.00")
        cell_fmt(ws[f"O{row_n}"], round(float(r['Booking_Value']), 2), fmt="INR #,##0.00")

    # Table 3: Top 5 Locations
    ws.merge_cells("J71:O71")
    ws["J71"] = "TOP 5 PICKUP LOCATIONS BY BOOKING VOLUME"
    ws["J71"].font = Font(name="Segoe UI", size=9, bold=True, color=NAVY_HEADER)

    sub_hdr(ws["J72"], "Rank")
    ws.merge_cells("K72:M72")
    sub_hdr(ws["K72"], "Pickup Location")
    sub_hdr(ws["N72"], "Bookings")
    sub_hdr(ws["O72"], "Fleet Share")

    top5 = df['Pickup Location'].value_counts().head(5)
    total_fleet = len(df)
    for rank, (loc, cnt) in enumerate(top5.items(), start=1):
        row_n = 72 + rank
        cell_fmt(ws[f"J{row_n}"], f"#{rank}")
        ws.merge_cells(f"K{row_n}:M{row_n}")
        cell_fmt(ws[f"K{row_n}"], str(loc), align="left")
        cell_fmt(ws[f"N{row_n}"], int(cnt), fmt="#,##0")
        cell_fmt(ws[f"O{row_n}"], round(cnt / total_fleet, 4), fmt="0.00%")

def build_ref_tables(ws, df, test_df):
    def ref_hdr(cell, text):
        cell.value = text
        cell.font = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def ref_cell(cell, val, fmt=None, align="center"):
        cell.value = val
        cell.font = Font(name="Segoe UI", size=8.5, color=TEXT_DARK)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border = thin_border()
        if fmt:
            cell.number_format = fmt

    ws.merge_cells("B82:O82")
    ws["B82"] = "DASHBOARD DATA SOURCES & REFERENCE CALCULATIONS (CHART FEED TABLES)"
    ws["B82"].font = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)
    ws["B82"].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
    ws.row_dimensions[82].height = 22

    ref_hdr(ws["B84"], "Status")
    ref_hdr(ws["C84"], "Count")
    ref_cell(ws["B85"], "Successful", align="left")
    ref_cell(ws["C85"], int((df['Actual_YES_NO'] == 'YES').sum()), fmt="#,##0")
    ref_cell(ws["B86"], "Unsuccessful", align="left")
    ref_cell(ws["C86"], int((df['Actual_YES_NO'] == 'NO').sum()), fmt="#,##0")

    ref_hdr(ws["E84"], "Vehicle Type")
    ref_hdr(ws["F84"], "Total Bookings")
    for idx, (veh, count) in enumerate(df['Vehicle Type'].value_counts().items(), start=85):
        ref_cell(ws[f"E{idx}"], str(veh), align="left")
        ref_cell(ws[f"F{idx}"], int(count), fmt="#,##0")

    ref_hdr(ws["H84"], "Outcome")
    ref_hdr(ws["I84"], "Actual")
    ref_hdr(ws["J84"], "Predicted")
    ref_cell(ws["H85"], "YES")
    ref_cell(ws["I85"], int((test_df['Actual_YES_NO'] == 'YES').sum()), fmt="#,##0")
    ref_cell(ws["J85"], int((test_df['Predicted_YES_NO'] == 'YES').sum()), fmt="#,##0")
    ref_cell(ws["H86"], "NO")
    ref_cell(ws["I86"], int((test_df['Actual_YES_NO'] == 'NO').sum()), fmt="#,##0")
    ref_cell(ws["J86"], int((test_df['Predicted_YES_NO'] == 'NO').sum()), fmt="#,##0")

    ref_hdr(ws["L84"], "Pickup Location")
    ref_hdr(ws["M84"], "Bookings")
    for idx, (loc, cnt) in enumerate(df['Pickup Location'].value_counts().head(10).items(), start=85):
        ref_cell(ws[f"L{idx}"], str(loc), align="left")
        ref_cell(ws[f"M{idx}"], int(cnt), fmt="#,##0")

    ref_hdr(ws["B96"], "Hour")
    ref_hdr(ws["C96"], "Booking Volume")
    for idx, (hr, count) in enumerate(df['Hour'].value_counts().sort_index().items(), start=97):
        ref_cell(ws[f"B{idx}"], f"{hr:02d}:00")
        ref_cell(ws[f"C{idx}"], int(count), fmt="#,##0")

    ref_hdr(ws["E96"], "Vehicle Type")
    ref_hdr(ws["F96"], "Avg Value (INR)")
    for idx, (veh, val) in enumerate(df.groupby('Vehicle Type')['Booking_Value'].mean().items(), start=97):
        ref_cell(ws[f"E{idx}"], str(veh), align="left")
        ref_cell(ws[f"F{idx}"], round(float(val), 2), fmt="INR #,##0.00")

    ref_hdr(ws["H96"], "Vehicle Type")
    ref_hdr(ws["I96"], "Avg Distance (km)")
    for idx, (veh, dist) in enumerate(df.groupby('Vehicle Type')['Ride_Distance'].mean().items(), start=97):
        ref_cell(ws[f"H{idx}"], str(veh), align="left")
        ref_cell(ws[f"I{idx}"], round(float(dist), 2), fmt="0.00")

def build_predictions(wb, pred_df):
    ws = wb.create_sheet(title="Predictions")
    ws.views.sheetView[0].showGridLines = True

    pred_export = pd.DataFrame({
        'Booking ID': pred_df['Booking ID'].astype(str).str.replace('"', ''),
        'Vehicle Type': pred_df['Vehicle Type'],
        'Pickup Location': pred_df['Pickup Location'],
        'Drop Location': pred_df['Drop Location'],
        'Ride Distance': np.round(pred_df['Ride Distance'], 2),
        'Booking Value': np.round(pred_df['Booking Value'], 2),
        'Payment Method': pred_df['Payment Method'].fillna('N/A (Cancelled)'),
        'Actual Booking Status': pred_df['Actual Status'],
        'Actual Success': pred_df['Actual_YES_NO'],
        'Predicted Success': pred_df['Predicted_YES_NO'],
        'Prediction Probability': np.round(pred_df['Prediction Probability'], 4),
        'Correct/Incorrect': pred_df['Correct/Incorrect']
    })

    headers = list(pred_export.columns)
    ws.append(headers)
    ws.row_dimensions[1].height = 28

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    rows = pred_export.values.tolist()
    for row in rows:
        ws.append(row)

    total_rows = len(rows) + 1
    last_col = get_column_letter(len(headers))

    dist_col = headers.index('Ride Distance') + 1
    val_col  = headers.index('Booking Value') + 1
    prob_col = headers.index('Prediction Probability') + 1

    for r in range(2, min(total_rows + 1, 1002)):
        ws.cell(row=r, column=dist_col).number_format = "0.00"
        ws.cell(row=r, column=val_col).number_format  = "INR #,##0.00"
        ws.cell(row=r, column=prob_col).number_format = "0.0%"

    tab = Table(displayName="PredictionsTable", ref=f"A1:{last_col}{total_rows}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)

    corr_letter = get_column_letter(headers.index('Correct/Incorrect') + 1)
    green_rule = CellIsRule(operator='equal', formula=['"Correct"'],
                            fill=PatternFill(start_color=SUCCESS_BG, end_color=SUCCESS_BG, fill_type="solid"),
                            font=Font(name="Segoe UI", size=9, bold=True, color=SUCCESS_GRN))
    red_rule = CellIsRule(operator='equal', formula=['"Incorrect"'],
                          fill=PatternFill(start_color=DANGER_BG, end_color=DANGER_BG, fill_type="solid"),
                          font=Font(name="Segoe UI", size=9, bold=True, color=DANGER_RED))

    ws.conditional_formatting.add(f"{corr_letter}2:{corr_letter}{total_rows}", green_rule)
    ws.conditional_formatting.add(f"{corr_letter}2:{corr_letter}{total_rows}", red_rule)

    ws.freeze_panes = "A2"
    auto_fit_cols(ws, max_cols=len(headers), sample_rows=50, min_width=14)

def build_model_perf(wb, metrics):
    ws = wb.create_sheet(title="Model Performance")
    ws.views.sheetView[0].showGridLines = True

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 22
    ws.column_dimensions["I"].width = 15
    ws.column_dimensions["J"].width = 15
    ws.column_dimensions["K"].width = 15

    ws.merge_cells("B2:K2")
    ws["B2"] = "DECISION TREE CLASSIFIER - MODEL PERFORMANCE & EVALUATION"
    ws["B2"].font = Font(name="Segoe UI", size=16, bold=True, color=WHITE)
    ws["B2"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 32

    overview_cards = [
        ("ALGORITHM", "Decision Tree Classifier", "B", "C"),
        ("TRAINING SAMPLES", f"{metrics['train_count']:,}", "D", "D"),
        ("TESTING SAMPLES", f"{metrics['test_count']:,}", "E", "E"),
        ("ACCURACY", f"{metrics['accuracy']:.4%}", "G", "H"),
        ("PRECISION", f"{metrics['precision']:.4%}", "I", "I"),
        ("RECALL", f"{metrics['recall']:.4%}", "J", "J"),
        ("F1 SCORE", f"{metrics['f1_score']:.4%}", "K", "K")
    ]

    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 26

    for title, val, c_start, c_end in overview_cards:
        if c_start != c_end:
            ws.merge_cells(f"{c_start}4:{c_end}4")
            ws.merge_cells(f"{c_start}5:{c_end}5")

        ws[f"{c_start}4"].value = title
        ws[f"{c_start}4"].font = Font(name="Segoe UI", size=8, bold=True, color=TEXT_MUTED)
        ws[f"{c_start}4"].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
        ws[f"{c_start}4"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"{c_start}5"].value = val
        ws[f"{c_start}5"].font = Font(name="Segoe UI", size=11, bold=True, color=NAVY_HEADER)
        ws[f"{c_start}5"].fill = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
        ws[f"{c_start}5"].alignment = Alignment(horizontal="center", vertical="center")

        for c in [c_start, c_end]:
            ws[f"{c}4"].border = Border(top=Side(style="medium", color=STEEL_BLUE), left=Side(style="thin", color=BORDER_GRAY), right=Side(style="thin", color=BORDER_GRAY))
            ws[f"{c}5"].border = Border(bottom=Side(style="thin", color=BORDER_GRAY), left=Side(style="thin", color=BORDER_GRAY), right=Side(style="thin", color=BORDER_GRAY))

    ws.row_dimensions[7].height = 12

    # 1. Metrics Table
    ws.merge_cells("B8:E8")
    ws["B8"] = "1. CLASSIFICATION METRICS TABLE"
    ws["B8"].font = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)

    ws["B9"] = "Metric"
    ws["C9"] = "Score"
    ws.merge_cells("D9:E9")
    ws["D9"] = "Formula / Definition"
    for col in ["B", "C", "D"]:
        ws[f"{col}9"].font = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        ws[f"{col}9"].fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
        ws[f"{col}9"].alignment = Alignment(horizontal="center", vertical="center")

    metric_rows = [
        ("Accuracy", metrics['accuracy'], "(TP + TN) / Total Predictions"),
        ("Precision", metrics['precision'], "TP / (TP + FP)"),
        ("Recall", metrics['recall'], "TP / (TP + FN)"),
        ("F1 Score", metrics['f1_score'], "2 * (Precision * Recall) / (Precision + Recall)")
    ]

    for idx, (m_name, m_val, m_desc) in enumerate(metric_rows, start=10):
        ws[f"B{idx}"] = m_name
        ws[f"B{idx}"].font = Font(name="Segoe UI", size=9, bold=True, color=TEXT_DARK)
        ws[f"B{idx}"].border = thin_border()

        ws[f"C{idx}"] = m_val
        ws[f"C{idx}"].font = Font(name="Segoe UI", size=9.5, bold=True, color=NAVY_HEADER)
        ws[f"C{idx}"].number_format = "0.0000%"
        ws[f"C{idx}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"C{idx}"].border = thin_border()

        ws.merge_cells(f"D{idx}:E{idx}")
        ws[f"D{idx}"] = m_desc
        ws[f"D{idx}"].font = Font(name="Segoe UI", size=8.5, italic=True, color=TEXT_MUTED)
        ws[f"D{idx}"].border = thin_border()

    # 2. Confusion Matrix
    ws.merge_cells("B15:E15")
    ws["B15"] = "2. CONFUSION MATRIX (TEST DATA: 30,000 SAMPLES)"
    ws["B15"].font = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)

    cm = metrics['confusion_matrix']
    tn, fp = int(cm[0, 0]), int(cm[0, 1])
    fn, tp = int(cm[1, 0]), int(cm[1, 1])

    ws["B16"] = "Actual \\ Predicted"
    ws["C16"] = "Predicted NO"
    ws["D16"] = "Predicted YES"
    ws["E16"] = "Total Actual"
    for col in ["B", "C", "D", "E"]:
        ws[f"{col}16"].font = Font(name="Segoe UI", size=8.5, bold=True, color=WHITE)
        ws[f"{col}16"].fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
        ws[f"{col}16"].alignment = Alignment(horizontal="center", vertical="center")

    ws["B17"] = "Actual NO"
    ws["B17"].font = Font(name="Segoe UI", size=8.5, bold=True, color=WHITE)
    ws["B17"].fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
    ws["B17"].alignment = Alignment(horizontal="center", vertical="center")

    ws["C17"] = f"{tn:,} (TN)"
    ws["C17"].font = Font(name="Segoe UI", size=9, bold=True, color=SUCCESS_GRN)
    ws["C17"].fill = PatternFill(start_color=SUCCESS_BG, end_color=SUCCESS_BG, fill_type="solid")
    ws["C17"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C17"].border = thin_border()

    ws["D17"] = f"{fp:,} (FP)"
    ws["D17"].font = Font(name="Segoe UI", size=9, bold=True, color=DANGER_RED)
    ws["D17"].fill = PatternFill(start_color=DANGER_BG, end_color=DANGER_BG, fill_type="solid")
    ws["D17"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D17"].border = thin_border()

    ws["E17"] = f"{tn + fp:,}"
    ws["E17"].font = Font(name="Segoe UI", size=9, bold=True)
    ws["E17"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E17"].border = thin_border()

    ws["B18"] = "Actual YES"
    ws["B18"].font = Font(name="Segoe UI", size=8.5, bold=True, color=WHITE)
    ws["B18"].fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
    ws["B18"].alignment = Alignment(horizontal="center", vertical="center")

    ws["C18"] = f"{fn:,} (FN)"
    ws["C18"].font = Font(name="Segoe UI", size=9, bold=True, color=GOLD_TEXT)
    ws["C18"].fill = PatternFill(start_color=GOLD_BG, end_color=GOLD_BG, fill_type="solid")
    ws["C18"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C18"].border = thin_border()

    ws["D18"] = f"{tp:,} (TP)"
    ws["D18"].font = Font(name="Segoe UI", size=9, bold=True, color=SUCCESS_GRN)
    ws["D18"].fill = PatternFill(start_color=SUCCESS_BG, end_color=SUCCESS_BG, fill_type="solid")
    ws["D18"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D18"].border = thin_border()

    ws["E18"] = f"{fn + tp:,}"
    ws["E18"].font = Font(name="Segoe UI", size=9, bold=True)
    ws["E18"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E18"].border = thin_border()

    ws["B19"] = "Total Predicted"
    ws["B19"].font = Font(name="Segoe UI", size=8.5, bold=True)
    ws["B19"].border = thin_border()

    ws["C19"] = f"{tn + fn:,}"
    ws["C19"].font = Font(name="Segoe UI", size=9, bold=True)
    ws["C19"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C19"].border = thin_border()

    ws["D19"] = f"{fp + tp:,}"
    ws["D19"].font = Font(name="Segoe UI", size=9, bold=True)
    ws["D19"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D19"].border = thin_border()

    ws["E19"] = f"{tn + fp + fn + tp:,}"
    ws["E19"].font = Font(name="Segoe UI", size=9.5, bold=True, color=NAVY_HEADER)
    ws["E19"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E19"].border = thin_border()

    # 5. Model Specs
    ws.merge_cells("B22:E22")
    ws["B22"] = "5. MODEL ARCHITECTURE & HYPERPARAMETERS"
    ws["B22"].font = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)

    model_specs = [
        ("Model Class", "DecisionTreeClassifier (scikit-learn CART)"),
        ("Splitting Criterion", "Gini Impurity"),
        ("Max Tree Depth", f"{DT_MAX_DEPTH}"),
        ("Min Samples Leaf", f"{DT_MIN_SAMPLES_LEAF}"),
        ("Random State", f"{RANDOM_STATE}"),
        ("Dataset Split", f"80% Train ({metrics['train_count']:,}) | 20% Test ({metrics['test_count']:,})"),
        ("Target Variable", "Booking Status (YES = Completed, NO = All Other)"),
        ("Model File", "decision_tree_model.pkl")
    ]

    for idx, (param, val) in enumerate(model_specs, start=23):
        ws[f"B{idx}"] = param
        ws[f"B{idx}"].font = Font(name="Segoe UI", size=8.5, bold=True, color=TEXT_MUTED)
        ws[f"B{idx}"].border = thin_border()

        ws.merge_cells(f"C{idx}:E{idx}")
        ws[f"C{idx}"] = val
        ws[f"C{idx}"].font = Font(name="Segoe UI", size=8.5, color=TEXT_DARK)
        ws[f"C{idx}"].border = thin_border()

    # 3. Feature Importance Table
    ws.merge_cells("G8:K8")
    ws["G8"] = "3. FEATURE IMPORTANCE RANKING"
    ws["G8"].font = Font(name="Segoe UI", size=10, bold=True, color=NAVY_HEADER)

    ws["G9"] = "Rank"
    ws["H9"] = "Feature Name"
    ws["I9"] = "Importance"
    ws["J9"] = "Share (%)"
    ws["K9"] = "Cumulative (%)"

    for col in ["G", "H", "I", "J", "K"]:
        ws[f"{col}9"].font = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        ws[f"{col}9"].fill = PatternFill(start_color=STEEL_BLUE, end_color=STEEL_BLUE, fill_type="solid")
        ws[f"{col}9"].alignment = Alignment(horizontal="center", vertical="center")

    importances = metrics['importances']
    cum_share = 0.0
    for idx, row in importances.iterrows():
        r = 10 + idx
        cum_share += row['Importance_Pct']

        ws[f"G{r}"] = int(row['Rank'])
        ws[f"G{r}"].alignment = Alignment(horizontal="center")
        ws[f"G{r}"].border = thin_border()

        ws[f"H{r}"] = str(row['Feature'])
        ws[f"H{r}"].font = Font(name="Segoe UI", size=8.5, bold=True)
        ws[f"H{r}"].border = thin_border()

        ws[f"I{r}"] = round(float(row['Importance']), 5)
        ws[f"I{r}"].number_format = "0.00000"
        ws[f"I{r}"].alignment = Alignment(horizontal="right")
        ws[f"I{r}"].border = thin_border()

        ws[f"J{r}"] = round(float(row['Importance_Pct']), 4)
        ws[f"J{r}"].number_format = "0.0%"
        ws[f"J{r}"].alignment = Alignment(horizontal="right")
        ws[f"J{r}"].border = thin_border()

        ws[f"K{r}"] = round(float(cum_share), 4)
        ws[f"K{r}"].number_format = "0.0%"
        ws[f"K{r}"].alignment = Alignment(horizontal="right")
        ws[f"K{r}"].border = thin_border()

    # 4. Feature Chart
    bar_fi = BarChart()
    bar_fi.type = "bar"
    bar_fi.style = 10
    bar_fi.title = "Decision Tree Feature Importances (Gini Reduction)"
    bar_fi.x_axis.title = "Gini Importance Score"
    bar_fi.y_axis.title = "Input Features"
    bar_fi.width = 15.6
    bar_fi.height = 8.5
    bar_fi.legend = None

    bar_fi.add_data(Reference(ws, min_col=9, min_row=9, max_row=9 + len(importances)), titles_from_data=True)
    bar_fi.set_categories(Reference(ws, min_col=8, min_row=10, max_row=9 + len(importances)))
    ws.add_chart(bar_fi, "G22")

def build_data(wb, df):
    ws = wb.create_sheet(title="Data")
    ws.views.sheetView[0].showGridLines = True

    cols = [
        'Booking ID', 'Date', 'Time', 'Hour', 'DayName', 'Booking Status',
        'Actual_YES_NO', 'Vehicle Type', 'Pickup Location', 'Drop Location',
        'Avg_VTAT', 'Avg_CTAT', 'Booking_Value', 'Ride_Distance', 'Payment Method'
    ]
    hdrs = [
        'Booking ID', 'Date', 'Time', 'Hour', 'Day of Week', 'Booking Status',
        'Actual Success', 'Vehicle Type', 'Pickup Location', 'Drop Location',
        'Avg VTAT', 'Avg CTAT', 'Booking Value', 'Ride Distance', 'Payment Method'
    ]

    ws.append(hdrs)
    ws.row_dimensions[1].height = 26

    for col_idx in range(1, len(hdrs) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    clean_df = df[cols].copy()
    clean_df['Booking ID'] = clean_df['Booking ID'].astype(str).str.replace('"', '')
    clean_df['Payment Method'] = clean_df['Payment Method'].fillna('N/A (Cancelled)')
    clean_df['Avg_VTAT'] = np.round(clean_df['Avg_VTAT'], 2)
    clean_df['Avg_CTAT'] = np.round(clean_df['Avg_CTAT'], 2)
    clean_df['Booking_Value'] = np.round(clean_df['Booking_Value'], 2)
    clean_df['Ride_Distance'] = np.round(clean_df['Ride_Distance'], 2)

    rows = clean_df.values.tolist()
    for row in rows:
        ws.append(row)

    total_rows = len(rows) + 1
    last_col = get_column_letter(len(hdrs))

    val_col  = hdrs.index('Booking Value') + 1
    dist_col = hdrs.index('Ride Distance') + 1
    for r in range(2, min(total_rows + 1, 502)):
        ws.cell(row=r, column=val_col).number_format  = "INR #,##0.00"
        ws.cell(row=r, column=dist_col).number_format = "0.00"

    tab = Table(displayName="CleanedRideData", ref=f"A1:{last_col}{total_rows}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)

    ws.freeze_panes = "A2"
    auto_fit_cols(ws, max_cols=len(hdrs), sample_rows=50, min_width=13)

def generate_dashboard():
    t0 = time.time()
    print("Loading data & trained model...")
    df, X, y, encoders = load_and_preprocess_data(DATA_PATH)
    model = joblib.load(MODEL_PATH)

    indices = df.index.to_numpy()
    X_train, X_test, y_train, y_test, idx_train, idx_test = train_test_split(
        X, y, indices, test_size=TEST_SIZE, random_state=RANDOM_STATE, stratify=y
    )

    y_pred = model.predict(X_test)
    y_prob = model.predict_proba(X_test)[:, 1]

    acc = accuracy_score(y_test, y_pred)
    prec = precision_score(y_test, y_pred)
    rec = recall_score(y_test, y_pred)
    f1 = f1_score(y_test, y_pred)
    cm = confusion_matrix(y_test, y_pred)

    importances = pd.DataFrame({
        'Feature': [FEATURE_LABELS.get(f, f) for f in INPUT_FEATURES],
        'Raw_Feature': INPUT_FEATURES,
        'Importance': model.feature_importances_
    }).sort_values(by='Importance', ascending=False).reset_index(drop=True)
    importances['Rank'] = importances.index + 1
    importances['Importance_Pct'] = importances['Importance'] / importances['Importance'].sum()

    metrics = {
        'accuracy': acc, 'precision': prec, 'recall': rec, 'f1_score': f1,
        'confusion_matrix': cm, 'importances': importances,
        'train_count': len(X_train), 'test_count': len(X_test), 'total_count': len(df)
    }

    test_df = df.loc[idx_test].copy().reset_index(drop=True)
    test_df['Predicted_Binary'] = y_pred
    test_df['Predicted_YES_NO'] = np.where(y_pred == 1, 'YES', 'NO')
    test_df['Prediction_Probability'] = np.round(y_prob, 4)
    test_df['Correct_Incorrect'] = np.where(test_df['Actual_YES_NO'] == test_df['Predicted_YES_NO'], 'Correct', 'Incorrect')

    pred_df = pd.DataFrame({
        'Booking ID': test_df['Booking ID'].astype(str).str.replace('"', ''),
        'Vehicle Type': test_df['Vehicle Type'],
        'Pickup Location': test_df['Pickup Location'],
        'Drop Location': test_df['Drop Location'],
        'Ride Distance': np.round(test_df['Ride_Distance'], 2),
        'Booking Value': np.round(test_df['Booking_Value'], 2),
        'Payment Method': test_df['Payment Method'],
        'Actual Status': test_df['Actual_Status'],
        'Actual_YES_NO': test_df['Actual_YES_NO'],
        'Predicted_YES_NO': test_df['Predicted_YES_NO'],
        'Prediction Probability': test_df['Prediction_Probability'],
        'Correct/Incorrect': test_df['Correct_Incorrect']
    })

    print("Building Excel Workbook...")
    wb = openpyxl.Workbook()
    default_ws = wb.active

    build_dashboard(wb, df, metrics, test_df)
    build_predictions(wb, pred_df)
    build_model_perf(wb, metrics)
    build_data(wb, df)

    wb.remove(default_ws)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"Excel dashboard successfully saved to: {OUTPUT_PATH}")
    print(f"Process completed in {time.time()-t0:.2f} seconds.")

if __name__ == "__main__":
    generate_dashboard()
