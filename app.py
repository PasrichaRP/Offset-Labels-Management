# app.py
# Install required packages:
# pip install streamlit pandas numpy openpyxl xlsxwriter

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from collections import defaultdict
import copy

st.set_page_config(page_title="Offset Printing Layout Optimizer", layout="wide")
st.title("📄 Offset Printing Layout Sheet Manager")
st.markdown("Optimize label layouts by merging multiple designs onto shared printing sheets to minimize plate sets.")

# ============================================
# Helper Functions
# ============================================

def load_excel(file):
    """Load Excel and extract relevant columns."""
    df = pd.read_excel(file)
    # Find relevant columns
    item_col = None
    qty_col = None
    for col in df.columns:
        if "ITEM" in str(col).upper() or "CODE" in str(col).upper():
            item_col = col
        if "QUANTITY REQUIRED" in str(col) or "REQUIRED" in str(col).upper():
            qty_col = col
    if item_col is None or qty_col is None:
        st.error("Could not find ITEM CODE and QUANTITY REQUIRED columns. Please check file format.")
        return None
    df = df[[item_col, qty_col]].copy()
    df.columns = ["ITEM CODE", "QUANTITY REQUIRED"]
    df["QUANTITY REQUIRED"] = pd.to_numeric(df["QUANTITY REQUIRED"], errors="coerce")
    df = df.dropna().reset_index(drop=True)
    df["ITEM CODE"] = df["ITEM CODE"].astype(str)
    return df

def calculate_set_parameters(group_items, capacity, sheets=None):
    """
    For a list of items with required quantities, find optimal sheets count (S)
    and pcs per item such that:
        sum(pcs_i) <= capacity
        printed_i = S * pcs_i >= required_i
    Minimizes extra and uses smallest S that satisfies constraints.
    Returns (S, pcs_dict, printed_dict, extra_dict, total_extra, utilization)
    """
    if not group_items:
        return None, {}, {}, {}, 0, 0
    
    reqs = {item["code"]: item["required"] for item in group_items}
    
    # Binary search for minimal S
    low = 1
    # Upper bound: max required (if pcs=1) or capacity scenario
    high = max(reqs.values())  # worst case sheets if pcs=1
    
    # Try to find feasible S
    best_S = None
    best_pcs = {}
    best_extra = float('inf')
    
    for S in range(low, high + 1):
        total_pcs = 0
        pcs_candidate = {}
        feasible = True
        for code, req in reqs.items():
            pcs = (req + S - 1) // S  # ceil division
            if pcs > capacity:
                feasible = False
                break
            pcs_candidate[code] = pcs
            total_pcs += pcs
        
        if feasible and total_pcs <= capacity:
            # Check extra total
            total_extra = sum((S * pcs_candidate[code] - req) for code, req in reqs.items())
            if total_extra < best_extra:
                best_extra = total_extra
                best_S = S
                best_pcs = pcs_candidate
    
    # If nothing found, increase S gradually until bound found
    if best_S is None:
        # Use S = max(required) * 2 or something
        for S in range(high, high * 5):
            total_pcs = 0
            pcs_candidate = {}
            feasible = True
            for code, req in reqs.items():
                pcs = (req + S - 1) // S
                if pcs > capacity:
                    feasible = False
                    break
                pcs_candidate[code] = pcs
                total_pcs += pcs
            if feasible and total_pcs <= capacity:
                best_S = S
                best_pcs = pcs_candidate
                best_extra = sum((S * best_pcs[code] - req) for code, req in reqs.items())
                break
    
    if best_S is None:
        # Fallback: each item in its own set with pcs=capacity, but we force here
        best_S = max((req + capacity - 1) // capacity for req in reqs.values())
        for code, req in reqs.items():
            best_pcs[code] = capacity
        
    printed = {code: best_S * best_pcs[code] for code in reqs}
    extra = {code: printed[code] - reqs[code] for code in reqs}
    total_extra = sum(extra.values())
    used_positions = sum(best_pcs.values())
    utilization = (used_positions / capacity) * 100 if capacity > 0 else 0
    
    return best_S, best_pcs, printed, extra, total_extra, utilization

def auto_optimize_sets(items_df, capacity):
    """
    Greedy merging: group items that can share a common sheet count S.
    Each set uses the same S for all items.
    """
    items = [{"code": row["ITEM CODE"], "required": row["QUANTITY REQUIRED"]} 
             for _, row in items_df.iterrows()]
    # Sort by required quantity descending for better packing
    items.sort(key=lambda x: x["required"], reverse=True)
    
    unassigned = items.copy()
    sets = []
    
    while unassigned:
        current_group = []
        # Try to add items greedily
        for item in unassigned[:]:
            # Test if this item can join current group
            test_group = current_group + [item]
            S, pcs, printed, extra, _, _ = calculate_set_parameters(test_group, capacity)
            if S is not None:
                current_group.append(item)
                unassigned.remove(item)
        if not current_group:
            # No item could be added - take largest remaining alone
            current_group = [unassigned.pop(0)]
        # Finalize group
        S, pcs_dict, printed_dict, extra_dict, total_extra, util = calculate_set_parameters(current_group, capacity)
        set_info = {
            "set_id": len(sets) + 1,
            "sheets": S,
            "capacity": capacity,
            "utilization": util,
            "items": [],
            "total_extra": total_extra
        }
        for item in current_group:
            code = item["code"]
            set_info["items"].append({
                "code": code,
                "required": item["required"],
                "pcs": pcs_dict[code],
                "printed": printed_dict[code],
                "extra": extra_dict[code],
                "extra_percent": (extra_dict[code] / item["required"]) * 100 if item["required"] > 0 else 0
            })
        sets.append(set_info)
    
    return sets

def recalc_from_manual(items_manual, capacity):
    """
    items_manual: list of dict with code, required, set_id, pcs, sheets (per set)
    Group by set_id and recalc.
    """
    # Group by set_id
    groups = defaultdict(list)
    for item in items_manual:
        groups[item["set_id"]].append(item)
    
    new_sets = []
    for set_id, group_items in groups.items():
        sheets = group_items[0]["sheets"]  # all share same sheets from manual
        # Verify consistency
        total_pcs = sum(item["pcs"] for item in group_items)
        if total_pcs > capacity:
            st.warning(f"Set {set_id} exceeds capacity ({total_pcs} > {capacity}). Please adjust pcs values.")
        # Compute printed, extra
        items_out = []
        total_extra = 0
        for item in group_items:
            printed = sheets * item["pcs"]
            extra = printed - item["required"]
            extra_pct = (extra / item["required"]) * 100 if item["required"] > 0 else 0
            items_out.append({
                "code": item["code"],
                "required": item["required"],
                "pcs": item["pcs"],
                "printed": printed,
                "extra": extra,
                "extra_percent": extra_pct
            })
            total_extra += extra
        utilization = (sum(item["pcs"] for item in group_items) / capacity) * 100
        new_sets.append({
            "set_id": set_id,
            "sheets": sheets,
            "capacity": capacity,
            "utilization": utilization,
            "items": items_out,
            "total_extra": total_extra
        })
    return new_sets

def sets_to_dataframe(sets):
    """Convert sets structure to editable dataframe rows."""
    rows = []
    for s in sets:
        for item in s["items"]:
            rows.append({
                "ITEM CODE": item["code"],
                "QUANTITY REQUIRED": item["required"],
                "SET ID": s["set_id"],
                "PCS per Sheet": item["pcs"],
                "Set Sheets": s["sheets"],
                "Printed Qty": item["printed"],
                "Extra Qty": item["extra"],
                "Extra %": round(item["extra_percent"], 1)
            })
    return pd.DataFrame(rows)

def apply_edits(df_edits, original_items_df):
    """Convert edited dataframe back to sets structure."""
    # Group by SET ID
    sets_dict = {}
    for _, row in df_edits.iterrows():
        set_id = int(row["SET ID"])
        if set_id not in sets_dict:
            sets_dict[set_id] = {"sheets": int(row["Set Sheets"]), "items": []}
        sets_dict[set_id]["items"].append({
            "code": row["ITEM CODE"],
            "required": row["QUANTITY REQUIRED"],
            "pcs": int(row["PCS per Sheet"])
        })
    # Rebuild sets with same sheets
    new_sets = []
    for set_id, data in sets_dict.items():
        sheets = data["sheets"]
        total_pcs = sum(item["pcs"] for item in data["items"])
        items_out = []
        total_extra = 0
        for item in data["items"]:
            printed = sheets * item["pcs"]
            extra = printed - item["required"]
            extra_pct = (extra / item["required"]) * 100 if item["required"] > 0 else 0
            items_out.append({
                "code": item["code"],
                "required": item["required"],
                "pcs": item["pcs"],
                "printed": printed,
                "extra": extra,
                "extra_percent": extra_pct
            })
            total_extra += extra
        capacity = st.session_state.get("capacity", 12)  # get current capacity
        utilization = (sum(item["pcs"] for item in data["items"]) / capacity) * 100 if capacity else 0
        new_sets.append({
            "set_id": set_id,
            "sheets": sheets,
            "capacity": capacity,
            "utilization": utilization,
            "items": items_out,
            "total_extra": total_extra
        })
    return new_sets

def export_to_excel(sets):
    """Export sets and items to Excel with multiple sheets."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Summary sheet
        summary = []
        for s in sets:
            summary.append({
                "Set No.": s["set_id"],
                "Total Sheets": s["sheets"],
                "Utilization %": round(s["utilization"], 1),
                "Total Extra Qty": s["total_extra"],
                "Labels in Set": len(s["items"])
            })
        pd.DataFrame(summary).to_excel(writer, sheet_name="Sets Summary", index=False)
        
        # Detailed items
        items_data = []
        for s in sets:
            for item in s["items"]:
                items_data.append({
                    "Set No.": s["set_id"],
                    "ITEM CODE": item["code"],
                    "Required": item["required"],
                    "Pcs/Sheet": item["pcs"],
                    "Sheets": s["sheets"],
                    "Printed Qty": item["printed"],
                    "Extra Qty": item["extra"],
                    "Extra %": item["extra_percent"]
                })
        pd.DataFrame(items_data).to_excel(writer, sheet_name="Detailed Layout", index=False)
        
        # Production sheet
        prod = []
        for s in sets:
            prod.append({
                "Set No.": s["set_id"],
                "Print Sheets Required": s["sheets"],
                "Total Positions per Sheet": s["capacity"],
                "Used Positions": sum(item["pcs"] for item in s["items"]),
                "Waste %": round(100 - s["utilization"], 1)
            })
        pd.DataFrame(prod).to_excel(writer, sheet_name="Production Sheet", index=False)
    output.seek(0)
    return output

# ============================================
# Main App
# ============================================

# Initialize session state
if "items_df" not in st.session_state:
    st.session_state.items_df = None
if "sets" not in st.session_state:
    st.session_state.sets = None
if "capacity" not in st.session_state:
    st.session_state.capacity = 12
if "warning_threshold" not in st.session_state:
    st.session_state.warning_threshold = 50

# Sidebar controls
with st.sidebar:
    st.header("⚙️ Settings")
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
    st.session_state.capacity = st.number_input("Labels per Sheet (Capacity)", min_value=1, value=12, step=1)
    st.session_state.warning_threshold = st.slider("Extra Quantity Warning Threshold (%)", 0, 100, 50)
    
    if st.button("🔄 Auto-Optimize Sets", type="primary"):
        if st.session_state.items_df is not None:
            with st.spinner("Optimizing..."):
                st.session_state.sets = auto_optimize_sets(st.session_state.items_df, st.session_state.capacity)
            st.success("Optimization complete!")
        else:
            st.warning("Please upload an Excel file first.")
    
    if st.session_state.sets is not None:
        st.download_button(
            label="📥 Export to Excel",
            data=export_to_excel(st.session_state.sets),
            file_name="printing_layout_optimized.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            label="📄 Export Production Sheet (Excel)",
            data=export_to_excel(st.session_state.sets),
            file_name="production_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Main area
if uploaded_file:
    df = load_excel(uploaded_file)
    if df is not None:
        st.session_state.items_df = df
        st.success(f"Loaded {len(df)} label designs.")
        
        # Show raw data
        with st.expander("📋 Uploaded Data"):
            st.dataframe(df, use_container_width=True)
        
        # Auto-optimize if no sets yet
        if st.session_state.sets is None:
            st.session_state.sets = auto_optimize_sets(st.session_state.items_df, st.session_state.capacity)

# Display editable sets
if st.session_state.sets is not None and st.session_state.items_df is not None:
    st.markdown("## 📊 Optimized Layout Sets")
    
    # Convert to editable dataframe
    edit_df = sets_to_dataframe(st.session_state.sets)
    
    # Show warning for high extra
    high_extra = edit_df[edit_df["Extra %"] > st.session_state.warning_threshold]
    if not high_extra.empty:
        st.warning(f"⚠️ {len(high_extra)} label(s) have extra quantity > {st.session_state.warning_threshold}%. Consider adjusting pcs or creating a new set.")
        st.dataframe(high_extra[["ITEM CODE", "Required Qty", "Printed Qty", "Extra %"]], use_container_width=True)
    
    st.markdown("### ✏️ Manual Editing - Adjust Set ID, PCS per Sheet, Set Sheets")
    st.info("Modify the table below (Set ID, PCS per Sheet, Set Sheets) and changes apply automatically. Ensure sum of PCS per Sheet ≤ Labels per Sheet capacity.")
    
    edited_df = st.data_editor(
        edit_df,
        use_container_width=True,
        column_config={
            "ITEM CODE": st.column_config.TextColumn("Item Code", disabled=True),
            "QUANTITY REQUIRED": st.column_config.NumberColumn("Required Qty", disabled=True),
            "SET ID": st.column_config.NumberColumn("Set ID", step=1),
            "PCS per Sheet": st.column_config.NumberColumn("PCS per Sheet", step=1, min_value=1),
            "Set Sheets": st.column_config.NumberColumn("Sheets per Set", step=1, min_value=1),
            "Printed Qty": st.column_config.NumberColumn("Printed Qty", disabled=True),
            "Extra Qty": st.column_config.NumberColumn("Extra Qty", disabled=True),
            "Extra %": st.column_config.NumberColumn("Extra %", disabled=True),
        },
        num_rows="dynamic"
    )
    
    # Apply edits if changed
    if not edited_df.equals(edit_df):
        with st.spinner("Recalculating..."):
            st.session_state.sets = apply_edits(edited_df, st.session_state.items_df)
        st.rerun()
    
    # Sets summary
    st.markdown("### 📈 Sets Summary")
    summary_data = []
    for s in st.session_state.sets:
        summary_data.append({
            "Set No.": s["set_id"],
            "Sheets": s["sheets"],
            "Positions Used": sum(item["pcs"] for item in s["items"]),
            "Capacity": s["capacity"],
            "Utilization %": round(s["utilization"], 1),
            "Total Extra": s["total_extra"]
        })
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True)
    
    # Visualization: Utilization bar chart
    if len(summary_df) > 0:
        st.markdown("### 📊 Sheet Utilization per Set")
        st.bar_chart(summary_df.set_index("Set No.")["Utilization %"])
    
    # Individual set details
    for s in st.session_state.sets:
        with st.expander(f"Set #{s['set_id']} - {s['sheets']} sheets, Utilization {round(s['utilization'],1)}%"):
            items_data = []
            for item in s["items"]:
                items_data.append({
                    "Item Code": item["code"],
                    "Required": item["required"],
                    "Pcs/Sheet": item["pcs"],
                    "Printed": item["printed"],
                    "Extra": item["extra"],
                    "Extra %": f"{item['extra_percent']:.1f}%"
                })
            st.table(pd.DataFrame(items_data))
            
            # Waste warning per set
            if s["utilization"] < 70:
                st.warning(f"⚠️ Set #{s['set_id']} utilization below 70%. Consider rearranging.")
    
    # Production sheet printable view
    st.markdown("## 🖨️ Printable Production Sheet")
    prod_data = []
    for s in st.session_state.sets:
        prod_data.append({
            "Set Number": s["set_id"],
            "Print Sheets": s["sheets"],
            "Sheet Capacity": s["capacity"],
            "Used Positions": sum(item["pcs"] for item in s["items"]),
            "Waste Positions": s["capacity"] - sum(item["pcs"] for item in s["items"]),
            "Job Notes": f"{len(s['items'])} label types"
        })
    prod_df = pd.DataFrame(prod_data)
    st.dataframe(prod_df, use_container_width=True)
    
    # Manual grouping suggestion: new set button
    if st.button("➕ Suggest New Set for high waste items"):
        st.info("Please manually assign a new Set ID in the table above for items with high extra %.")

else:
    st.info("👈 Upload an Excel file to begin. The system will automatically merge labels onto shared sheets to minimize plate sets.")

# Footer
st.markdown("---")
st.caption("Offset Printing Layout Optimizer | Minimizes plate sets by intelligently merging labels per sheet.")
