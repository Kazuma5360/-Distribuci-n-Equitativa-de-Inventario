import pandas as pd
import numpy as np

# ========== CONFIGURATION ==========
# Update these paths before running
DEMAND_FILE = "demands.xlsx"       # Your demand/orders file
DEMAND_SHEET = "Sheet1"            # Sheet name
DEMAND_HEADER_ROW = 2              # Row where headers start (0-indexed)

INVENTORY_FILE = "inventory.xls"   # Your inventory file

OUTPUT_FILE = "Equitable_Distribution_Result.xlsx"

# ========== COLUMNS TO DROP FROM DEMAND FILE ==========
# Adjust these to match your actual file columns
COLUMNS_TO_DROP = [
    "KEY", "SUB_FAMILY", "FAMILY_CODE",
    "FAMILY_DESCRIPTION", "WHOLESALE", "ORDERS",
    "SUPPLIED_QTY", "REQUESTED_QTY"
]

# ========== LOAD DEMAND FILE ==========
df_demand = pd.read_excel(DEMAND_FILE, sheet_name=DEMAND_SHEET, header=DEMAND_HEADER_ROW)

# Clean and prepare demand DataFrame
df_demand = df_demand.drop(columns=COLUMNS_TO_DROP, errors="ignore")
df_demand["ORDER"] = np.nan
df_demand = df_demand.sort_values(by=["SKU", "PENDING_QTY"], ascending=[True, True])

# Keep only FRACTIONED type orders
df_demand = df_demand[df_demand["TYPE"] == "FRACTIONED"]

# Keep only rows with PENDING_QTY > 0
df_demand = df_demand[df_demand["PENDING_QTY"] != 0]
df_demand = df_demand.dropna(subset=["PENDING_QTY"])

# ========== LOAD INVENTORY FILE ==========
df_inventory = pd.read_excel(INVENTORY_FILE)
df_inventory = df_inventory.rename(columns={"PRODUCT_ID": "SKU"})

# Find the lot with the HIGHEST available stock per SKU
idx_max_available = df_inventory.groupby("SKU")["AVAILABLE"].idxmax()
df_main_lot = df_inventory.loc[
    idx_max_available,
    ["SKU", "LOT_ID", "AVAILABLE", "WAREHOUSE_ID", "EXEMPT"]
].rename(columns={
    "LOT_ID": "MAIN_LOT_ID",
    "AVAILABLE": "LOT_AVAILABLE",
    "WAREHOUSE_ID": "WAREHOUSE",
    "EXEMPT": "EXEMPT"
})

# Sum TOTAL available stock per SKU (all lots)
df_total_available = (
    df_inventory
    .groupby("SKU", as_index=False)["AVAILABLE"]
    .sum()
    .rename(columns={"AVAILABLE": "TOTAL_AVAILABLE"})
)

# Merge total available + main lot info
df_inventory = df_total_available.merge(df_main_lot, on="SKU", how="left")
df_inventory = df_inventory.rename(columns={"TOTAL_AVAILABLE": "AVAILABLE"})

# ========== MERGE DEMAND + INVENTORY ==========
df = pd.merge(df_demand, df_inventory, on="SKU", how="left")
df = df.dropna(subset=["AVAILABLE"])
df = df[df["AVAILABLE"] != 0]


# ========== EQUITABLE DISTRIBUTION FUNCTION ==========
def distribute_equitably(group):
    """
    Distributes available inventory proportionally across all orders for a given SKU.

    Each order receives a share based on what it requested relative to the total demand.
    Remainders from floor rounding are assigned to orders with the highest decimal residuals.
    Any final leftover (after rounding) is given to the last order without exceeding its request.

    Args:
        group (pd.DataFrame): Subset of rows sharing the same SKU.

    Returns:
        pd.DataFrame: Group with added columns:
            - DISTRIBUTED: Units assigned to each order.
            - PCT_SERVED: Percentage of the order fulfilled.
            - SHORTAGE: Units still pending after distribution.
            - LEFTOVER_STOCK: Remaining inventory after all assignments.
    """
    available = group["AVAILABLE"].iloc[0]
    pending = group["PENDING_QTY"].values
    total_requested = pending.sum()

    if available >= total_requested:
        # Enough stock for everyone
        group["DISTRIBUTED"] = pending
        group["PCT_SERVED"] = 100.0
    else:
        # Proportional share
        ratio = available / total_requested
        assigned = np.floor(pending * ratio).astype(int)

        # Distribute rounding remainders to orders with highest decimal residuals
        leftover_units = int(available - assigned.sum())
        if leftover_units > 0:
            residuals = (pending * ratio) - assigned
            sorted_indices = np.argsort(-residuals)
            for i in range(leftover_units):
                assigned[sorted_indices[i]] += 1

        # Assign any final surplus to the last order (without exceeding its request)
        final_surplus = int(available - assigned.sum())
        if final_surplus > 0:
            last_idx = len(assigned) - 1
            can_receive = int(pending[last_idx] - assigned[last_idx])
            assigned[last_idx] += min(final_surplus, can_receive)

        group["DISTRIBUTED"] = assigned
        group["PCT_SERVED"] = np.round((assigned / pending) * 100, 2)

    group["SHORTAGE"] = group["PENDING_QTY"] - group["DISTRIBUTED"]
    group["LEFTOVER_STOCK"] = available - group["DISTRIBUTED"].sum()

    return group


# ========== APPLY DISTRIBUTION ==========
df = df.groupby("SKU", group_keys=False).apply(distribute_equitably)


# ========== PRINT SUMMARY BY SKU ==========
print("\n" + "=" * 100)
print("SUMMARY BY SKU:")
print("=" * 100)
summary = df.groupby("SKU").agg({
    "AVAILABLE":    "first",
    "PENDING_QTY":  "sum",
    "DISTRIBUTED":  "sum",
    "SHORTAGE":     "sum",
    "PCT_SERVED":   "first",
    "LEFTOVER_STOCK": "first"
}).round(2)
print(summary.head(20))


# ========== PRINT GENERAL STATISTICS ==========
print("\n" + "=" * 100)
print("GENERAL STATISTICS:")
print("=" * 100)
print(f"Total orders:              {len(df)}")
print(f"Unique SKUs:               {df['SKU'].nunique()}")
print(f"Total pending requested:   {df['PENDING_QTY'].sum():.0f}")
print(f"Total available inventory: {df.groupby('SKU')['AVAILABLE'].first().sum():.0f}")
print(f"Total distributed:         {df['DISTRIBUTED'].sum():.0f}")
print(f"Total shortage:            {df['SHORTAGE'].sum():.0f}")
print(f"Average % served:          {df.groupby('SKU')['PCT_SERVED'].first().mean():.2f}%")


# ========== EXPORT TO EXCEL ==========
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    # Sheet 1: Full distribution detail
    df.to_excel(writer, sheet_name="Full Distribution", index=False)

    # Sheet 2: Summary by SKU
    summary.to_excel(writer, sheet_name="Summary by SKU")

    # Sheet 3: Only orders with shortage
    df_shortage = df[df["SHORTAGE"] > 0].copy()
    df_shortage.to_excel(writer, sheet_name="Orders with Shortage", index=False)

    # Sheet 4: General statistics
    stats = pd.DataFrame({
        "Metric": [
            "Total orders",
            "Unique SKUs",
            "Total pending requested",
            "Total available inventory",
            "Total distributed",
            "Total shortage",
            "Average % served"
        ],
        "Value": [
            len(df),
            df["SKU"].nunique(),
            df["PENDING_QTY"].sum(),
            df.groupby("SKU")["AVAILABLE"].first().sum(),
            df["DISTRIBUTED"].sum(),
            df["SHORTAGE"].sum(),
            f"{df.groupby('SKU')['PCT_SERVED'].first().mean():.2f}%"
        ]
    })
    stats.to_excel(writer, sheet_name="General Statistics", index=False)

print("\n" + "=" * 100)
print(f"✅ EXCEL FILE GENERATED: {OUTPUT_FILE}")
print("=" * 100)
print("\nThe file contains 4 sheets:")
print("  1. Full Distribution  — All orders with their assigned quantities")
print("  2. Summary by SKU     — Totals grouped by product")
print("  3. Orders with Shortage — Only orders not fully covered")
print("  4. General Statistics — Overall summary")
print("=" * 100)
