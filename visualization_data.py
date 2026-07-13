"""
Project 4: Data Visualization (Optional Milestone)
DecodeLabs Data Analytics Internship
Author: Shafi
Dataset: Share_Dataset_for_Data_Analytics.xlsx (1200 records, Jan 2023 - Jun 2025)

This script pulls the three numbers used in the Project4_DataVisualization.pptx
deck: revenue by product, the monthly revenue trend, and the order status
breakdown. The actual chart design decisions (chart type, zero-baseline axis,
direct labeling, single accent color, no gridlines/chartjunk) live in the deck
itself - this is just the data behind it, kept here so the numbers are easy
to check or refresh from an updated dataset.
"""

import pandas as pd

DATA_PATH = "orders_data.xlsx"

df = pd.read_excel(DATA_PATH)
df["Month"] = df["Date"].dt.to_period("M").astype(str)


def revenue_by_product():
    """Bar chart data - comparing categories, so a column chart is correct."""
    return df.groupby("Product")["TotalPrice"].sum().sort_values(ascending=False)


def monthly_revenue_trend():
    """Line chart data - tracking a continuous value over time."""
    return df.groupby("Month")["TotalPrice"].sum()


def order_status_breakdown():
    """Bar chart data - highlights Cancelled/Pending as the actionable slice."""
    return df["OrderStatus"].value_counts()


def main():
    print("=== Revenue by Product (Slide 2) ===")
    print(revenue_by_product().round(0))

    print("\n=== Monthly Revenue Trend (Slide 3) ===")
    monthly = monthly_revenue_trend()
    print(f"Average: ${monthly.mean():,.2f} | Std Dev: ${monthly.std():,.2f}")
    print(f"Peak: {monthly.idxmax()} (${monthly.max():,.2f})")
    print(f"Low:  {monthly.idxmin()} (${monthly.min():,.2f})")

    print("\n=== Order Status Breakdown (Slide 4) ===")
    status = order_status_breakdown()
    print(status)
    flagged = status.get("Cancelled", 0) + status.get("Pending", 0)
    print(f"Cancelled + Pending: {flagged} orders ({flagged / len(df) * 100:.1f}% of total)")


if __name__ == "__main__":
    main()
