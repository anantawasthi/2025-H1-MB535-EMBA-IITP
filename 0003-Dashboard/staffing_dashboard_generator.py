import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.gridspec as gridspec

# Load the dataset
data = pd.read_excel("Staffing_Dashboard_Data.xlsx")

# Prepare summary KPIs
summary = {
    "Average Time to Hire": f"{data['Time_to_Hire_Days'].mean():.1f} days",
    "Offer Acceptance Rate": f"{(data['Offer_Accepted'].mean() * 100):.1f}%",
    "Average Cost per Hire": f"₹{data['Cost_per_Hire'].mean():,.0f}",
    "First-Year Attrition Rate": f"{(data['First_Year_Attrition'].mean() * 100):.1f}%",
}

# Create dashboard layout
plt.figure(figsize=(18, 10))
gs = gridspec.GridSpec(4, 3)

# KPI Tiles
kpi_positions = [(0, 0), (0, 1), (0, 2), (1, 0)]
for pos, (kpi, value) in zip(kpi_positions, summary.items()):
    ax = plt.subplot(gs[pos])
    ax.text(0.5, 0.6, value, fontsize=22, fontweight="bold", ha="center", va="center")
    ax.text(0.5, 0.2, kpi, fontsize=14, ha="center", va="center")
    ax.axis("off")
    ax.set_facecolor("#f0f0f0")

# Bar Chart: Time to Hire by Department
ax1 = plt.subplot(gs[1, 1:])
sns.barplot(
    data=data, x="Department", y="Time_to_Hire_Days", ci=None, ax=ax1, palette="Blues"
)
ax1.set_title("Average Time to Hire by Department")
ax1.set_ylabel("Days")
ax1.set_xlabel("")

# Pie Chart: Source of Hire
ax2 = plt.subplot(gs[2, 0])
source_counts = data["Source_of_Hire"].value_counts()
ax2.pie(source_counts, labels=source_counts.index, autopct="%1.0f%%", startangle=140)
ax2.set_title("Source of Hire Distribution")

# Line Chart: First-Year Attrition by Quarter
ax3 = plt.subplot(gs[2, 1])
attrition_by_quarter = data.groupby("Quarter")["First_Year_Attrition"].mean() * 100
attrition_by_quarter = attrition_by_quarter.sort_index()
ax3.plot(
    attrition_by_quarter.index,
    attrition_by_quarter.values,
    marker="o",
    linestyle="-",
    color="red",
)
ax3.set_title("First-Year Attrition Rate by Quarter")
ax3.set_ylabel("Attrition Rate (%)")
ax3.set_ylim(0, max(attrition_by_quarter.values) + 10)

# Bar Chart: Offer Acceptance Rate by Department
ax4 = plt.subplot(gs[2, 2])
acceptance_by_dept = data.groupby("Department")["Offer_Accepted"].mean() * 100
sns.barplot(
    x=acceptance_by_dept.index, y=acceptance_by_dept.values, ax=ax4, palette="Greens"
)
ax4.set_title("Offer Acceptance Rate by Department")
ax4.set_ylabel("Acceptance Rate (%)")
ax4.set_xlabel("")

# Final Layout and Save
plt.tight_layout()
plt.savefig("ABC_Tech_Staffing_Dashboard.png", dpi=300)
plt.show()
