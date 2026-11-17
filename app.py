
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="Sales & Revenue Dashboard",
    page_icon="📊",
    layout="wide"
)

# --------------------------------------------------
# CUSTOM CORPORATE DESIGN
# --------------------------------------------------

st.markdown("""
<style>

/* MAIN PAGE */
.stApp {
    background-color: #F4F7FB;
    color: #172B4D;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

/* LIGHT BLUE SIDEBAR */
[data-testid="stSidebar"] {
    background-color: #EAF2FA !important;
    border-right: 1px solid #D3E1EF;
}

/* SIDEBAR TEXT */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    color: #173B5E !important;
}

/* SIDEBAR DROPDOWN BOXES */
[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border: 1px solid #B8CCE0 !important;
    border-radius: 8px !important;
    color: #173B5E !important;
}

/* DROPDOWN TEXT */
[data-testid="stSidebar"] [data-baseweb="select"] span,
[data-testid="stSidebar"] [data-baseweb="select"] input {
    color: #173B5E !important;
    -webkit-text-fill-color: #173B5E !important;
}

/* SELECTED FILTER TAGS */
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: #DCEBFA !important;
    border: 1px solid #A9C5E0 !important;
    border-radius: 6px !important;
}

/* SELECTED FILTER TAG TEXT */
[data-testid="stSidebar"] [data-baseweb="tag"] span {
    color: #173B5E !important;
}

/* FILTER REMOVE ICON */
[data-testid="stSidebar"] [data-baseweb="tag"] svg {
    fill: #2463A6 !important;
    color: #2463A6 !important;
}

/* DROPDOWN ARROWS */
[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: #2463A6 !important;
    color: #2463A6 !important;
}

/* DATE INPUT BOXES */
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    background-color: #FFFFFF !important;
    color: #173B5E !important;
    -webkit-text-fill-color: #173B5E !important;
    border: 1px solid #B8CCE0 !important;
    border-radius: 8px !important;
}

/* DATE CALENDAR ICON */
[data-testid="stSidebar"] [data-testid="stDateInput"] svg {
    fill: #2463A6 !important;
    color: #2463A6 !important;
}

/* RESET BUTTON */
[data-testid="stSidebar"] [data-testid="stButton"] button {
    background-color: #FFFFFF !important;
    color: #2463A6 !important;
    border: 1px solid #2463A6 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    min-height: 42px !important;
}

/* RESET BUTTON HOVER */
[data-testid="stSidebar"] [data-testid="stButton"] button:hover {
    background-color: #DCEBFA !important;
    color: #173B5E !important;
    border-color: #2463A6 !important;
}

/* APPLY FILTERS BUTTON */
[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button {
    background-color: #2463A6 !important;
    color: #FFFFFF !important;
    border: 1px solid #2463A6 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    min-height: 42px !important;
}

/* APPLY BUTTON HOVER */
[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button:hover {
    background-color: #173B5E !important;
    color: #FFFFFF !important;
}

/* MAIN HEADINGS */
h1, h2, h3 {
    color: #102A43 !important;
    font-weight: 700 !important;
}

/* KPI CARDS */
.kpi-card {
    background-color: #FFFFFF;
    padding: 22px;
    border-radius: 12px;
    border: 1px solid #E2E8F0;
    border-left: 5px solid #2463A6;
    box-shadow: 0 3px 12px rgba(16, 42, 67, 0.06);
    margin-bottom: 12px;
}

.kpi-label {
    color: #627D98;
    font-size: 14px;
    font-weight: 600;
    margin-bottom: 10px;
}

.kpi-value {
    color: #102A43;
    font-size: 25px;
    font-weight: 700;
}

/* SECTION HEADINGS */
.section-heading {
    color: #102A43;
    font-size: 21px;
    font-weight: 700;
    padding-bottom: 10px;
    border-bottom: 2px solid #D9E6F2;
    margin-top: 20px;
    margin-bottom: 18px;
}

/* DOWNLOAD BUTTONS */
.stDownloadButton button {
    background-color: #174A7E !important;
    color: #FFFFFF !important;
    border-radius: 8px;
    border: none;
    font-weight: 600;
    padding: 10px;
}

.stDownloadButton button:hover {
    background-color: #2463A6 !important;
    color: #FFFFFF !important;
}

/* DATAFRAME */
[data-testid="stDataFrame"] {
    border: 1px solid #D9E2EC;
    border-radius: 8px;
}

/* ALERTS */
.stAlert {
    border-radius: 8px;
}

</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# LOAD DATA
# --------------------------------------------------

@st.cache_data
def load_data():
    df = pd.read_csv("sales_data.csv")
    df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df


df = load_data()

# --------------------------------------------------
# INITIALIZE FILTER VALUES
# --------------------------------------------------

if "regions" not in st.session_state:
    st.session_state["regions"] = df["Region"].unique().tolist()

if "categories" not in st.session_state:
    st.session_state["categories"] = df["Category"].unique().tolist()

if "products" not in st.session_state:
    st.session_state["products"] = df["Product"].unique().tolist()

if "start_date" not in st.session_state:
    st.session_state["start_date"] = df["Order Date"].min().date()

if "end_date" not in st.session_state:
    st.session_state["end_date"] = df["Order Date"].max().date()

# --------------------------------------------------
# DASHBOARD HEADER
# --------------------------------------------------

st.markdown(
    "<h1>📊 Sales & Revenue Dashboard</h1>",
    unsafe_allow_html=True
)

st.markdown(
    "<p style='color:#627D98; font-size:16px;'>"
    "Business Intelligence | Sales Performance | Revenue Insights"
    "</p>",
    unsafe_allow_html=True
)

st.divider()

# --------------------------------------------------
# SIDEBAR FILTERS
# --------------------------------------------------

st.sidebar.title("⚙️ Dashboard Controls")
st.sidebar.markdown("Customize your sales analysis below.")

# RESET BUTTON
if st.sidebar.button(
    "🔄 RESET ALL FILTERS",
    use_container_width=True,
    key="reset_filters"
):
    st.session_state["regions"] = df["Region"].unique().tolist()
    st.session_state["categories"] = df["Category"].unique().tolist()
    st.session_state["products"] = df["Product"].unique().tolist()
    st.session_state["start_date"] = df["Order Date"].min().date()
    st.session_state["end_date"] = df["Order Date"].max().date()

    st.rerun()

with st.sidebar.form(key="dashboard_filters"):

    st.markdown("### 🌍 Region")

    regions = st.multiselect(
        "Select Region",
        options=df["Region"].unique().tolist(),
        key="regions"
    )

    st.markdown("### 📦 Category")

    categories = st.multiselect(
        "Select Category",
        options=df["Category"].unique().tolist(),
        key="categories"
    )

    available_products = df[
        df["Category"].isin(categories)
    ]["Product"].unique().tolist()

    st.session_state["products"] = [
        product for product in st.session_state["products"]
        if product in available_products
    ]

    st.markdown("### 🛍️ Product")

    products = st.multiselect(
        "Select Product",
        options=available_products,
        key="products"
    )

    st.markdown("### 📅 Date Range")

    start_date = st.date_input(
        "Start Date",
        key="start_date"
    )

    end_date = st.date_input(
        "End Date",
        key="end_date"
    )

    apply_filters = st.form_submit_button(
        "Apply Filters",
        use_container_width=True
    )

# --------------------------------------------------
# FILTER DATA
# --------------------------------------------------

@st.cache_data
def filter_sales_data(
    data,
    selected_regions,
    selected_categories,
    selected_products,
    start,
    end
):
    return data[
        (data["Region"].isin(selected_regions)) &
        (data["Category"].isin(selected_categories)) &
        (data["Product"].isin(selected_products)) &
        (data["Order Date"].dt.date >= start) &
        (data["Order Date"].dt.date <= end)
    ]


filtered_df = filter_sales_data(
    df,
    tuple(regions),
    tuple(categories),
    tuple(products),
    start_date,
    end_date
)

# --------------------------------------------------
# KPI CALCULATIONS
# --------------------------------------------------

st.markdown(
    "<div class='section-heading'>📌 Business Performance Overview</div>",
    unsafe_allow_html=True
)

total_revenue = filtered_df["Revenue"].sum()
total_profit = filtered_df["Profit"].sum()
total_orders = filtered_df["Order ID"].nunique()
total_quantity = filtered_df["Quantity"].sum()

average_order_value = (
    total_revenue / total_orders
    if total_orders > 0 else 0
)

profit_margin = (
    (total_profit / total_revenue) * 100
    if total_revenue > 0 else 0
)


def show_kpi(label, value):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# FIRST KPI ROW
col1, col2, col3 = st.columns(3)

with col1:
    show_kpi("💰 Total Revenue", f"₹{total_revenue:,.0f}")

with col2:
    show_kpi("📈 Total Profit", f"₹{total_profit:,.0f}")

with col3:
    show_kpi("🧾 Total Orders", f"{total_orders:,}")


# SECOND KPI ROW
col4, col5, col6 = st.columns(3)

with col4:
    show_kpi("📦 Units Sold", f"{total_quantity:,}")

with col5:
    show_kpi("🛒 Average Order Value", f"₹{average_order_value:,.0f}")

with col6:
    show_kpi("📊 Profit Margin", f"{profit_margin:.2f}%")


st.divider()

# --------------------------------------------------
# CHARTS
# --------------------------------------------------

if not filtered_df.empty:

    # MONTHLY REVENUE TREND
    st.markdown(
        "<div class='section-heading'>📈 Monthly Revenue Trend</div>",
        unsafe_allow_html=True
    )

    monthly_sales = (
        filtered_df.groupby(
            filtered_df["Order Date"].dt.to_period("M")
        )["Revenue"].sum().reset_index()
    )

    monthly_sales["Order Date"] = (
        monthly_sales["Order Date"].astype(str)
    )

    fig1 = px.line(
        monthly_sales,
        x="Order Date",
        y="Revenue",
        markers=True,
        title="Monthly Revenue Performance",
        template="plotly_white"
    )

    fig1.update_traces(
        line=dict(color="#2463A6", width=3),
        marker=dict(size=8)
    )

    fig1.update_layout(
        title_font_color="#102A43",
        font_color="#334E68",
        plot_bgcolor="white",
        paper_bgcolor="white"
    )

    st.plotly_chart(
        fig1,
        use_container_width=True,
        config={"displayModeBar": False}
    )

    # PRODUCT AND CATEGORY ANALYSIS
    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            "<div class='section-heading'>🏆 Top Performing Products</div>",
            unsafe_allow_html=True
        )

        product_sales = (
            filtered_df.groupby("Product")["Revenue"]
            .sum()
            .reset_index()
            .sort_values("Revenue", ascending=False)
        )

        fig2 = px.bar(
            product_sales,
            x="Product",
            y="Revenue",
            color="Revenue",
            title="Revenue by Product",
            color_continuous_scale=["#B8D5F0", "#174A7E"],
            template="plotly_white"
        )

        fig2.update_layout(
            title_font_color="#102A43",
            font_color="#334E68",
            plot_bgcolor="white",
            paper_bgcolor="white",
            coloraxis_showscale=False
        )

        st.plotly_chart(
            fig2,
            use_container_width=True,
            config={"displayModeBar": False}
        )

    with col2:

        st.markdown(
            "<div class='section-heading'>📦 Category Analysis</div>",
            unsafe_allow_html=True
        )

        category_sales = (
            filtered_df.groupby("Category")["Revenue"]
            .sum()
            .reset_index()
        )

        fig3 = px.pie(
            category_sales,
            names="Category",
            values="Revenue",
            hole=0.45,
            title="Revenue by Category",
            color_discrete_sequence=[
                "#174A7E",
                "#3B82B8",
                "#8AB6D6",
                "#B8D5F0"
            ]
        )

        fig3.update_layout(
            title_font_color="#102A43",
            font_color="#334E68",
            paper_bgcolor="white"
        )

        st.plotly_chart(
            fig3,
            use_container_width=True,
            config={"displayModeBar": False}
        )

    # SALES QUANTITY ANALYSIS
    st.markdown(
        "<div class='section-heading'>📊 Sales Quantity Analysis</div>",
        unsafe_allow_html=True
    )

    quantity_sales = (
        filtered_df.groupby("Product")["Quantity"]
        .sum()
        .reset_index()
        .sort_values("Quantity", ascending=False)
    )

    fig4 = px.bar(
        quantity_sales,
        x="Product",
        y="Quantity",
        color="Product",
        title="Units Sold by Product",
        color_discrete_sequence=[
            "#B8D5F0",
            "#174A7E",
            "#3B82B8",
            "#8AB6D6",
            "#2463A6",
            "#527FA5",
            "#9ABAD3",
            "#D0E2F0"
        ],
        template="plotly_white"
    )

    fig4.update_layout(
        title_font_color="#102A43",
        font_color="#334E68",
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False
    )

    st.plotly_chart(
        fig4,
        use_container_width=True,
        config={"displayModeBar": False}
    )

    # REGIONAL ANALYSIS
    st.markdown(
        "<div class='section-heading'>🌍 Regional Sales Analysis</div>",
        unsafe_allow_html=True
    )

    region_sales = (
        filtered_df.groupby("Region")["Revenue"]
        .sum()
        .reset_index()
        .sort_values("Revenue", ascending=False)
    )

    fig5 = px.bar(
        region_sales,
        x="Region",
        y="Revenue",
        color="Region",
        title="Revenue by Region",
        color_discrete_sequence=[
            "#174A7E",
            "#3B82B8",
            "#8AB6D6",
            "#2463A6",
            "#9ABAD3"
        ],
        template="plotly_white"
    )

    fig5.update_layout(
        title_font_color="#102A43",
        font_color="#334E68",
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False
    )

    st.plotly_chart(
        fig5,
        use_container_width=True,
        config={"displayModeBar": False}
    )

    # PROFIT ANALYSIS
    st.markdown(
        "<div class='section-heading'>💰 Profit Analysis by Product</div>",
        unsafe_allow_html=True
    )

    product_profit = (
        filtered_df.groupby("Product")["Profit"]
        .sum()
        .reset_index()
        .sort_values("Profit", ascending=False)
    )

    fig6 = px.bar(
        product_profit,
        x="Product",
        y="Profit",
        color="Profit",
        title="Profit Generated by Each Product",
        color_continuous_scale=["#B8D5F0", "#174A7E"],
        template="plotly_white"
    )

    fig6.update_layout(
        title_font_color="#102A43",
        font_color="#334E68",
        plot_bgcolor="white",
        paper_bgcolor="white",
        coloraxis_showscale=False
    )

    st.plotly_chart(
        fig6,
        use_container_width=True,
        config={"displayModeBar": False}
    )

    # --------------------------------------------------
    # TRANSACTION DETAILS
    # --------------------------------------------------

    st.markdown(
        "<div class='section-heading'>📋 Sales Transaction Details</div>",
        unsafe_allow_html=True
    )

    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------
    # DOWNLOAD REPORTS
    # --------------------------------------------------

    st.markdown(
        "<div class='section-heading'>📥 Download Sales Reports</div>",
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    # CSV DOWNLOAD
    csv_data = filtered_df.to_csv(
        index=False
    ).encode("utf-8")

    with col1:
        st.download_button(
            label="📥 Download CSV Report",
            data=csv_data,
            file_name="sales_report.csv",
            mime="text/csv",
            key="csv_download",
            use_container_width=True
        )

    # EXCEL DOWNLOAD
    excel_buffer = BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        filtered_df.to_excel(
            writer,
            index=False,
            sheet_name="Sales Report"
        )

        summary_data = pd.DataFrame({
            "Business Metric": [
                "Total Revenue",
                "Total Profit",
                "Total Orders",
                "Units Sold",
                "Average Order Value",
                "Profit Margin"
            ],
            "Value": [
                total_revenue,
                total_profit,
                total_orders,
                total_quantity,
                average_order_value,
                profit_margin
            ]
        })

        summary_data.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

        workbook = writer.book

        for worksheet in workbook.worksheets:

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for cell in worksheet[1]:
                cell.font = Font(
                    bold=True,
                    color="FFFFFF",
                    size=11
                )

                cell.fill = PatternFill(
                    start_color="174A7E",
                    end_color="174A7E",
                    fill_type="solid"
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )

            for column in worksheet.columns:
                max_length = max(
                    len(str(cell.value or ""))
                    for cell in column
                )

                worksheet.column_dimensions[
                    column[0].column_letter
                ].width = max_length + 4

            thin_border = Border(
                bottom=Side(
                    style="thin",
                    color="D9E2EC"
                )
            )

            for row in worksheet.iter_rows():
                for cell in row:
                    cell.border = thin_border

        summary_sheet = workbook["Summary"]

        for row in range(2, summary_sheet.max_row + 1):

            metric = summary_sheet.cell(
                row=row,
                column=1
            ).value

            value_cell = summary_sheet.cell(
                row=row,
                column=2
            )

            if metric in [
                "Total Revenue",
                "Total Profit",
                "Average Order Value"
            ]:
                value_cell.number_format = '₹#,##0.00'

            elif metric == "Profit Margin":
                value_cell.number_format = '0.00"%"'

            else:
                value_cell.number_format = '#,##0'

    excel_buffer.seek(0)
    excel_data = excel_buffer.read()

    with col2:
        st.download_button(
            label="📊 Download Excel Report",
            data=excel_data,
            file_name="sales_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="excel_download",
            use_container_width=True
        )

else:
    st.warning(
        "No sales records match the selected filters. "
        "Please change your filters."
    )

# --------------------------------------------------
# FOOTER
# --------------------------------------------------

st.divider()

st.markdown(
    """
    <div style="text-align:center; color:#627D98; padding:10px;">
        Sales & Revenue Analysis Dashboard |
        Built with Python, Streamlit, Pandas and Plotly
    </div>
    """,
    unsafe_allow_html=True
)
