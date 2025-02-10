import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from io import BytesIO

EXCEL_LOG_FILE = "flagged_content.xlsx"  # Ensure this matches your log file

# Load Data
def load_data():
    try:
        return pd.read_excel(EXCEL_LOG_FILE)
    except FileNotFoundError:
        st.error("Log file not found. Please ensure the log file exists.")
        return pd.DataFrame(columns=["Timestamp", "Platform", "Flagged Content"])

# Authentication
def authenticate():
    password = st.text_input("Enter Password:", type="password")
    if password == "1519":  # Replace 'your_password' with your desired password
        st.success("Access Granted")
        return True
    else:
        st.error("Access Denied")
        return False

# Dashboard
def main():
    st.title("Flagged Content Analytics Dashboard")

    # Load the data
    data = load_data()

    if data.empty:
        st.warning("No data available to display.")
        return

    # Data Overview
    st.header("Overview")
    st.write(f"Total Records: {len(data)}")
    st.write(f"Time Range: {data['Timestamp'].min()} to {data['Timestamp'].max()}")

    # Show Raw Data
    if st.checkbox("Show Raw Data"):
        st.dataframe(data)

    # Filter Data
    st.header("Filter Data")
    platforms = data["Platform"].unique()
    selected_platforms = st.multiselect("Filter by Platform", platforms, default=platforms)
    filtered_data = data[data["Platform"].isin(selected_platforms)]

    # Download Filtered Data
    if st.button("Download Filtered Data as CSV"):
        csv = filtered_data.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download",
            data=csv,
            file_name="filtered_data.csv",
            mime="text/csv",
        )

    # Visualize Most Frequent Platforms
    st.header("Most Frequent Platforms")
    platform_counts = filtered_data["Platform"].value_counts()
    st.bar_chart(platform_counts)

    # Flagged Words Analysis
    st.header("Most Common Flagged Words")
    word_counts = filtered_data["Flagged Content"].str.split(expand=True).stack().value_counts().head(10)
    st.bar_chart(word_counts)

    # Word Cloud Visualization
    st.header("Generate Word Cloud")
    if st.checkbox("Generate Word Cloud"):
        word_cloud = WordCloud(width=800, height=400, background_color="white").generate(
            " ".join(filtered_data["Flagged Content"].dropna())
        )
        fig, ax = plt.subplots()
        ax.imshow(word_cloud, interpolation="bilinear")
        ax.axis("off")
        st.pyplot(fig)

    # Trends Over Time
    st.header("Flagged Events Over Time")
    try:
        filtered_data["Date"] = pd.to_datetime(filtered_data["Timestamp"], format="%Y-%m-%d_%H-%M-%S", errors="coerce").dt.date
        daily_counts = filtered_data.groupby("Date").size()
        st.line_chart(daily_counts)
    except Exception as e:
        st.error(f"Error processing date data: {e}")

# Run the App
if authenticate():
    main()
