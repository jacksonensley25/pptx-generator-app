import streamlit as st
import datetime
from main5 import generate_pptx

st.title("PPTX Generator")

# Text inputs
retailer_id  = st.text_input("Retailer ID")
brand_name   = st.text_input("Brand Name")
campaign_id  = st.text_input("Campaign ID (comma-separated if multiple)")
line_item_id = st.text_input("Line-Item ID (comma-separated if multiple)")

# Slide 1 date-range widget
default_start_1 = datetime.date.today() - datetime.timedelta(days=30)
default_end_1   = datetime.date.today()
slide1_range = st.date_input(
    "Slide 1 Date Range",
    value=(default_start_1, default_end_1)
)

# Slide 2 date-range widget
default_start_2 = datetime.date.today().replace(day=1)
default_end_2   = datetime.date.today()
slide2_range = st.date_input(
    "Slide 2 Date Range",
    value=(default_start_2, default_end_2)
)

if st.button("Generate Presentation"):
    # 1) Validate date ranges
    if not (isinstance(slide1_range, (tuple, list)) and len(slide1_range) == 2):
        st.error("🔴 Please select both a start and an end date for Slide 1.")
        st.stop()
    if not (isinstance(slide2_range, (tuple, list)) and len(slide2_range) == 2):
        st.error("🔴 Please select both a start and an end date for Slide 2.")
        st.stop()

    # 2) Parse date values
    slide1_start, slide1_end = slide1_range
    slide2_start, slide2_end = slide2_range
    s1, e1 = slide1_start.strftime("%Y-%m-%d"), slide1_end.strftime("%Y-%m-%d")
    s2, e2 = slide2_start.strftime("%Y-%m-%d"), slide2_end.strftime("%Y-%m-%d")

    # 3) Parse campaign IDs
    try:
        campaign_ids = [
            int(x.strip()) for x in campaign_id.split(",") if x.strip()
        ]
        if not campaign_ids:
            raise ValueError
    except ValueError:
        st.error("🔴 Please enter one or more numeric Campaign IDs, separated by commas.")
        st.stop()

    # 4) Parse line-item IDs
    try:
        line_item_ids = [
            int(x.strip()) for x in line_item_id.split(",") if x.strip()
        ]
        if not line_item_ids:
            raise ValueError
    except ValueError:
        st.error("🔴 Please enter one or more numeric Line-Item IDs, separated by commas.")
        st.stop()

    # 5) Generate PPTX
    with st.spinner("Building your PowerPoint…"):
        pptx_path = generate_pptx(
            retailer_id=retailer_id,
            brand_name=brand_name,
            campaign_id=campaign_ids,
            line_items_pivot=line_item_ids,
            line_items_capout=",".join(str(i) for i in line_item_ids),
            slide1_start=s1,
            slide1_end=e1,
            slide2_start=s2,
            slide2_end=e2,
        )

    st.success(f"✅ Presentation generated: {pptx_path}")
    with open(pptx_path, "rb") as f:
        st.download_button("Download PPTX", f, file_name=pptx_path)
