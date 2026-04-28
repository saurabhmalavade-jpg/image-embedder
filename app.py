# app.py — Streamlit UI for Excel Image Embedder
import streamlit as st
from core import process_excel

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel Image Embedder",
    page_icon="🖼️",
    layout="centered"
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fb; }
    .stButton > button {
        background-color: #2563eb;
        color: white;
        border: none;
        padding: 0.6rem 2rem;
        border-radius: 8px;
        font-size: 16px;
        font-weight: 600;
        width: 100%;
        transition: background 0.2s;
    }
    .stButton > button:hover { background-color: #1d4ed8; }
    .stat-box {
        background: white;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("🖼️ Excel Image Embedder")
st.markdown(
    "Upload an Excel file containing **image URLs**. "
    "The tool will download each image and embed it directly into the cell — "
    "all other text stays untouched."
)
st.divider()

# ── File Upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📂 Upload your Excel file (.xlsx)",
    type=["xlsx"],
    help="The file should have cells containing http:// or https:// image URLs."
)

if uploaded_file:
    st.success(f"✅ File uploaded: **{uploaded_file.name}**")

    # Output filename suggestion
    default_name = uploaded_file.name.replace(".xlsx", "_with_images.xlsx")
    output_filename = st.text_input(
        "💾 Output filename",
        value=default_name,
        help="Name for the processed file you will download."
    )
    if not output_filename.endswith(".xlsx"):
        output_filename += ".xlsx"

    st.divider()

    # ── Process Button ────────────────────────────────────────────────────────
    if st.button("🚀 Start Processing"):
        input_bytes = uploaded_file.read()

        # Progress bar + status text
        progress_bar = st.progress(0, text="Starting…")
        status_text  = st.empty()

        def update_progress(completed, total):
            pct  = int((completed / total) * 100)
            progress_bar.progress(pct, text=f"Processing… {completed}/{total} images")
            status_text.markdown(f"⏳ **{completed}** of **{total}** images done")

        with st.spinner("Processing your Excel file…"):
            try:
                output_bytes, stats = process_excel(input_bytes, progress_callback=update_progress)
            except Exception as e:
                st.error(f"❌ Something went wrong: {e}")
                st.stop()

        progress_bar.progress(100, text="Done!")
        status_text.empty()

        # ── Results Summary ───────────────────────────────────────────────────
        st.divider()
        st.subheader("📊 Results")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(
                f'<div class="stat-box"><h2>{stats["total"]}</h2><p>Total URLs found</p></div>',
                unsafe_allow_html=True
            )
        with col2:
            st.markdown(
                f'<div class="stat-box" style="border-top:3px solid #16a34a">'
                f'<h2 style="color:#16a34a">{stats["success"]}</h2>'
                f'<p>Images embedded ✅</p></div>',
                unsafe_allow_html=True
            )
        with col3:
            st.markdown(
                f'<div class="stat-box" style="border-top:3px solid #dc2626">'
                f'<h2 style="color:#dc2626">{stats["error"]}</h2>'
                f'<p>Failed / Corrupt ❌</p></div>',
                unsafe_allow_html=True
            )

        if stats["total"] == 0:
            st.warning("⚠️ No image URLs were found in the uploaded file.")
        else:
            st.success("🎉 Processing complete! Download your file below.")

        # ── Download Button ───────────────────────────────────────────────────
        st.divider()
        st.download_button(
            label="⬇️ Download Processed Excel File",
            data=output_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ── Footer when no file uploaded ─────────────────────────────────────────────
else:
    st.info("👆 Upload an Excel file to get started.")
    with st.expander("ℹ️ How does this work?"):
        st.markdown("""
        1. **Upload** your `.xlsx` file containing image URLs in any cell.
        2. Click **Start Processing** — the tool downloads all images concurrently (fast!).
        3. Each image is embedded directly into the cell that had the URL.
        4. Cells with text/numbers are **not touched**.
        5. If an image fails to load, the cell shows a red error message.
        6. **Download** your finished file.
        """)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.footer {
    position: fixed;
    bottom: 0;
    left: 0;
    width: 100%;
    background-color: #fff;
    color: #888ea8;
    text-align: center;
    padding: 10px 0;
    font-size: 0.75rem;
    border-top: 1px solid #dde1ef;
    font-family: 'Inter', sans-serif;
    font-weight: 500;
    letter-spacing: 0.03em;
    z-index: 999;
}
</style>
<div class="footer">
    © Designed and Developed by Saurabh Malavade
</div>
""", unsafe_allow_html=True)
