from pathlib import Path
import shutil
import streamlit as st

from t1_generate_cd import generate_cd
from t2_generate_e_base import generate_e_base
from t3_generate_x import generate_x


INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")


def clear_folder(folder: Path):
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)


def save_uploaded_file(uploaded_file, folder: Path):
    folder.mkdir(parents=True, exist_ok=True)
    file_path = folder / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path


def download_button(label, file_path):
    file_path = Path(file_path)
    if file_path.exists():
        with open(file_path, "rb") as f:
            st.download_button(
                label=label,
                data=f,
                file_name=file_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


st.title("Billing Automation")

st.warning("Use the clear button before a new run to avoid mixing old files.")

if st.button("Clear all input/output files"):
    clear_folder(INPUT_DIR)
    clear_folder(OUTPUT_DIR)
    st.success("Cleared input and output folders.")


st.header("T1: Generate C and D")

shipment_file = st.file_uploader("Upload Shipment File", type=["xlsx", "xlsb", "xls"], key="shipment")
trip_file = st.file_uploader("Upload Trip File", type=["xlsx", "xlsb", "xls"], key="trip")

if st.button("Run T1"):
    if not shipment_file or not trip_file:
        st.error("Upload both Shipment and Trip files.")
    else:
        clear_folder(Path("input/t1/shipment"))
        clear_folder(Path("input/t1/trip"))
        Path("output/t1").mkdir(parents=True, exist_ok=True)

        save_uploaded_file(shipment_file, Path("input/t1/shipment"))
        save_uploaded_file(trip_file, Path("input/t1/trip"))

        generate_cd()
        st.success("T1 complete.")

download_button("Download C_Delivered", "output/t1/C_Delivered.xlsx")
download_button("Download D_Pickup", "output/t1/D_Pickup.xlsx")


st.header("T2: Generate E Base")

c_file = st.file_uploader("Upload C' Cleaned File", type=["xlsx", "xlsb", "xls"], key="c_cleaned")
d_file = st.file_uploader("Upload D' Cleaned File", type=["xlsx", "xlsb", "xls"], key="d_cleaned")
f_file = st.file_uploader("Upload F Pincode File", type=["xlsx", "xlsb", "xls"], key="f")
g_file = st.file_uploader("Upload G Hub Zone File", type=["xlsx", "xlsb", "xls"], key="g")
h_file = st.file_uploader("Upload H Rate Card File", type=["xlsx", "xlsb", "xls"], key="h")
hub_master_file = st.file_uploader("Upload Hub Master File", type=["xlsx", "xlsb", "xls"], key="hub_master")

if st.button("Run T2"):
    required = [c_file, d_file, f_file, g_file, h_file, hub_master_file]

    if not all(required):
        st.error("Upload all T2 files.")
    else:
        clear_folder(Path("input/t2/c_cleaned"))
        clear_folder(Path("input/t2/d_cleaned"))
        clear_folder(Path("input/t2/f"))
        clear_folder(Path("input/t2/g"))
        clear_folder(Path("input/t2/h"))
        clear_folder(Path("input/t2/hub_master"))
        Path("output/t2").mkdir(parents=True, exist_ok=True)

        save_uploaded_file(c_file, Path("input/t2/c_cleaned"))
        save_uploaded_file(d_file, Path("input/t2/d_cleaned"))
        save_uploaded_file(f_file, Path("input/t2/f"))
        save_uploaded_file(g_file, Path("input/t2/g"))
        save_uploaded_file(h_file, Path("input/t2/h"))
        save_uploaded_file(hub_master_file, Path("input/t2/hub_master"))

        generate_e_base()
        st.success("T2 complete.")

download_button("Download E_Base", "output/t2/E_Base.xlsx")

st.header("T3: Generate X and Final E")

e_base_file = st.file_uploader(
    "Upload E_Base File",
    type=["xlsx", "xlsb", "xls"],
    key="t3_e_base",
)

x_base_file = st.file_uploader(
    "Upload X_Base / E2E Rate Card File",
    type=["xlsx", "xlsb", "xls"],
    key="t3_x_base",
)

if st.button("Run T3"):
    if not e_base_file or not x_base_file:
        st.error("Upload both E_Base and X_Base files.")
    else:
        clear_folder(Path("input/t3/e_base"))
        clear_folder(Path("input/t3/x_base"))
        Path("output/t3").mkdir(parents=True, exist_ok=True)

        save_uploaded_file(e_base_file, Path("input/t3/e_base"))
        save_uploaded_file(x_base_file, Path("input/t3/x_base"))

        generate_x()
        st.success("T3 complete.")

download_button("Download X", "output/t3/X.xlsx")
download_button("Download E", "output/t3/E.xlsx")