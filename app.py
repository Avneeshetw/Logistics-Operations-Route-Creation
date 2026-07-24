import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import itertools
import os
import requests
import datetime
import io
from folium.plugins import PolyLineTextPath, Fullscreen

# Page Layout Configurations
st.set_page_config(
    page_title="Interactive Logistics Router", 
    layout="wide",
    page_icon="DrishtiLogo.png"
)

# Safe Logo Placement 
logo_path = "DrishtiLogo.png"
if os.path.exists(logo_path):
    st.logo(logo_path)

st.markdown("<style>.block-container { padding-top: 1rem; padding-bottom: 1rem; }</style>", unsafe_allow_html=True)

# Header
st.title("🛠️ Logistics Operations & Route Creation")
st.write("---")

excel_file = "Location.xlsx"
backend_excel = "dispatch_master.xlsx"
vehicle_master_file = "vehicle_master.xlsx"

# Helper: Load/Create Dispatch Excel Data
def get_backend_excel_data():
    standard_cols = ["Date", "Warehouse / Plant", "Distributors / DBRs", "Total Load (Ton)", "Dispatch Status", "Vehicle Alloted"]
    if os.path.exists(backend_excel):
        try:
            df = pd.read_excel(backend_excel)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed|Route')]
            for col in standard_cols:
                if col not in df.columns:
                    df[col] = ""
            return df[standard_cols]
        except Exception:
            pass
    return pd.DataFrame(columns=standard_cols)

# 🚛 Load / Initialize Vehicle Master File
def load_vehicle_master():
    standard_cols = ["Vehicle Number", "OwnershipType", "Location", "Transporter Name", "CapacityTonnage", "Remarks"]
    if os.path.exists(vehicle_master_file):
        try:
            v_df = pd.read_excel(vehicle_master_file)
            v_df.columns = v_df.columns.str.strip()
            for col in standard_cols:
                if col not in v_df.columns:
                    v_df[col] = ""
            return v_df[standard_cols]
        except Exception:
            pass
    
    default_data = [
        {"Vehicle Number": "UP32CZ7228", "OwnershipType": "Own", "Location": "Safedabad", "Transporter Name": "Company Vehicles", "CapacityTonnage": 3.3, "Remarks": "Operational"},
        {"Vehicle Number": "UP32HN2137", "OwnershipType": "Own", "Location": "Safedabad", "Transporter Name": "Company Vehicles", "CapacityTonnage": 3.7, "Remarks": "Operational"},
        {"Vehicle Number": "UP32CZ1713", "OwnershipType": "Own", "Location": "Safedabad", "Transporter Name": "Company Vehicles", "CapacityTonnage": 4.0, "Remarks": "Operational"},
        {"Vehicle Number": "UP32FT6643", "OwnershipType": "Own", "Location": "Safedabad", "Transporter Name": "Company Vehicles", "CapacityTonnage": 5.0, "Remarks": "Operational"}
    ]
    df = pd.DataFrame(default_data)
    df.to_excel(vehicle_master_file, index=False)
    return df

# 🟢 STRICTLY OPERATIONAL & AVAILABLE VEHICLES
def get_only_available_vehicles():
    v_df = load_vehicle_master()
    
    valid_remarks = ["operational", "returned / available", "available"]
    op_mask = v_df["Remarks"].astype(str).str.strip().str.lower().isin(valid_remarks)
    operational_vehicles = v_df[op_mask]["Vehicle Number"].dropna().astype(str).str.strip().tolist()
    
    dispatch_df = get_backend_excel_data()
    active_dispatches = dispatch_df[dispatch_df["Dispatch Status"].isin(["Alloted", "Dispatched"])]
    currently_busy = active_dispatches["Vehicle Alloted"].dropna().astype(str).str.strip().tolist()
    
    truly_available = [v for v in operational_vehicles if v not in currently_busy and v != ""]
    return truly_available

# Road Router Engine
def get_road_route_and_distance(coords_list):
    loc_string = ";".join([f"{lon},{lat}" for lat, lon in coords_list])
    url = f"http://router.project-osrm.org/route/v1/driving/{loc_string}?overview=full&geometries=geojson&continue_straight=true"
    try:
        response = requests.get(url, timeout=5).json()
        if response.get("code") == "Ok":
            route = response["routes"][0]
            distance_km = route["distance"] / 1000.0
            road_geometry = [(lat, lon) for lon, lat in route["geometry"]["coordinates"]]
            return distance_km, road_geometry
    except Exception:
        pass
    return None, None

# Generate Excel Bytes
def generate_excel_download(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

if not os.path.exists(excel_file):
    st.error(f"⚠️ '{excel_file}' file nahi mili! Folder me Check karein.")
else:
    raw_df = pd.read_excel(excel_file)
    raw_df.columns = raw_df.columns.str.strip()
    cols = list(raw_df.columns)
    
    name_col = next((c for c in cols if 'location' in c.lower() or 'name' in c.lower()), cols[1])
    type_col = next((c for c in cols if 'type' in c.lower()), cols[2])
    lat_col = next((c for c in cols if 'lat' in c.lower()), cols[3])
    lon_col = next((c for c in cols if 'lon' in c.lower() or 'long' in c.lower()), cols[4])

    df = raw_df.dropna(subset=[lat_col, lon_col]).copy()

    wh_df = df[df[type_col].astype(str).str.upper().isin(['WAREHOUSE', 'PLANT', 'WH'])]
    if wh_df.empty: wh_df = df.copy()

    dbr_df = df[~df[type_col].astype(str).str.upper().isin(['WAREHOUSE', 'PLANT', 'WH'])]
    if dbr_df.empty: dbr_df = df.copy()
    
    wh_list = list(wh_df[name_col].unique())
    dbr_list = list(dbr_df[name_col].unique())

    if 'form_key' not in st.session_state:
        st.session_state.form_key = 0
    if 'status_tab_key' not in st.session_state:
        st.session_state.status_tab_key = 0

    # =========================================================
    # 🌟 TOP PANEL
    # =========================================================
    left_col, right_col = st.columns([1.1, 1.9])

    with left_col:
        tab1, tab2, tab3 = st.tabs(["📝 Dispatch Form", "🛠️ Status Manager", "📁 Bulk Uploads"])

        # -------------------------------------------------------------
        # TAB 1: DISPATCH FORM
        # -------------------------------------------------------------
        with tab1:
            entry_date = st.date_input("Date", datetime.date.today(), key=f"date_{st.session_state.form_key}")
            entry_wh = st.selectbox("Warehouse / Plant", options=wh_list, key=f"wh_{st.session_state.form_key}")
            
            entry_dbrs = st.multiselect(
                "Distributors / DBRs", 
                options=dbr_list, 
                key=f"dbrs_{st.session_state.form_key}", 
                placeholder="Select DBRs"
            )
            
            f_c1, f_c2 = st.columns(2)
            with f_c1:
                entry_load = st.number_input("Total Load (Ton)", min_value=0.1, max_value=100.0, value=5.0, step=0.1, key=f"load_{st.session_state.form_key}")
            with f_c2:
                entry_status = st.selectbox("Dispatch Status", options=["Pending", "Alloted", "Dispatched", "Completed / Delivered"], key=f"status_{st.session_state.form_key}")
                
            available_vehicles = get_only_available_vehicles()
            vehicle_options = ["Select Vehicle"] + available_vehicles
            
            if entry_status == "Pending":
                entry_vehicle = st.selectbox("Vehicle Alloted (Optional for Pending)", options=["Not Required"] + available_vehicles, key=f"veh_{st.session_state.form_key}")
            else:
                if len(available_vehicles) == 0:
                    st.warning("⚠️ Abhi koi bhi Operational/Available vehicle bacha nahi hai!")
                entry_vehicle = st.selectbox("Vehicle Alloted (Only Available)", options=vehicle_options, key=f"veh_{st.session_state.form_key}")

            b1, b2 = st.columns(2)
            with b1:
                save_btn = st.button("💾 Save Entry", use_container_width=True)
            with b2:
                if st.button("🧹 Clear / Reset Filter", use_container_width=True):
                    st.session_state.form_key += 1
                    st.rerun()

            if save_btn:
                if not entry_dbrs:
                    st.warning("⚠️ Kam se kam ek DBR zaroor chunein!")
                elif entry_status != "Pending" and (entry_vehicle == "Select Vehicle" or not entry_vehicle):
                    st.error("⚠️ Status 'Alloted' ya 'Dispatched' ke liye Available Vehicle select karein!")
                else:
                    try:
                        existing_df = get_backend_excel_data()
                        veh_val = "" if (entry_status == "Pending" and entry_vehicle in ["Not Required", "Select Vehicle"]) else entry_vehicle
                        
                        new_row = {
                            "Date": entry_date.strftime("%d/%m/%y"),
                            "Warehouse / Plant": entry_wh,
                            "Distributors / DBRs": ", ".join(entry_dbrs),
                            "Total Load (Ton)": float(entry_load),
                            "Dispatch Status": entry_status,
                            "Vehicle Alloted": veh_val
                        }
                        updated_df = pd.concat([existing_df, pd.DataFrame([new_row])], ignore_index=True)
                        updated_df.to_excel(backend_excel, index=False)
                        st.success("✅ Entry Save Ho Gayi!")
                        st.session_state.form_key += 1
                        st.rerun()
                    except PermissionError:
                        st.error("❌ 'dispatch_master.xlsx' Excel me KHULI HUI HAI! Band karein.")

        # -------------------------------------------------------------
        # TAB 2: STATUS MANAGER
        # -------------------------------------------------------------
        with tab2:
            st.markdown("##### 🚚 1. Update Vehicle Master Status")
            v_master_df = load_vehicle_master()
            all_vehicles = v_master_df["Vehicle Number"].dropna().astype(str).str.strip().tolist()
            
            selected_v_update = st.selectbox("Select Vehicle Number", options=all_vehicles, key=f"st_tab_v_{st.session_state.status_tab_key}")
            status_options = ["Operational", "Vehicle Breakdown", "Driver not avaialble", "Returned / Available", "Not Available"]
            
            curr_row = v_master_df[v_master_df["Vehicle Number"] == selected_v_update]
            curr_remark = curr_row["Remarks"].values[0] if len(curr_row) > 0 else "Operational"
            
            new_remark = st.selectbox("Update Status / Remarks", options=status_options, index=status_options.index(curr_remark) if curr_remark in status_options else 0, key=f"st_tab_r_{st.session_state.status_tab_key}")
            
            if st.button("🔄 Update Vehicle Status", use_container_width=True):
                v_master_df.loc[v_master_df["Vehicle Number"] == selected_v_update, "Remarks"] = new_remark
                v_master_df.to_excel(vehicle_master_file, index=False)
                
                if new_remark.lower() in ["vehicle breakdown", "driver not avaialble", "not available"]:
                    dispatch_df = get_backend_excel_data()
                    mask = (dispatch_df["Vehicle Alloted"].astype(str).str.strip() == selected_v_update) & (dispatch_df["Dispatch Status"] == "Alloted")
                    
                    if mask.any():
                        dispatch_df.loc[mask, "Dispatch Status"] = "Pending"
                        dispatch_df.loc[mask, "Vehicle Alloted"] = ""
                        dispatch_df.to_excel(backend_excel, index=False)
                        st.warning(f"⚠️ Vehicle {selected_v_update} breakdown hone ki wajah se alloted dispatch entry ko wapas 'Pending' kar diya gaya hai!")
                
                st.session_state.status_tab_key += 1
                st.session_state.form_key += 1
                st.success(f"✅ Vehicle {selected_v_update} -> Status '{new_remark}' Updated!")
                st.rerun()

            st.write("---")
            st.markdown("##### 📦 2. Update Dispatch Status")
            dispatch_data = get_backend_excel_data()
            
            if not dispatch_data.empty:
                dispatch_options = []
                for idx, row in dispatch_data.iterrows():
                    veh_str = str(row['Vehicle Alloted']).strip()
                    veh_display = veh_str if veh_str and veh_str.lower() != 'nan' and veh_str.lower() != 'none' else 'No Vehicle'
                    date_str = str(row['Date']).strip()
                    status_str = str(row['Dispatch Status']).strip()
                    
                    # Exact format matching screenshot: #0 | UP32CZ7228 | Completed / Delivered (25/07/26)
                    disp_str = f"#{idx} | {veh_display} | {status_str} ({date_str})"
                    dispatch_options.append((idx, disp_str))
                
                selected_dispatch_tuple = st.selectbox(
                    "Select Entry", 
                    options=dispatch_options, 
                    format_func=lambda x: x[1], 
                    key=f"st_disp_{st.session_state.status_tab_key}"
                )
                
                new_disp_status = st.selectbox(
                    "New Status", 
                    options=["Pending", "Alloted", "Dispatched", "Completed / Delivered"], 
                    key=f"st_disp_stat_{st.session_state.status_tab_key}"
                )
                
                if st.button("🔄 Update Status", use_container_width=True):
                    selected_idx = selected_dispatch_tuple[0]
                    dispatch_data.at[selected_idx, "Dispatch Status"] = new_disp_status
                    if new_disp_status == "Pending":
                        dispatch_data.at[selected_idx, "Vehicle Alloted"] = ""
                        
                    dispatch_data.to_excel(backend_excel, index=False)
                    st.session_state.status_tab_key += 1
                    st.session_state.form_key += 1
                    st.success(f"✅ Entry #{selected_idx} updated to '{new_disp_status}'!")
                    st.rerun()
            else:
                st.info("No saved dispatches found.")

        # -------------------------------------------------------------
        # TAB 3: BULK UPLOADS
        # -------------------------------------------------------------
        with tab3:
            st.markdown("##### 📤 Bulk Vehicle & Dispatch Upload")
            st.subheader("1. Bulk Vehicle Status Master")
            
            sample_v_df = load_vehicle_master()
            v_template_bytes = generate_excel_download(sample_v_df)
            st.download_button(
                label="📥 Download Vehicle Master Template",
                data=v_template_bytes,
                file_name="vehicle_master_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            v_file = st.file_uploader("Upload Updated Vehicle Master Excel", type=["xlsx", "xls"], key="v_uploader")
            if v_file is not None:
                if st.button("🚀 Process Bulk Vehicle Upload", use_container_width=True):
                    try:
                        uploaded_v_df = pd.read_excel(v_file)
                        uploaded_v_df.columns = uploaded_v_df.columns.str.strip()
                        uploaded_v_df.to_excel(vehicle_master_file, index=False)
                        st.session_state.form_key += 1
                        st.session_state.status_tab_key += 1
                        st.success(f"✅ Bulk Upload Successful! {len(uploaded_v_df)} Vehicles Updated.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error uploading file: {e}")

            st.write("---")
            st.subheader("2. Bulk Dispatch Load Master")
            
            sample_d_df = pd.DataFrame([{
                "Date": datetime.date.today().strftime("%d/%m/%y"),
                "Warehouse / Plant": wh_list[0] if wh_list else "Safedabad",
                "Distributors / DBRs": "Sample DBR 1, Sample DBR 2",
                "Total Load (Ton)": 10.5,
                "Dispatch Status": "Pending",
                "Vehicle Alloted": ""
            }])
            d_template_bytes = generate_excel_download(sample_d_df)
            st.download_button(
                label="📥 Download Dispatch Load Template",
                data=d_template_bytes,
                file_name="dispatch_bulk_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            d_file = st.file_uploader("Upload Bulk Dispatch Excel", type=["xlsx", "xls"], key="d_uploader")
            if d_file is not None:
                if st.button("🚀 Append Bulk Dispatch Entries", use_container_width=True):
                    try:
                        uploaded_d_df = pd.read_excel(d_file)
                        uploaded_d_df.columns = uploaded_d_df.columns.str.strip()
                        existing_d = get_backend_excel_data()
                        final_d = pd.concat([existing_d, uploaded_d_df], ignore_index=True)
                        final_d.to_excel(backend_excel, index=False)
                        st.session_state.form_key += 1
                        st.session_state.status_tab_key += 1
                        st.success(f"✅ {len(uploaded_d_df)} New Dispatch Records Added!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error in Dispatch Upload: {e}")

    # Route Calculation logic
    selected_coords = []
    selected_names = []
    
    w_r = df[df[name_col] == entry_wh].iloc[0]
    selected_coords.append((float(w_r[lat_col]), float(w_r[lon_col])))
    selected_names.append(w_r[name_col])
    center_lat, center_lon = float(w_r[lat_col]), float(w_r[lon_col])
    
    for d_name in entry_dbrs:
        r = df[df[name_col] == d_name].iloc[0]
        selected_coords.append((float(r[lat_col]), float(r[lon_col])))
        selected_names.append(r[name_col])

    min_road_distance = 0
    best_road_geometry = None
    best_route_indices = [0] + list(range(1, len(selected_names))) + [0]
    
    if len(entry_dbrs) > 0:
        min_road_distance = float('inf')
        dbr_num = len(entry_dbrs)
        for perm in itertools.permutations(range(1, dbr_num + 1)):
            current_perm_indices = [0] + list(perm) + [0]
            current_perm_coords = [selected_coords[idx] for idx in current_perm_indices]
            dist, geom = get_road_route_and_distance(current_perm_coords)
            
            if dist is not None and dist < min_road_distance:
                min_road_distance = dist
                best_route_indices = current_perm_indices
                best_road_geometry = geom

        if not best_road_geometry:
            min_road_distance = 0
            best_road_geometry = [selected_coords[idx] for idx in best_route_indices]

    route_names = [selected_names[idx] for idx in best_route_indices] if len(entry_dbrs) > 0 else []
    
    stop_order_dict = {}
    for step, name in enumerate(route_names):
        if step != 0 and step != len(route_names) - 1:
            if name not in stop_order_dict:
                stop_order_dict[name] = len(stop_order_dict) + 1

    with right_col:
        st.subheader("🗺️ Live Route Calculation & Map")
        
        if len(entry_dbrs) > 0:
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("Total Distance", f"{round(min_road_distance, 2)} KM")
            m_col2.metric("Total Load", f"{entry_load} Ton")

        # Folium Map
        m = folium.Map(location=(center_lat, center_lon), zoom_start=11)
        
        # Tile Layers
        folium.TileLayer('OpenStreetMap', name='Street View').add_to(m)
        folium.TileLayer(
            tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
            attr='Esri',
            name='Satellite View'
        ).add_to(m)

        # Fullscreen & Layer Control
        Fullscreen(position="topleft").add_to(m)
        folium.LayerControl(position="topright").add_to(m)
        
        for _, row in df.iterrows():
            loc_name = row[name_col]
            is_wh = str(row[type_col]).upper() in ['WAREHOUSE', 'PLANT', 'WH']
            
            if is_wh:
                icon_html = """
                <div style="background-color: #e74c3c; border: 3px solid #ffffff; border-radius: 50%; width: 42px; height: 42px; display: flex; justify-content: center; align-items: center; box-shadow: 0px 4px 10px rgba(0,0,0,0.4);">
                    <span style="font-size: 20px; color: white;">🏭</span>
                </div>
                """
                marker_icon = folium.DivIcon(html=icon_html, icon_size=(42, 42), icon_anchor=(21, 21))
                
                if loc_name == entry_wh:
                    label_html = f"<div style='font-size: 12px; font-weight: bold; color: white; background: #e74c3c; padding: 2px 5px; border-radius: 4px; border: 2px solid white; white-space: nowrap;'>🏭 START: {loc_name}</div>"
                    folium.map.Marker((row[lat_col], row[lon_col]), icon=folium.features.DivIcon(html=label_html, icon_size=(0,0), icon_anchor=(-20,35))).add_to(m)
                
                folium.Marker(location=(row[lat_col], row[lon_col]), icon=marker_icon).add_to(m)
                
            else:
                if loc_name in entry_dbrs:
                    color, icon = 'blue', 'shopping-cart'
                    stop_num = stop_order_dict.get(loc_name, 1)
                    label_html = f"<div style='font-size: 12px; font-weight: bold; color: #d93838; background: #fff0f0; padding: 2px 5px; border-radius: 4px; border: 2.5px solid #d93838; white-space: nowrap;'>🚨 [{stop_num}] {loc_name}</div>"
                else:
                    color, icon = 'gray', 'info-sign'
                    label_html = f"<div style='font-size: 10px; font-weight: bold; color: #333; background: #ffffff; padding: 2px 4px; border: 1px solid #999; border-radius: 3px; white-space: nowrap;'>{loc_name}</div>"
            
                folium.Marker(location=(row[lat_col], row[lon_col]), icon=folium.Icon(color=color, icon=icon)).add_to(m)
                folium.map.Marker((row[lat_col], row[lon_col]), icon=folium.features.DivIcon(html=label_html, icon_size=(0,0), icon_anchor=(-10,15))).add_to(m)
                
        if len(entry_dbrs) > 0 and best_road_geometry:
            line = folium.PolyLine(best_road_geometry, color="#1b4fd2", weight=6, opacity=0.85).add_to(m)
            PolyLineTextPath(
                line, '                ►                ', repeat=True, offset=8,
                attributes={'fill': '#ffffff', 'font-weight': 'bold', 'font-size': '14px'} 
            ).add_to(m)
        
        st_folium(m, width=850, height=520, key="right_full_map", returned_objects=[])

    # =========================================================
    # 🌟 BOTTOM PANEL: MASTER TABLES
    # =========================================================
    st.write("---")
    t_bottom1, t_bottom2 = st.tabs(["📊 Saved Dispatch Entries Master", "🚚 Vehicle Master Status Sheet"])
    
    with t_bottom1:
        current_excel_df = get_backend_excel_data()
        st.dataframe(current_excel_df, use_container_width=True, height=220)

    with t_bottom2:
        v_sheet = load_vehicle_master()
        st.dataframe(v_sheet, use_container_width=True, height=220)