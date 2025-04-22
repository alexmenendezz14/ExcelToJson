import pandas as pd
import json
import subprocess
from datetime import datetime, timezone

# Leer archivo Excel
excel_file = 'assets_limpio.xlsx'
xls = pd.ExcelFile(excel_file)

# Convertir fechas a formato ISO
def to_iso_format(date):
    if pd.isnull(date):
        return datetime.now(timezone.utc).isoformat()
    if isinstance(date, datetime):
        return date.astimezone(timezone.utc).isoformat()
    return pd.to_datetime(date).astimezone(timezone.utc).isoformat()

# Crear estructura JSON
json_data = {"ausers": [], "assets": []}

# Cargar hojas
assets_df = pd.read_excel(xls, sheet_name='Assets')
usuarios_df = pd.read_excel(xls, sheet_name='Usuarios')
interfaces_df = pd.read_excel(xls, sheet_name='Interfaces')
ips_df = pd.read_excel(xls, sheet_name='IPs')
cpes_df = pd.read_excel(xls, sheet_name='CPEs')
propiedades_df = pd.read_excel(xls, sheet_name='Propiedades')

# IDs incrementales
user_id_counter = 1
asset_id_counter = 1001
asaus_id_counter = 5001
asint_id_counter = 2001
ainip_id_counter = 3001
ascpe_id_counter = 4001
aspro_id_counter = 6001

# Usuarios únicos
usuarios_unicos = usuarios_df['Usuario'].unique()
for usuario in usuarios_unicos:
    json_data["ausers"].append({
        "auser_id": user_id_counter,
        "customer_id": 101,
        "auser_username": usuario,
        "auser_first_seen_dt": datetime.now(timezone.utc).isoformat(),
        "auser_first_seen_by": "system",
        "auser_last_seen_dt": datetime.now(timezone.utc).isoformat(),
        "auser_last_seen_by": "audit_scan",
        "created_dt": datetime.now(timezone.utc).isoformat(),
        "updated_dt": datetime.now(timezone.utc).isoformat()
    })
    user_id_counter += 1

# Assets
for _, asset_row in assets_df.iterrows():
    asset_unique_id = asset_row.get("Nombre único")
    asset = {
        "asset_id": asset_id_counter,
        "customer_id": 101,
        "asset_unique_id": asset_unique_id,
        "exple_id": asset_row.get("Nivel de exposición"),
        "safie_id": asset_row.get("Safety"),
        "zone_id": asset_row.get("Zona"),
        "asset_name": asset_row.get("Nombre"),
        "asset_alias": asset_row.get("Alias"),
        "asset_desc": asset_row.get("Descripción"),
        "asset_type": asset_row.get("Tipo"),
        "asset_status": asset_row.get("Estado"),
        "asset_value": asset_row.get("Valor económico"),
        "asset_acquisition_dt": to_iso_format(asset_row.get("Fecha adquisición")),
        "asset_location": asset_row.get("Localización"),
        "asset_additional_info": asset_row.get("Información adicional"),
        "asset_owner": asset_row.get("Propietario"),
        "asset_owner_contact": asset_row.get("Contacto del propietario"),
        "asset_business_owner": asset_row.get("Propietario (negocio)"),
        "asset_business_owner_contact": asset_row.get("Contacto del propietario (negocio)"),
        "asset_tech_owner": asset_row.get("Propietario (Técnico)"),
        "asset_tech_owner_contact": asset_row.get("Contacto del propietario (Técnico)"),
        "asset_risk_score_sum": 65,
        "asset_last_reported_dt": datetime.now(timezone.utc).isoformat(),
        "asset_impact_confidentiality": 3,
        "asset_impact_integrity": 4,
        "asset_impact_availability": 5,
        "asset_impact_criticality": 4,
        "socia_risk": "MEDIO",
        "socia_risk_updated_dt": datetime.now(timezone.utc).isoformat(),
        "created_dt": datetime.now(timezone.utc).isoformat(),
        "updated_dt": datetime.now(timezone.utc).isoformat(),
        "assets_ausers": [],
        "assets_interfaces": [],
        "assets_cpes": [],
        "assets_properties": []
    }

    asset_id_counter += 1

    # Assets_Usuarios
    usuarios_asset = usuarios_df[usuarios_df['Nombre único de Asset'] == asset_unique_id]
    for _, user_asset_row in usuarios_asset.iterrows():
        asset["assets_ausers"].append({
            "asaus_id": asaus_id_counter,
            "auser_id": next(u["auser_id"] for u in json_data["ausers"] if u["auser_username"] == user_asset_row["Usuario"]),
            "asaus_first_seen_dt": datetime.now(timezone.utc).isoformat(),
            "asaus_first_seen_by": "scan",
            "asaus_last_seen_dt": datetime.now(timezone.utc).isoformat(),
            "asaus_last_seen_by": "monitor",
            "created_dt": datetime.now(timezone.utc).isoformat(),
            "updated_dt": datetime.now(timezone.utc).isoformat()
        })
        asaus_id_counter += 1

    # Interfaces y IPs
    interfaces = interfaces_df[interfaces_df['Nombre único de Asset'] == asset_unique_id]
    for _, interface_row in interfaces.iterrows():
        interface = {
            "asint_id": asint_id_counter,
            "asint_unique_id": interface_row.get("Dirección MAC"),
            "mac_address": interface_row.get("Dirección MAC"),
            "first_seen_dt": datetime.now(timezone.utc).isoformat(),
            "first_seen_by": "deployment",
            "last_seen_dt": datetime.now(timezone.utc).isoformat(),
            "last_seen_by": "monitoring",
            "created_dt": datetime.now(timezone.utc).isoformat(),
            "updated_dt": datetime.now(timezone.utc).isoformat(),
            "assets_interfaces_ips": []
        }
        asint_id_counter += 1

        ips = ips_df[ips_df['Dirección MAC'] == interface["mac_address"]]
        for _, ip_row in ips.iterrows():
            interface["assets_interfaces_ips"].append({
                "ainip_id": ainip_id_counter,
                "ainip_unique_id": f"IP-{asset_unique_id}-{interface['mac_address']}",
                "ainip_type": "IPv4",
                "ainip_value": ip_row.get("Dirección IP"),
                "ainip_first_seen_dt": datetime.now(timezone.utc).isoformat(),
                "ainip_first_seen_by": "deployment",
                "ainip_last_seen_dt": datetime.now(timezone.utc).isoformat(),
                "ainip_last_seen_by": "monitoring",
                "created_dt": datetime.now(timezone.utc).isoformat(),
                "updated_dt": datetime.now(timezone.utc).isoformat()
            })
            ainip_id_counter += 1

        asset["assets_interfaces"].append(interface)

    json_data["assets"].append(asset)

# Guardar JSON
with open('resultado.json', 'w', encoding='utf-8') as f:
    json.dump(json_data, f, indent=2, ensure_ascii=False)

print("JSON generado con éxito: resultado.json")


def subir_a_github():
    try:
        # Agregar archivo
        subprocess.run("git add resultado.json", shell=True, check=True)

        # Commit con mensaje
        commit_msg = f'git commit -m "Auto-update: JSON generado el {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}"'
        subprocess.run(commit_msg, shell=True, check=True)

        # Push
        subprocess.run("git push", shell=True, check=True)

        print("✅ JSON subido a GitHub correctamente.")
    except subprocess.CalledProcessError as e:
        print("❌ Error al subir el JSON a GitHub:", e)


subir_a_github()


