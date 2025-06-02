import pandas as pd
import json
import urllib.request
from datetime import datetime

def fmt_dt(dt):
    """Formatea datetime o str a 'YYYY-MM-DD HH:MM:SS.SSS'; devuelve '' si NaN."""
    if pd.isna(dt):
        return ""
    if not isinstance(dt, datetime):
        dt = pd.to_datetime(dt)
    return dt.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

# 1) Cargo todas las hojas
xls            = pd.ExcelFile('assets_limpio.xlsx')
assets_df      = pd.read_excel(xls, sheet_name='Assets')
hostnames_df   = pd.read_excel(xls, sheet_name='HostNames')
interfaces_df  = pd.read_excel(xls, sheet_name='Interfaces')
ips_df         = pd.read_excel(xls, sheet_name='IPs')
cpes_df        = pd.read_excel(xls, sheet_name='CPEs')
usuarios_df    = pd.read_excel(xls, sheet_name='Usuarios')

# 2) Inicializo contadores
asset_ctr = 1001
iface_ctr = 2001
ip_ctr    = 3001
user_ctr  = 1
cpe_ctr   = 4001

output = {"ausers": [], "assets": []}

# --- a) Usuarios ---
for _, u in usuarios_df.dropna(subset=['Usuario']).iterrows():
    ts = fmt_dt(u.get('FirstSeen'))
    output["ausers"].append({
        "auser_id":            user_ctr,
        "customer_id":         101,
        "auser_username":      u.get("Usuario", ""),
        "auser_first_seen_dt": ts,
        "auser_first_seen_by": "system",
        "auser_last_seen_dt":  ts,
        "auser_last_seen_by":  "audit_scan",
        "created_dt":          ts,
        "updated_dt":          ts,
    })
    user_ctr += 1

# --- b) Activos ---
for _, a in assets_df.iterrows():
    uid      = a.get('Nombre único', '')
    acq_date = fmt_dt(a.get('Fecha adquisición'))

    # Alias desde HostNames
    row_h = hostnames_df[hostnames_df.get('Nombre único de Asset', '') == uid]
    alias = row_h.get('Nombre del host', pd.Series()).iat[0] if not row_h.empty else ""

    asset = {
        "AssetId":              asset_ctr,
        "customer_id":          101,
        "asset_unique_id":      uid,
        "exple_id":             a.get("Nivel de exposición", None),
        "safie_id":             a.get("Safety", None),
        "zone_id":              a.get("Zona", None),
        "Name":                 a.get("Nombre", ""),
        "asset_alias":          alias,
        "asset_desc":           a.get("Descripción", ""),
        "Type":                 a.get("Tipo", ""),
        "Status":               a.get("Estado", ""),
        "Value":                a.get("Valor económico", 0),
        "AcquisitionDate":      acq_date,
        "Location":             a.get("Localización", ""),
        "AdditionalInfo":       a.get("Información adicional", ""),
        "Owner":                a.get("Propietario", ""),
        "Owner_contact":        a.get("Contacto del propietario", ""),
        "asset_business_owner":         a.get("Propietario (negocio)", ""),
        "asset_business_owner_contact": a.get("Contacto del propietario (negocio)", ""),
        "asset_tech_owner":             a.get("Propietario (Técnico)", ""),
        "asset_tech_owner_contact":     a.get("Contacto del propietario (Técnico)", ""),
        "asset_risk_score_sum":         a.get("Impact", 0),
        "asset_last_reported_dt":       acq_date,
        "asset_impact_confidentiality": a.get("Impact_conf", 0),
        "asset_impact_integrity":       a.get("Impact_int", 0),
        "asset_impact_availability":    a.get("Impact_avail", 0),
        "asset_impact_criticality":     a.get("Impact_crit", 0),
        "socia_risk":                   a.get("Safety_label", a.get("Safety", "")),
        "socia_risk_updated_dt":        acq_date,
        "created_dt":                   acq_date,
        "updated_dt":                   acq_date,
        "Interfaces": [],
        "CPEs": []
    }

    # --- c) Interfaces + IPs ---
    for _, itf in interfaces_df[interfaces_df.get('Nombre único de Asset', '') == uid].iterrows():
        mac = itf.get('Dirección MAC', '')
        iface = {
            "InterfaceId":     iface_ctr,
            "asint_unique_id": mac,
            "MAC":             mac,
            "first_seen_dt":   acq_date,
            "first_seen_by":   "deployment",
            "last_seen_dt":    acq_date,
            "last_seen_by":    "monitoring",
            "created_dt":      acq_date,
            "updated_dt":      acq_date,
            "IPs": []
        }
        for _, ipr in ips_df[ips_df.get('Dirección MAC', '') == mac].iterrows():
            ip_obj = {
                "IpId":                ip_ctr,
                "ainip_unique_id":     f"IP-{uid}-{mac}",
                "Type":                ipr.get("Tipo", "IPv4"),
                "Value":               ipr.get("Dirección IP", ""),
                "ainip_first_seen_dt": acq_date,
                "ainip_first_seen_by": "deployment",
                "ainip_last_seen_dt":  acq_date,
                "ainip_last_seen_by":  "monitoring",
                "created_dt":          acq_date,
                "updated_dt":          acq_date,
            }
            iface["IPs"].append(ip_obj)
            ip_ctr += 1

        asset["Interfaces"].append(iface)
        iface_ctr += 1

    # --- d) CPEs dentro de cada asset ---
    for _, c in cpes_df[cpes_df.get('Nombre único de Asset', '') == uid].iterrows():
        asset["CPEs"].append({
            "ascpe_id":    cpe_ctr,
            "ascpe_value": c.get("CPE", ""),
            "created_dt":  acq_date,
            "updated_dt":  acq_date
        })
        cpe_ctr += 1

    output["assets"].append(asset)
    asset_ctr += 1

# 3) Guardar JSON localmente
with open('resultado.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, indent=2, ensure_ascii=False)

print("✅ JSON generado: resultado.json")

# 4) Enviar JSON a la API con customer_id=3 utilizando urllib
url = "http://127.0.0.1:8000/dev/socia/assets/ingest?customer_id=3"
data_bytes = json.dumps(output).encode('utf-8')
req = urllib.request.Request(url, data=data_bytes, headers={"Content-Type": "application/json"}, method="POST")

try:
    with urllib.request.urlopen(req) as response:
        status = response.getcode()
        body = response.read().decode('utf-8')
        print(f"Response status: {status}")
        print(f"Response body:   {body}")
except urllib.error.URLError as e:
    print(f"Error al enviar a la API: {e}")
