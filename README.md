**Descripción**
Este script en Python tiene como objetivo procesar datos provenientes de un archivo Excel (`assets_limpio.xlsx`), estructurarlos en un formato JSON y luego enviarlos a una API REST para ingesta de información relacionada con usuarios, activos, interfaces, direcciones IP y CPEs.

---

**Requisitos**

* Python 3.7 o superior
* Bibliotecas de Python:

  * `pandas`
  * `json` (incluida en la librería estándar)
  * `urllib` (incluida en la librería estándar)
  * `datetime` (incluida en la librería estándar)

Para instalar `pandas`, ejecutar:

```bash
pip install pandas
```

---

**Estructura de Archivos**

* `assets_limpio.xlsx`

  * Hoja **Assets**
  * Hoja **HostNames**
  * Hoja **Interfaces**
  * Hoja **IPs**
  * Hoja **CPEs**
  * Hoja **Usuarios**
* `script.py` (el código que se describe a continuación)
* `resultado.json` (archivo JSON que genera el script)

---

**Instalación / Preparación**

1. Clonar o copiar este repositorio en tu máquina local.
2. Colocar el archivo `assets_limpio.xlsx` en la misma carpeta que `script.py`.
3. Asegurarse de tener instalado Python 3.7+ y la biblioteca `pandas`.

---

**Uso**

1. Abrir una terminal y navegar hasta la carpeta donde están `script.py` y `assets_limpio.xlsx`.
2. Ejecutar el script:

   ```bash
   python script.py
   ```
3. El flujo general que realiza el script es:

   * Carga todas las hojas de Excel mencionadas.
   * Formatea fechas y horas al formato `YYYY-MM-DD HH:MM:SS.SSS`, o devuelve cadena vacía si falta el valor.
   * Recorre la hoja de usuarios para generar objetos JSON con campos como `auser_id`, `auser_username`, fechas de “first\_seen” y “last\_seen”, etc.
   * Recorre la hoja de activos (“Assets”), obtiene datos generales por cada activo e identifica alias de host en la hoja “HostNames”.
   * Para cada activo, construye lista de interfaces a partir de la hoja “Interfaces”.
   * Para cada interfaz, recopila direcciones IP de la hoja “IPs” y las agrega bajo la clave `IPs` de la interfaz.
   * Para cada activo, recopila registros CPE de la hoja “CPEs” y los agrega en una lista bajo la clave `CPEs`.
   * Incrementa contadores internos (`asset_ctr`, `iface_ctr`, `ip_ctr`, `user_ctr`, `cpe_ctr`) para asignar identificadores únicos.
   * Arma un diccionario Python con dos claves principales:

     * `"ausers"`: lista de usuarios procesados.
     * `"assets"`: lista de activos, cada uno con sus interfaces, IPs y CPEs.
   * Guarda este diccionario en un archivo local llamado `resultado.json` con codificación UTF-8 e indentado.
   * Intenta enviar el JSON resultante a la API REST en la URL:

     ```
     http://127.0.0.1:8000/dev/socia/assets/ingest?customer_id=3
     ```
   * Imprime en consola un mensaje de confirmación (“✅ JSON generado: resultado.json”).
   * Luego, muestra el estado HTTP y la respuesta de la API o un mensaje de error si la conexión falla.

---

**Explicación de Funciones y Bloques Principales**

1. **Función `fmt_dt(dt)`**

   * Recibe un objeto `datetime`, una cadena de texto o `NaN`.
   * Si el valor es `NaN`, devuelve `""`.
   * Si no es instancia de `datetime`, lo convierte usando `pd.to_datetime()`.
   * Formatea la fecha como `"YYYY-MM-DD HH:MM:SS.SSS"`, truncando los microsegundos a milisegundos.

2. **Carga de Hojas de Excel**

   ```python
   xls            = pd.ExcelFile('assets_limpio.xlsx')
   assets_df      = pd.read_excel(xls, sheet_name='Assets')
   hostnames_df   = pd.read_excel(xls, sheet_name='HostNames')
   interfaces_df  = pd.read_excel(xls, sheet_name='Interfaces')
   ips_df         = pd.read_excel(xls, sheet_name='IPs')
   cpes_df        = pd.read_excel(xls, sheet_name='CPEs')
   usuarios_df    = pd.read_excel(xls, sheet_name='Usuarios')
   ```

   * Se crea un objeto `ExcelFile` para poder leer cada hoja por su nombre.
   * Cada `*_df` es un DataFrame de pandas que contiene los datos de la respectiva hoja.

3. **Inicialización de Contadores**

   ```python
   asset_ctr = 1001
   iface_ctr = 2001
   ip_ctr    = 3001
   user_ctr  = 1
   cpe_ctr   = 4001
   ```

   * Se utilizan para asignar identificadores únicos a usuarios, activos, interfaces, IPs y CPEs.

4. **Procesamiento de Usuarios**

   ```python
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
   ```

   * Filtra filas donde la columna `Usuario` NO sea nula.
   * Por cada fila, formatea la fecha `FirstSeen` y arma un diccionario con campos que incluyen:

     * `auser_id`: identificador incremental.
     * `customer_id`: fijo en 101 (puede ajustarse según necesidad).
     * `auser_username`: nombre de usuario.
     * Fechas de “first\_seen” y “last\_seen” ambos iguales a `FirstSeen`.
     * Metadatos de “first\_seen\_by” y “last\_seen\_by” con valores fijos `system` y `audit_scan`.
     * Campos `created_dt` y `updated_dt` también con la misma marca de tiempo.

5. **Procesamiento de Activos**

   ```python
   for _, a in assets_df.iterrows():
       uid      = a.get('Nombre único', '')
       acq_date = fmt_dt(a.get('Fecha adquisición'))
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
       ...
       output["assets"].append(asset)
       asset_ctr += 1
   ```

   * Por cada fila de la hoja “Assets”, se extraen campos básicos (nombre, descripción, valor, dueño, impactos, etc.).
   * Se busca en la hoja “HostNames” el alias de host correspondiente usando `Nombre único de Asset`. Si existe, se extrae el primer valor de `Nombre del host`.
   * Se prepara un diccionario `asset` con todos los campos relevantes (IDs, metadatos, fechas).
   * Dentro del mismo bucle, se invocan dos subprocesos:

     * **Interfaces + Direcciones IP** (parte c).
     * **CPEs asociados** (parte d).
   * Finalmente, se agrega el objeto `asset` completo a `output["assets"]` y se incrementa `asset_ctr`.

6. **Interfaces + Direcciones IP**

   ```python
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
   ```

   * Filtra la hoja “Interfaces” para aquellas filas cuyo `Nombre único de Asset` coincida con `uid` del activo actual.
   * Para cada interfaz:

     * Se toma la dirección MAC y se crea un diccionario `iface` con metadatos de “first\_seen”, “last\_seen”, etc.
     * Se buscan en la hoja “IPs” todas las filas cuya `Dirección MAC` sea igual a la MAC actual.
     * Por cada entrada IP, se genera un objeto con:

       * `IpId` incremental.
       * `ainip_unique_id`: cadena formada por `IP-{uid}-{mac}`.
       * Tipo de IP (por defecto "IPv4").
       * Valor de la IP.
       * Fechas y metadatos idénticos a los de la interfaz.
     * Se agrega cada `ip_obj` a la lista `iface["IPs"]`.
   * Al terminar de procesar las IPs de esta interfaz, se agrega `iface` a `asset["Interfaces"]` y se incrementa `iface_ctr`.

7. **CPEs dentro de cada Activo**

   ```python
   for _, c in cpes_df[cpes_df.get('Nombre único de Asset', '') == uid].iterrows():
       asset["CPEs"].append({
           "ascpe_id":    cpe_ctr,
           "ascpe_value": c.get("CPE", ""),
           "created_dt":  acq_date,
           "updated_dt":  acq_date
       })
       cpe_ctr += 1
   ```

   * Filtra la hoja “CPEs” para filas cuyo `Nombre único de Asset` coincida con `uid`.
   * Por cada entrada, agrega un diccionario con:

     * `ascpe_id`: identificador incremental.
     * `ascpe_value`: valor de la columna “CPE”.
     * Fechas `created_dt` y `updated_dt` (mismo valor de `acq_date` del activo).
   * Incrementa `cpe_ctr` para el siguiente CPE.

8. **Generación y Guardado de JSON**

   ```python
   with open('resultado.json', 'w', encoding='utf-8') as f:
       json.dump(output, f, indent=2, ensure_ascii=False)
   print("✅ JSON generado: resultado.json")
   ```

   * El diccionario `output` (con claves `"ausers"` y `"assets"`) se escribe en un archivo llamado `resultado.json`, con indentación de 2 espacios y sin escape de caracteres Unicode.
   * Se notifica en consola que el JSON se generó exitosamente.

9. **Envío del JSON a la API**

   ```python
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
   ```

   * Define la URL de la API local (puerto 8000) con parámetro `customer_id=3`.
   * Convierte el diccionario `output` a bytes JSON (`utf-8`).
   * Crea un objeto `Request` configurado para una petición HTTP `POST` con cabecera `Content-Type: application/json`.
   * Intenta abrir la URL, leer el código de estado HTTP y el cuerpo de respuesta, e imprimirlos.
   * Si hay un error de red (por ejemplo, la API no está levantada), captura `URLError` y lo muestra en consola.

---

**Variables y Contadores**

* `asset_ctr`, `iface_ctr`, `ip_ctr`, `user_ctr`, `cpe_ctr`: valores iniciales fijos para simular identificadores únicos.
* `output`: diccionario principal que contendrá dos listas:

  * `"ausers"`: usuarios procesados.
  * `"assets"`: activos procesados (cada uno con sus interfaces, IPs y CPEs).

---

**Puntos a Tener en Cuenta**

* El script asume que las hojas del archivo Excel están correctamente nombradas y contienen las columnas esperadas.
* Si alguna hoja o columna no existe, el script lanzará un error de `KeyError`.
* La función de formato de fechas (`fmt_dt`) convierte valores nulos en cadenas vacías, para evitar errores al formatear.
* En el bloque de envío a la API, se usa `customer_id=3` aunque en la construcción de objetos se está usando `customer_id=101`. Ajustar según corresponda.
* Si la API no está disponible en `127.0.0.1:8000`, la petición fallará y se imprimirá el error en consola.

---

**Ejemplo de Estructura JSON de Salida**

```jsonc
{
  "ausers": [
    {
      "auser_id": 1,
      "customer_id": 101,
      "auser_username": "usuario1",
      "auser_first_seen_dt": "2023-05-10 12:34:56.000",
      "auser_first_seen_by": "system",
      "auser_last_seen_dt": "2023-05-10 12:34:56.000",
      "auser_last_seen_by": "audit_scan",
      "created_dt": "2023-05-10 12:34:56.000",
      "updated_dt": "2023-05-10 12:34:56.000"
    },
    ...
  ],
  "assets": [
    {
      "AssetId": 1001,
      "customer_id": 101,
      "asset_unique_id": "ACT-001",
      "exple_id": 2,
      "safie_id": "High",
      "zone_id": "Zona A",
      "Name": "Servidor Principal",
      "asset_alias": "srv-prin-01",
      "asset_desc": "Servidor de base de datos principal",
      "Type": "Servidor",
      "Status": "En Producción",
      "Value": 15000,
      "AcquisitionDate": "2022-01-15 00:00:00.000",
      "Location": "CPD Central",
      "AdditionalInfo": "Instalado en rack 3",
      "Owner": "IT Manager",
      "Owner_contact": "it.manager@ejemplo.com",
      "asset_business_owner": "CIO",
      "asset_business_owner_contact": "cio@ejemplo.com",
      "asset_tech_owner": "Administrador de Sistemas",
      "asset_tech_owner_contact": "sysadmin@ejemplo.com",
      "asset_risk_score_sum": 7,
      "asset_last_reported_dt": "2022-01-15 00:00:00.000",
      "asset_impact_confidentiality": 3,
      "asset_impact_integrity": 2,
      "asset_impact_availability": 2,
      "asset_impact_criticality": 3,
      "socia_risk": "Medium",
      "socia_risk_updated_dt": "2022-01-15 00:00:00.000",
      "created_dt": "2022-01-15 00:00:00.000",
      "updated_dt": "2022-01-15 00:00:00.000",
      "Interfaces": [
        {
          "InterfaceId": 2001,
          "asint_unique_id": "00:1A:2B:3C:4D:5E",
          "MAC": "00:1A:2B:3C:4D:5E",
          "first_seen_dt": "2022-01-15 00:00:00.000",
          "first_seen_by": "deployment",
          "last_seen_dt": "2022-01-15 00:00:00.000",
          "last_seen_by": "monitoring",
          "created_dt": "2022-01-15 00:00:00.000",
          "updated_dt": "2022-01-15 00:00:00.000",
          "IPs": [
            {
              "IpId": 3001,
              "ainip_unique_id": "IP-ACT-001-00:1A:2B:3C:4D:5E",
              "Type": "IPv4",
              "Value": "192.168.1.10",
              "ainip_first_seen_dt": "2022-01-15 00:00:00.000",
              "ainip_first_seen_by": "deployment",
              "ainip_last_seen_dt": "2022-01-15 00:00:00.000",
              "ainip_last_seen_by": "monitoring",
              "created_dt": "2022-01-15 00:00:00.000",
              "updated_dt": "2022-01-15 00:00:00.000"
            },
            ...
          ]
        },
        ...
      ],
      "CPEs": [
        {
          "ascpe_id": 4001,
          "ascpe_value": "cpe:/o:microsoft:windows_10:1709",
          "created_dt": "2022-01-15 00:00:00.000",
          "updated_dt": "2022-01-15 00:00:00.000"
        },
        ...
      ]
    },
    ...
  ]
}
```

> **Nota:** Los valores mostrados en este ejemplo son ilustrativos.

---

**Cómo Adaptarlo a Otro Entorno**

* Si la API de destino no está en `127.0.0.1:8000`, modificar la variable `url` al host y puerto adecuados.
* Si el parámetro `customer_id` debe ser otro número, cambiar `?customer_id=3` en la URL.
* Ajustar los IDs iniciales (`asset_ctr`, `iface_ctr`, etc.) según la convención de tu base de datos o sistema de numeración.
* Si las hojas de Excel tienen columnas con nombres distintos, actualizar los `get("Nombre de columna", ...)` a los nombres reales.

---

**Licencia**
Este proyecto se distribuye bajo licencia [MIT](https://opensource.org/licenses/MIT) (o la que corresponda), permitiendo uso, modificación y distribución libres, siempre que se mantenga la atribución y se incluya un archivo de licencia.

---

**Autor**

* Nombre del desarrollador: *Tu Nombre Aquí*
* Contacto: *[tu.email@ejemplo.com](mailto:tu.email@ejemplo.com)*

---

Con este README tendrás toda la información necesaria para entender, ejecutar y adaptar el script a tus necesidades. ¡Éxito!
