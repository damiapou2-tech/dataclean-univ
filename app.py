import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from xlsxwriter.utility import xl_col_to_name
import io
import re

st.set_page_config(page_title="DataClean Pro | Automatització d'Informes", page_icon="📊", layout="wide")

st.title("📊 DataClean Pro — Neteja i Automatització d'Informes")
st.markdown("Compatible amb **Excel (.xlsx, .xls), CSV, TSV i ODS**. Detecta pestanyes, codificacions, formats europeus i genera gràfiques i fórmules executives automàticament.")

FORMAT_ADMESOS = ["xlsx", "xls", "ods", "csv", "tsv", "txt"]

def netejar_nom_arxiu(text):
    """Genera un nom d'arxiu segur per a Windows/Mac."""
    text_net = re.sub(r'[\\/*?:"<>|\s]+', '_', str(text)).strip('_')
    return text_net[:80] if text_net else "consolidat"

def deduir_etiqueta_global(nom_fitxer, llista_pestanyes):
    """Deduïx automàticament l'etiqueta del fitxer segons les pestanyes triades."""
    nom_base = nom_fitxer.rsplit(".", 1)[0]
    n = len(llista_pestanyes)
    if n == 0:
        return netejar_nom_arxiu(nom_base)
    if n == 1:
        return netejar_nom_arxiu(llista_pestanyes[0])
    
    mapa_periodes = {
        2: "Bimestral",
        3: "Trimestral",
        4: "Quadrimestral",
        6: "Semestral",
        12: "Anual"
    }
    ambit = mapa_periodes.get(n, f"Consolidat_{n}_Fulls")
    return netejar_nom_arxiu(f"{ambit}_{nom_base}")

def detectar_fila_capcalera(df_sense_cap):
    """Escaneja les primeres 20 files i retorna l'índex on comencen les capçaleres reals."""
    paraules_clau = ("codi", "ref", "sku", "article", "quant", "data", "fecha", "factura", "tiquet", "ticket", "preu", "pvp", "import", "client")
    millor_fila = 0
    max_coincidencies = 0

    for idx in range(min(20, len(df_sense_cap))):
        valors_fila = [str(v).strip().lower() for v in df_sense_cap.iloc[idx].values if pd.notna(v)]
        coincidencies = sum(1 for val in valors_fila if any(k in val for k in paraules_clau))
        if coincidencies > max_coincidencies:
            max_coincidencies = coincidencies
            millor_fila = idx

    return millor_fila if max_coincidencies >= 2 else 0

def detectar_separador_csv(text_complet, fila_inici=0):
    """Detecta de forma fiable si el CSV usa ';' o ',' evitant confusions amb decimals europeus."""
    linies = text_complet.splitlines()
    if len(linies) > fila_inici:
        primera_linia = linies[fila_inici]
        if primera_linia.count(";") > 0 and primera_linia.count(";") >= primera_linia.count(","):
            return ";"
        if primera_linia.count("\t") > 0:
            return "\t"
    return ","

@st.cache_data(show_spinner=False)
def analitzar_metadades_fitxer(raw_bytes, nom_fitxer):
    """Cacheja la inspecció inicial del fitxer detectant la capçalera de cada pestanya."""
    ext = nom_fitxer.split(".")[-1].lower()
    if ext in ("xlsx", "xls", "ods"):
        excel_file = pd.ExcelFile(io.BytesIO(raw_bytes))
        pestanyes = excel_file.sheet_names
        files_per_pestanya = {}
        for p in pestanyes:
            df_preview = pd.read_excel(excel_file, sheet_name=p, header=None, nrows=20)
            files_per_pestanya[p] = detectar_fila_capcalera(df_preview)
        return {
            "tipus": "excel",
            "pestanyes": pestanyes,
            "fila_auto": files_per_pestanya[pestanyes[0]],
            "files_per_pestanya": files_per_pestanya
        }
    else:
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text_tmp = raw_bytes.decode(enc)
                sep_tmp = detectar_separador_csv(text_tmp, 0)
                df_preview = pd.read_csv(io.StringIO(text_tmp), sep=sep_tmp, header=None, nrows=20)
                fila_auto = detectar_fila_capcalera(df_preview)
                return {"tipus": "csv", "encoding": enc, "fila_auto": fila_auto}
            except Exception:
                continue
        return {"tipus": "error"}

@st.cache_data(show_spinner=False)
def llegir_dades_cachejades(raw_bytes, nom_fitxer, fila_inici, usar_auto_per_full, pestanyes_seleccionades=None):
    """Llegeix i consolida només les pestanyes seleccionades adaptant la capçalera de cada full."""
    ext = nom_fitxer.split(".")[-1].lower()
    pestanyes_depurades = []
    info_lectura = ""

    if ext in ("xlsx", "xls", "ods"):
        meta = analitzar_metadades_fitxer(raw_bytes, nom_fitxer)
        files_map = meta.get("files_per_pestanya", {})
        excel_file = pd.ExcelFile(io.BytesIO(raw_bytes))
        pestanyes_a_llegir = list(pestanyes_seleccionades) if pestanyes_seleccionades else [excel_file.sheet_names[0]]

        llista_dfs = []
        detall_files = []
        for p in pestanyes_a_llegir:
            skip_p = files_map.get(p, fila_inici) if usar_auto_per_full else fila_inici
            df_temp = pd.read_excel(excel_file, sheet_name=p, skiprows=skip_p)
            if not df_temp.empty:
                df_temp["_origen_pestanya"] = p
                llista_dfs.append(df_temp)
                pestanyes_depurades.append(p)
                detall_files.append(f"{p} (fila {skip_p})")

        df_final = pd.concat(llista_dfs, ignore_index=True) if llista_dfs else pd.DataFrame()
        if len(pestanyes_depurades) == 1:
            p_unica = pestanyes_depurades[0]
            skip_u = files_map.get(p_unica, fila_inici) if usar_auto_per_full else fila_inici
            info_lectura = f"Pestanya seleccionada: **{p_unica}** (inici a fila {skip_u})"
        elif len(pestanyes_depurades) > 1:
            files_distintes = len({files_map.get(p, fila_inici) for p in pestanyes_depurades}) > 1
            if usar_auto_per_full and files_distintes:
                info_lectura = f"{len(pestanyes_depurades)} pestanyes fusionades amb alineació per full: {', '.join(detall_files)}"
            else:
                info_lectura = f"{len(pestanyes_depurades)} pestanyes seleccionades i fusionades correctament."
    else:
        meta = analitzar_metadades_fitxer(raw_bytes, nom_fitxer)
        if meta["tipus"] == "error":
            return pd.DataFrame(), "", "Error de codificació"
        enc = meta["encoding"]
        text_complet = raw_bytes.decode(enc)
        separador = detectar_separador_csv(text_complet, fila_inici)
        df_final = pd.read_csv(io.StringIO(text_complet), sep=separador, skiprows=fila_inici)
        nom_base = nom_fitxer.rsplit(".", 1)[0]
        df_final["_origen_pestanya"] = nom_base
        pestanyes_depurades.append(nom_base)
        info_lectura = f"Arxiu de text llegit correctament (Codificació: **{enc}**, Separador: **`{separador}`**)"

    df_final = df_final.dropna(axis=1, how="all")
    etiqueta_sortida = deduir_etiqueta_global(nom_fitxer, pestanyes_depurades)
    return df_final, etiqueta_sortida, info_lectura

def carregar_fitxer_universal(fitxer, nom_fitxer_label):
    """Interfície d'entrada amb selecció múltiple de pestanyes i detecció de capçaleres variables."""
    raw_bytes = fitxer.getvalue()
    meta = analitzar_metadades_fitxer(raw_bytes, fitxer.name)

    if meta["tipus"] == "error":
        st.error("No s'ha pogut llegir la codificació de l'arxiu CSV.")
        return pd.DataFrame(), ""

    with st.expander(f"⚙️ Configuració d'entrada: {fitxer.name}", expanded=True):
        col_a, col_b = st.columns([2, 1])

        pestanyes_triades = None
        if meta["tipus"] == "excel":
            totes_pestanyes = meta["pestanyes"]
            if len(totes_pestanyes) > 1:
                pestanyes_triades = col_a.multiselect(
                    f"Selecciona les pestanyes a depurar ({len(totes_pestanyes)} disponibles):",
                    options=totes_pestanyes,
                    default=totes_pestanyes,
                    key=f"sheets_{nom_fitxer_label}_{fitxer.name}",
                    help="Pots treure les pestanyes que no vulguis processar o deixar-ne només una."
                )
                if not pestanyes_triades:
                    st.warning("⚠️ Selecciona com a mínim una pestanya per continuar.")
                    return pd.DataFrame(), ""
            else:
                pestanyes_triades = totes_pestanyes

        fila_defecte = int(meta["fila_auto"])
        if meta["tipus"] == "excel" and pestanyes_triades:
            fila_defecte = int(meta["files_per_pestanya"].get(pestanyes_triades[0], fila_defecte))

        clau_skip = f"skip_{nom_fitxer_label}_{fitxer.name}_{'_'.join(pestanyes_triades) if pestanyes_triades else 'csv'}"
        fila_inici = col_b.number_input(
            "Fila d'inici (Auto per full)",
            min_value=0, max_value=50, value=fila_defecte,
            key=clau_skip,
            help="Detecta automàticament la fila de capçalera de cada pestanya. Si la canvies manualment, s'aplicarà el valor fixat."
        )

        usar_auto_per_full = (meta["tipus"] == "excel" and fila_inici == fila_defecte)

        df_final, etiqueta_sortida, info_lectura = llegir_dades_cachejades(
            raw_bytes,
            fitxer.name,
            fila_inici,
            usar_auto_per_full,
            tuple(pestanyes_triades) if pestanyes_triades else None
        )

        if info_lectura:
            col_a.info(info_lectura)

        if df_final.empty:
            st.warning("Les pestanyes seleccionades no contenen dades.")
            return pd.DataFrame(), ""

        cols_mostrar = [c for c in df_final.columns if c != "_origen_pestanya"]
        st.caption(f"Vista prèvia ràpida ({len(df_final)} files x {len(cols_mostrar)} columnes):")
        st.dataframe(df_final[cols_mostrar].head(3), use_container_width=True)

    return df_final, etiqueta_sortida

def netejar_numero_europeu(serie):
    """Conversió vectoritzada ultraràpida de formats europeus ('1.250,50 €') a float net."""
    if pd.api.types.is_numeric_dtype(serie):
        return serie.fillna(0.0).astype(float)
    
    s = serie.fillna("0").astype(str).str.replace(r'[€$£\s]', '', regex=True)
    te_ambdos = s.str.contains(r'\.', regex=True) & s.str.contains(',', regex=False)
    s = s.where(~te_ambdos, s.str.replace('.', '', regex=False))
    s = s.str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

# --- INTERFÍCIE PRINCIPAL ---
c1, c2 = st.columns(2)
with c1:
    file_vendes = st.file_uploader("1. Dades de Vendes / Moviments", type=FORMAT_ADMESOS)
with c2:
    file_cataleg = st.file_uploader("2. Catàleg / Mestres de Productes", type=FORMAT_ADMESOS)

if file_vendes and file_cataleg:
    df_v_raw, etiqueta_defecte = carregar_fitxer_universal(file_vendes, "vendes")
    df_c_raw, _ = carregar_fitxer_universal(file_cataleg, "cataleg")

    if not df_v_raw.empty and not df_c_raw.empty:
        st.divider()
        st.subheader("🔗 Assignació de Columnes (Mapeig Intel·ligent)")
        st.caption("El sistema detecta les columnes automàticament, inclòs l'ID de Venda/Tiquet per distingir duplicats reals de compres seguides.")

        cols_v = [c for c in df_v_raw.columns if c != "_origen_pestanya"]
        cols_c = [c for c in df_c_raw.columns if c != "_origen_pestanya"]

        def trobar_col(llista, paraules_clau, excluir=None):
            excluir = excluir or []
            for idx, col in enumerate(llista):
                col_low = str(col).lower()
                if any(ex in col_low for ex in excluir):
                    continue
                if any(k in col_low for k in paraules_clau):
                    return idx
            return 0

        m1, m2, m3, m4, m5 = st.columns(5)
        col_v_tiquet = m1.selectbox(
            "ID Venda / Tiquet (Clau)",
            ["(Cap)"] + cols_v,
            index=trobar_col(["(Cap)"] + cols_v, ("factura", "tiquet", "ticket", "comanda", "pedido", "doc", "num_"), excluir=["client", "sku"]),
            help="Permet distingir un duplicat real d'exportació de dues vendes seguides del mateix producte."
        )
        col_v_codi = m2.selectbox("Codi Producte (Vendes)", cols_v, index=trobar_col(cols_v, ("sku", "ref", "codi", "art", "prod"), excluir=["factura", "tiquet", "client"]))
        col_v_quant = m3.selectbox("Quantitat (Vendes)", cols_v, index=trobar_col(cols_v, ("quant", "unit", "qty", "cant")))
        col_v_data = m4.selectbox("Data (Vendes)", cols_v, index=trobar_col(cols_v, ("data", "fecha", "date", "dia")))
        col_v_canal = m5.selectbox("Canal / Client (Opcional)", ["(Cap)"] + cols_v, index=trobar_col(["(Cap)"] + cols_v, ("canal", "botiga", "client", "origen")))

        mc1, mc2, mc3, mc4, mc5 = st.columns(5)
        col_c_codi = mc1.selectbox("Codi Producte (Catàleg)", cols_c, index=trobar_col(cols_c, ("codi", "ref", "sku", "id", "prod")))
        col_c_nom = mc2.selectbox("Nom Producte",cols_c,index=trobar_col(cols_c, ("nom", "desc", "article", "producte", "name"), excluir=["codi", "ref", "sku", "id"]))
        col_c_cat = mc3.selectbox("Categoria / Família", cols_c, index=trobar_col(cols_c, ("cat", "fam", "grup", "tipus")))
        col_c_preu = mc4.selectbox("Preu Venda", cols_c, index=trobar_col(cols_c, ("preu", "pvp", "venta", "price", "import")))
        col_c_cost = mc5.selectbox("Cost Unitari", cols_c, index=trobar_col(cols_c, ("cost", "compra", "base")))

        opt1, opt2 = st.columns([2, 1])
        etiqueta_personalitzada = opt1.text_input(
            "🏷️ Etiqueta de l'informe final (detectada automàticament segons les pestanyes, editable):",
            value=etiqueta_defecte
        )
        generar_grafics = opt2.checkbox(
            "📈 Generar les 3 gràfiques executives (Web + Excel)",
            value=True,
            help="Inclou els 3 gràfics interactius a la web i incrustats dins l'Excel amb fórmules dinàmiques."
        )

        if st.button("🚀 Executar Neteja i Generar Informe Executiu", type="primary", use_container_width=True):
            te_id_venda = col_v_tiquet != "(Cap)"

            # 1. Estandardització vectoritzada amb ID de Venda / Tiquet
            df_v = pd.DataFrame({
                "id_venda": df_v_raw[col_v_tiquet].fillna("").astype(str).str.strip().str.upper() if te_id_venda else "SENSE_ID",
                "data": df_v_raw[col_v_data],
                "codi_producte": df_v_raw[col_v_codi].fillna("").astype(str).str.strip().str.upper(),
                "quantitat": netejar_numero_europeu(df_v_raw[col_v_quant]),
                "canal": df_v_raw[col_v_canal] if col_v_canal != "(Cap)" else "General",
                "pestanya_origen": df_v_raw["_origen_pestanya"] if "_origen_pestanya" in df_v_raw.columns else "Principal"
            })

            df_c = pd.DataFrame({
                "codi_producte": df_c_raw[col_c_codi].fillna("").astype(str).str.strip().str.upper(),
                "nom_producte": df_c_raw[col_c_nom],
                "categoria": df_c_raw[col_c_cat].fillna("Sense Categoria"),
                "preu_venda": netejar_numero_europeu(df_c_raw[col_c_preu]),
                "cost_unitari": netejar_numero_europeu(df_c_raw[col_c_cost])
            }).drop_duplicates(subset=["codi_producte"], keep="first")

            # 2. Auditoria de files buides agrupada per full
            es_fila_buida = df_v["codi_producte"].isin({"", "NAN", "NONE", "<NA>"})
            files_buides_raw = df_v[es_fila_buida]
            total_files_buides = len(files_buides_raw)

            resum_buides_list = []
            recompte_per_full = pd.DataFrame()
            if total_files_buides > 0:
                recompte_per_full = files_buides_raw.groupby("pestanya_origen", sort=False).size().reset_index(name="recompte")
                for row_b in recompte_per_full.itertuples(index=False):
                    n_b = int(row_b.recompte)
                    resum_buides_list.append({
                        "id_venda": "—",
                        "data": "—",
                        "codi_producte": f"({n_b} files buides)",
                        "quantitat": 0,
                        "canal": "—",
                        "pestanya_origen": row_b.pestanya_origen,
                        "motiu_alerta": f"Avís: {n_b} {'fila en blanc o sense codi ignorada' if n_b == 1 else 'files en blanc o sense codi ignorades'} en aquest full"
                    })
            df_resum_buides = pd.DataFrame(resum_buides_list)

            df_v_amb_codi = df_v[~es_fila_buida].copy()

            # 3. Detecció intel·ligent de duplicats vs. coincidències reals
            if te_id_venda:
                mask_dup = df_v_amb_codi.duplicated(subset=["id_venda", "data", "codi_producte", "quantitat", "canal"], keep="first")
                duplicats = df_v_amb_codi[mask_dup].copy()
                duplicats["motiu_alerta"] = "Duplicat real eliminat (mateix nº tiquet/factura i línia)"
                df_v_net = df_v_amb_codi[~mask_dup]
                possibles_coincidencies = pd.DataFrame()
            else:
                mask_coinc = df_v_amb_codi.duplicated(subset=["data", "codi_producte", "quantitat", "canal"], keep="first")
                possibles_coincidencies = df_v_amb_codi[mask_coinc].copy()
                possibles_coincidencies["motiu_alerta"] = "Avís: Possible venda repetida (mateix dia/SKU sense nº tiquet) — Conservada al càlcul"
                duplicats = pd.DataFrame()
                df_v_net = df_v_amb_codi

            # 4. Creuament amb catàleg
            df_merged = pd.merge(df_v_net, df_c, on="codi_producte", how="left")

            mask_no_trobat = df_merged["nom_producte"].isna()
            no_trobats = df_merged.loc[mask_no_trobat, ["id_venda", "data", "codi_producte", "quantitat", "canal", "pestanya_origen"]].copy()
            no_trobats["motiu_alerta"] = "Codi inexistent al catàleg"

            blocs_critics = [df for df in (duplicats, no_trobats) if not df.empty]
            blocs_avisos = [df for df in (possibles_coincidencies, df_resum_buides) if not df.empty]
            tots_blocs_alerta = blocs_critics + blocs_avisos

            df_alertes = pd.concat(tots_blocs_alerta, ignore_index=True) if tots_blocs_alerta else pd.DataFrame()
            errors_critics_count = len(duplicats) + len(no_trobats)
            avisos_count = len(possibles_coincidencies) + total_files_buides

            # 5. Càlculs financers vectoritzats
            df_valid = df_merged[~mask_no_trobat].reset_index(drop=True)
            df_valid["facturacio"] = df_valid["quantitat"] * df_valid["preu_venda"]
            df_valid["cost_total"] = df_valid["quantitat"] * df_valid["cost_unitari"]
            df_valid["marge_net"] = df_valid["facturacio"] - df_valid["cost_total"]

            df_resum = (
                df_valid.groupby("categoria", as_index=False)
                .agg(
                    unitats_venudes=("quantitat", "sum"),
                    facturacio_total=("facturacio", "sum"),
                    cost_total=("cost_total", "sum"),
                    marge_net=("marge_net", "sum")
                )
                .sort_values(by="facturacio_total", ascending=False)
                .reset_index(drop=True)
            )
            df_resum["marges_pct"] = (df_resum["marge_net"] / df_resum["facturacio_total"].replace(0, pd.NA)).fillna(0.0)

            df_prod_resum = (
                df_valid.groupby(["codi_producte", "nom_producte", "categoria"], as_index=False)
                .agg(
                    unitats=("quantitat", "sum"),
                    preu_venda=("preu_venda", "first"),
                    cost_unitari=("cost_unitari", "first"),
                    facturacio_total=("facturacio", "sum"),
                    marge_net_total=("marge_net", "sum")
                )
                .sort_values(by="marge_net_total", ascending=False)
                .reset_index(drop=True)
            )
            df_prod_resum["marge_unitari"] = df_prod_resum["preu_venda"] - df_prod_resum["cost_unitari"]
            df_prod_resum["rendibilitat_pct"] = (df_prod_resum["marge_unitari"] / df_prod_resum["preu_venda"].replace(0, pd.NA)).fillna(0.0)

            # --- EXPORTACIÓ RÀPIDA A EXCEL AMB FÓRMULES DINÀMIQUES I GRÀFICS ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_resum.to_excel(writer, sheet_name="1_Resum_Executiu", index=False)
                df_valid.to_excel(writer, sheet_name="2_Vendes_Netes", index=False)
                df_prod_resum.to_excel(writer, sheet_name="3_Analisi_Productes", index=False)
                if not df_alertes.empty:
                    df_alertes.to_excel(writer, sheet_name="4_Alertes_Errors", index=False)

                wb = writer.book
                fmt_cap = wb.add_format({"bold": True, "bg_color": "#1B365D", "font_color": "white", "border": 1})
                fmt_alerta = wb.add_format({"bold": True, "bg_color": "#9C0006", "font_color": "white", "border": 1})
                fmt_avis_suau = wb.add_format({"italic": True, "bg_color": "#FEF9E7", "font_color": "#7D6608"})
                fmt_tot_txt = wb.add_format({"bold": True, "bg_color": "#EAEDED", "border": 1, "top": 2})
                fmt_tot_int = wb.add_format({"bold": True, "bg_color": "#EAEDED", "border": 1, "top": 2, "num_format": "#,##0"})
                fmt_tot_eur = wb.add_format({"bold": True, "bg_color": "#EAEDED", "border": 1, "top": 2, "num_format": "#,##0.00 €"})
                fmt_tot_pct = wb.add_format({"bold": True, "bg_color": "#EAEDED", "border": 1, "top": 2, "num_format": "0.0%"})
                fmt_eur = wb.add_format({"num_format": "#,##0.00 €"})
                fmt_pct = wb.add_format({"num_format": "0.0%"})
                fmt_int = wb.add_format({"num_format": "#,##0"})

                # PESTANYA 2: 2_Vendes_Netes
                ws_valid = writer.sheets["2_Vendes_Netes"]
                cols_valid = list(df_valid.columns)
                idx_codi_v = cols_valid.index("codi_producte")
                idx_qty = cols_valid.index("quantitat")
                idx_cat_v = cols_valid.index("categoria")
                idx_pvp = cols_valid.index("preu_venda")
                idx_cost_u = cols_valid.index("cost_unitari")
                idx_fact = cols_valid.index("facturacio")
                idx_cost_t = cols_valid.index("cost_total")
                idx_marge = cols_valid.index("marge_net")

                col_l_codi_v = xl_col_to_name(idx_codi_v)
                col_l_qty = xl_col_to_name(idx_qty)
                col_l_cat_v = xl_col_to_name(idx_cat_v)
                col_l_pvp = xl_col_to_name(idx_pvp)
                col_l_cost_u = xl_col_to_name(idx_cost_u)
                col_l_fact = xl_col_to_name(idx_fact)
                col_l_cost_t = xl_col_to_name(idx_cost_t)
                col_l_marge = xl_col_to_name(idx_marge)

                n_valid = len(df_valid)
                ultima_fila_valid = n_valid + 1

                fact_arr = df_valid["facturacio"].to_numpy()
                cost_arr = df_valid["cost_total"].to_numpy()
                marge_arr = df_valid["marge_net"].to_numpy()

                for r in range(n_valid):
                    er = r + 2
                    ws_valid.write_formula(r + 1, idx_fact, f"={col_l_qty}{er}*{col_l_pvp}{er}", fmt_eur, float(fact_arr[r]))
                    ws_valid.write_formula(r + 1, idx_cost_t, f"={col_l_qty}{er}*{col_l_cost_u}{er}", fmt_eur, float(cost_arr[r]))
                    ws_valid.write_formula(r + 1, idx_marge, f"={col_l_fact}{er}-{col_l_cost_t}{er}", fmt_eur, float(marge_arr[r]))

                if n_valid > 0:
                    fila_tot_v = ultima_fila_valid
                    ws_valid.write(fila_tot_v, 0, "TOTAL GENERAL", fmt_tot_txt)
                    for c_i in range(1, len(cols_valid)):
                        ws_valid.write(fila_tot_v, c_i, "", fmt_tot_txt)
                    ws_valid.write_formula(fila_tot_v, idx_qty, f"=SUM({col_l_qty}2:{col_l_qty}{ultima_fila_valid})", fmt_tot_int, float(df_valid["quantitat"].sum()))
                    ws_valid.write_formula(fila_tot_v, idx_fact, f"=SUM({col_l_fact}2:{col_l_fact}{ultima_fila_valid})", fmt_tot_eur, float(fact_arr.sum()))
                    ws_valid.write_formula(fila_tot_v, idx_cost_t, f"=SUM({col_l_cost_t}2:{col_l_cost_t}{ultima_fila_valid})", fmt_tot_eur, float(cost_arr.sum()))
                    ws_valid.write_formula(fila_tot_v, idx_marge, f"=SUM({col_l_marge}2:{col_l_marge}{ultima_fila_valid})", fmt_tot_eur, float(marge_arr.sum()))

                # PESTANYA 1: 1_Resum_Executiu
                ws_resum = writer.sheets["1_Resum_Executiu"]
                n_cats = len(df_resum)
                rang_cat_v = f"'2_Vendes_Netes'!${col_l_cat_v}$2:${col_l_cat_v}${ultima_fila_valid}"
                rang_qty_v = f"'2_Vendes_Netes'!${col_l_qty}$2:${col_l_qty}${ultima_fila_valid}"
                rang_fact_v = f"'2_Vendes_Netes'!${col_l_fact}$2:${col_l_fact}${ultima_fila_valid}"
                rang_cost_v = f"'2_Vendes_Netes'!${col_l_cost_t}$2:${col_l_cost_t}${ultima_fila_valid}"

                for r in range(n_cats):
                    er = r + 2
                    ws_resum.write_formula(r + 1, 1, f"=SUMIF({rang_cat_v}, A{er}, {rang_qty_v})", fmt_int, float(df_resum.loc[r, "unitats_venudes"]))
                    ws_resum.write_formula(r + 1, 2, f"=SUMIF({rang_cat_v}, A{er}, {rang_fact_v})", fmt_eur, float(df_resum.loc[r, "facturacio_total"]))
                    ws_resum.write_formula(r + 1, 3, f"=SUMIF({rang_cat_v}, A{er}, {rang_cost_v})", fmt_eur, float(df_resum.loc[r, "cost_total"]))
                    ws_resum.write_formula(r + 1, 4, f"=C{er}-D{er}", fmt_eur, float(df_resum.loc[r, "marge_net"]))
                    ws_resum.write_formula(r + 1, 5, f"=IF(C{er}=0,0,E{er}/C{er})", fmt_pct, float(df_resum.loc[r, "marges_pct"]))

                if n_cats > 0:
                    fila_tot_r = n_cats + 1
                    er_tot_r = fila_tot_r + 1
                    ws_resum.write(fila_tot_r, 0, "TOTAL GENERAL", fmt_tot_txt)
                    ws_resum.write_formula(fila_tot_r, 1, f"=SUM(B2:B{fila_tot_r})", fmt_tot_int, float(df_resum["unitats_venudes"].sum()))
                    ws_resum.write_formula(fila_tot_r, 2, f"=SUM(C2:C{fila_tot_r})", fmt_tot_eur, float(df_resum["facturacio_total"].sum()))
                    ws_resum.write_formula(fila_tot_r, 3, f"=SUM(D2:D{fila_tot_r})", fmt_tot_eur, float(df_resum["cost_total"].sum()))
                    ws_resum.write_formula(fila_tot_r, 4, f"=C{er_tot_r}-D{er_tot_r}", fmt_tot_eur, float(df_resum["marge_net"].sum()))
                    tot_fact = float(df_resum["facturacio_total"].sum())
                    tot_pct = float(df_resum["marge_net"].sum() / tot_fact) if tot_fact else 0.0
                    ws_resum.write_formula(fila_tot_r, 5, f"=IF(C{er_tot_r}=0,0,E{er_tot_r}/C{er_tot_r})", fmt_tot_pct, tot_pct)

                # PESTANYA 3: 3_Analisi_Productes
                ws_prod = writer.sheets["3_Analisi_Productes"]
                n_prods = len(df_prod_resum)
                rang_codi_v = f"'2_Vendes_Netes'!${col_l_codi_v}$2:${col_l_codi_v}${ultima_fila_valid}"

                for r in range(n_prods):
                    er = r + 2
                    ws_prod.write_formula(r + 1, 3, f"=SUMIF({rang_codi_v}, A{er}, {rang_qty_v})", fmt_int, float(df_prod_resum.loc[r, "unitats"]))
                    ws_prod.write_formula(r + 1, 6, f"=D{er}*E{er}", fmt_eur, float(df_prod_resum.loc[r, "facturacio_total"]))
                    ws_prod.write_formula(r + 1, 7, f"=G{er}-(D{er}*F{er})", fmt_eur, float(df_prod_resum.loc[r, "marge_net_total"]))
                    ws_prod.write_formula(r + 1, 8, f"=E{er}-F{er}", fmt_eur, float(df_prod_resum.loc[r, "marge_unitari"]))
                    ws_prod.write_formula(r + 1, 9, f"=IF(E{er}=0,0,I{er}/E{er})", fmt_pct, float(df_prod_resum.loc[r, "rendibilitat_pct"]))

                if n_prods > 0:
                    fila_tot_p = n_prods + 1
                    ws_prod.write(fila_tot_p, 0, "TOTAL GENERAL", fmt_tot_txt)
                    for c_i in (1, 2, 4, 5, 8, 9):
                        ws_prod.write(fila_tot_p, c_i, "", fmt_tot_txt)
                    ws_prod.write_formula(fila_tot_p, 3, f"=SUM(D2:D{fila_tot_p})", fmt_tot_int, float(df_prod_resum["unitats"].sum()))
                    ws_prod.write_formula(fila_tot_p, 6, f"=SUM(G2:G{fila_tot_p})", fmt_tot_eur, float(df_prod_resum["facturacio_total"].sum()))
                    ws_prod.write_formula(fila_tot_p, 7, f"=SUM(H2:H{fila_tot_p})", fmt_tot_eur, float(df_prod_resum["marge_net_total"].sum()))

                # Formatació ràpida de capçaleres i amplades
                pestanyes_out = [
                    ("1_Resum_Executiu", df_resum, fmt_cap),
                    ("2_Vendes_Netes", df_valid, fmt_cap),
                    ("3_Analisi_Productes", df_prod_resum, fmt_cap)
                ]
                if not df_alertes.empty:
                    pestanyes_out.append(("4_Alertes_Errors", df_alertes, fmt_alerta))

                for nom_p, df_act, estil in pestanyes_out:
                    ws = writer.sheets[nom_p]
                    mostra = df_act.head(200)
                    for idx, col in enumerate(df_act.columns):
                        ws.write(0, idx, str(col).upper(), estil)
                        max_dades = int(mostra[col].fillna("").astype(str).str.len().max()) if not mostra.empty else 0
                        ample = min(max(max_dades, len(str(col))) + 5, 52)
                        ws.set_column(idx, idx, ample)

                # Ressaltem en groc suau els avisos informatius (coincidències sense tiquet + resum files buides)
                if not df_alertes.empty and len(blocs_avisos) > 0:
                    ws_al = writer.sheets["4_Alertes_Errors"]
                    inici_avisos_row = errors_critics_count + 1
                    total_files_avisos = sum(len(b) for b in blocs_avisos)
                    for r_a in range(total_files_avisos):
                        ws_al.set_row(inici_avisos_row + r_a, None, fmt_avis_suau)

                ws_resum.set_column("B:B", 18, fmt_int)
                ws_resum.set_column("C:E", 18, fmt_eur)
                ws_resum.set_column("F:F", 15, fmt_pct)

                ws_valid.set_column(idx_qty, idx_qty, 14, fmt_int)
                ws_valid.set_column(idx_pvp, idx_marge, 16, fmt_eur)

                ws_prod.set_column("D:D", 14, fmt_int)
                ws_prod.set_column("E:I", 18, fmt_eur)
                ws_prod.set_column("J:J", 16, fmt_pct)

                # --- INCRUSTAR ELS 3 GRÀFICS NATIUS D'EXCEL A 1_Resum_Executiu ---
                if generar_grafics and not df_resum.empty:
                    fila_inici_grafics = n_cats + 4

                    chart_pie = wb.add_chart({"type": "doughnut"})
                    chart_pie.add_series({
                        "name": "% Vendes per Categoria",
                        "categories": ["1_Resum_Executiu", 1, 0, n_cats, 0],
                        "values": ["1_Resum_Executiu", 1, 1, n_cats, 1],
                        "data_labels": {"percentage": True, "category": True}
                    })
                    chart_pie.set_title({"name": "% Distribució de Vendes per Categoria (Unitats)"})
                    chart_pie.set_style(10)
                    ws_resum.insert_chart(fila_inici_grafics, 0, chart_pie, {"x_scale": 1.15, "y_scale": 1.15})

                    chart_cat = wb.add_chart({"type": "column", "subtype": "stacked"})
                    chart_cat.add_series({
                        "name": "Cost Total (€)",
                        "categories": ["1_Resum_Executiu", 1, 0, n_cats, 0],
                        "values": ["1_Resum_Executiu", 1, 3, n_cats, 3],
                        "fill": {"color": "#95A5A6"}
                    })
                    chart_cat.add_series({
                        "name": "Marge Net (€)",
                        "categories": ["1_Resum_Executiu", 1, 0, n_cats, 0],
                        "values": ["1_Resum_Executiu", 1, 4, n_cats, 4],
                        "fill": {"color": "#1E8449"}
                    })
                    chart_cat.set_title({"name": "Desglossament de Facturació: Cost vs. Marge Net"})
                    chart_cat.set_y_axis({"num_format": "#,##0 €"})
                    ws_resum.insert_chart(fila_inici_grafics, 4, chart_cat, {"x_scale": 1.25, "y_scale": 1.15})

                    num_prods_chart = min(12, len(df_prod_resum))
                    chart_prod = wb.add_chart({"type": "column", "subtype": "stacked"})
                    chart_prod.add_series({
                        "name": "Cost Unitari (€)",
                        "categories": ["3_Analisi_Productes", 1, 1, num_prods_chart, 1],
                        "values": ["3_Analisi_Productes", 1, 5, num_prods_chart, 5],
                        "fill": {"color": "#AEB6BF"}
                    })
                    chart_prod.add_series({
                        "name": "Marge Unitari (€)",
                        "categories": ["3_Analisi_Productes", 1, 1, num_prods_chart, 1],
                        "values": ["3_Analisi_Productes", 1, 8, num_prods_chart, 8],
                        "fill": {"color": "#1B4F72"},
                        "data_labels": {"value": True, "num_format": "#,##0.00 €"}
                    })
                    chart_prod.set_title({"name": "Estructura de Preu i Marge Real per Producte (PVP = Cost + Marge)"})
                    chart_prod.set_y_axis({"num_format": "#,##0.00 €"})
                    ws_resum.insert_chart(fila_inici_grafics + 18, 0, chart_prod, {"x_scale": 2.42, "y_scale": 1.25})

            # KPIs superiors a la web
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Facturació Total", f"{df_valid['facturacio'].sum():,.2f} €")
            k2.metric("Marge Net", f"{df_valid['marge_net'].sum():,.2f} €")
            k3.metric("Registres Vàlids", len(df_valid))
            k4.metric(
                "Errors Crítics (Duplicats / SKU)",
                errors_critics_count,
                delta=f"+ {avisos_count} avisos preventius" if avisos_count > 0 else None,
                delta_color="off"
            )

            if te_id_venda:
                st.success(f"🛡️ **Control d'integritat per Tiquet/Factura actiu (`{col_v_tiquet}`):** S'han eliminat **{len(duplicats)} duplicats reals** amb el mateix número de document i s'han conservat totes les compres coincidents amb número de tiquet diferent.")
            elif len(possibles_coincidencies) > 0:
                st.warning(f"⚠️ **Mode preventiu (sense columna de Tiquet/Factura):** S'han detectat **{len(possibles_coincidencies)} vendes coincidents** (mateix dia, SKU i quantitat). S'han conservat al càlcul per no perdre facturació real i s'han marcat a `4_Alertes_Errors` per revisar.")

            if total_files_buides > 0:
                detall_text = " · ".join([f"**{r.pestanya_origen}**: {int(r.recompte)}" for r in recompte_per_full.itertuples(index=False)])
                st.info(f"🧹 **Control de files en blanc:** S'han detectat i ignorat **{total_files_buides} files buides** sense codi ({detall_text}).")

            st.dataframe(df_resum, use_container_width=True)

            # Panell Visual Interactiu a la web
            if generar_grafics and not df_valid.empty:
                st.divider()
                st.subheader("📈 Anàlisi Visual Executiva")
                
                g1, g2 = st.columns(2)
                with g1:
                    fig_pie = px.pie(
                        df_resum,
                        values="unitats_venudes",
                        names="categoria",
                        hole=0.42,
                        title="🍰 % Distribució de Vendes per Categoria / Estil (Unitats)",
                        color_discrete_sequence=px.colors.qualitative.Set2
                    )
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    st.plotly_chart(fig_pie, use_container_width=True)

                with g2:
                    fig_cat_marge = go.Figure(data=[
                        go.Bar(name="Cost Total (€)", x=df_resum["categoria"], y=df_resum["cost_total"], marker_color="#95A5A6"),
                        go.Bar(name="Marge Net (€)", x=df_resum["categoria"], y=df_resum["marge_net"], marker_color="#1E8449")
                    ])
                    fig_cat_marge.update_layout(
                        barmode="stack",
                        title="💰 Desglossament de Facturació: Cost vs. Marge Net",
                        yaxis_title="Euros (€)"
                    )
                    st.plotly_chart(fig_cat_marge, use_container_width=True)

                top_prods = df_prod_resum.head(12)
                fig_prod = go.Figure()
                fig_prod.add_trace(go.Bar(
                    name="Cost Unitari (€)",
                    x=top_prods["nom_producte"],
                    y=top_prods["cost_unitari"],
                    marker_color="#AEB6BF"
                ))
                fig_prod.add_trace(go.Bar(
                    name="Marge Unitari (€)",
                    x=top_prods["nom_producte"],
                    y=top_prods["marge_unitari"],
                    marker_color="#1B4F72",
                    text=(top_prods["rendibilitat_pct"] * 100).round(1).astype(str) + "% marge",
                    textposition="outside"
                ))
                fig_prod.update_layout(
                    barmode="stack",
                    title="🏷️ Estructura de Preu i Marge Real per Producte (PVP = Cost + Marge)",
                    yaxis_title="Preu Unitari (€)",
                    xaxis_tickangle=-25,
                    height=460
                )
                st.plotly_chart(fig_prod, use_container_width=True)

            sufix_final = netejar_nom_arxiu(etiqueta_personalitzada)
            nom_fitxer_sortida = f"informe_depurat_{sufix_final}.xlsx"

            st.download_button(
                label=f"📥 Descarregar {nom_fitxer_sortida}",
                data=buffer.getvalue(),
                file_name=nom_fitxer_sortida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )