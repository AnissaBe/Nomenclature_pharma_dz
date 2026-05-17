import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Nomenclature Pharmaceutique DZ",
    page_icon="",
    layout="wide"
)

# ── Mot de passe ─────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    password = st.text_input("Mot de passe", type="password")
    if password == "APA_appnomenc_26":
        st.session_state.authenticated = True
        st.rerun()
    else:
        st.stop()

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700;800&display=swap');
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        .main { background-color: #F8F9FA; }
        .hero {
            background: linear-gradient(135deg, #1A1A2E 0%, #16213E 60%, #0F3460 100%);
            padding: 44px 48px;
            border-radius: 16px;
            margin-bottom: 20px;
        }
        .hero-tag {
            display: inline-block;
            background: rgba(0,200,150,0.15);
            border: 0.5px solid rgba(0,200,150,0.3);
            border-radius: 20px;
            padding: 5px 14px;
            font-size: 11px;
            font-weight: 600;
            color: #00C896;
            letter-spacing: 1px;
            text-transform: uppercase;
            margin-bottom: 18px;
        }
        .hero h1 {
            color: white;
            font-size: 30px;
            font-weight: 800;
            line-height: 1.4;
            margin-bottom: 16px;
            letter-spacing: -0.5px;
        }
        .hero h1 span { color: #00C896; }
        .hero-rep {
            color: white;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 4px;
        }
        .hero-min {
            color: rgba(255,255,255,0.6);
            font-size: 12px;
            font-weight: 400;
        }
        .cards {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
        }
        .card {
            flex: 1;
            background: white;
            border-radius: 12px;
            padding: 20px 22px;
            border: 0.5px solid #E5E7EB;
        }
        .card-accent {
            width: 32px; height: 3px;
            background: #00C896;
            border-radius: 2px;
            margin-bottom: 14px;
        }
        .card-number {
            font-size: 10px;
            font-weight: 700;
            color: #00C896;
            letter-spacing: 1.5px;
            margin-bottom: 6px;
            text-transform: uppercase;
        }
        .card-label {
            font-size: 13px;
            color: #111827;
            font-weight: 700;
            line-height: 1.4;
            margin-bottom: 4px;
        }
        .card-desc {
            font-size: 12px;
            color: #6B7280;
            line-height: 1.5;
        }
    </style>
""", unsafe_allow_html=True)

# ── Hero ─────────────────────────────────────────────────────────
st.markdown("""
    <div class="hero">
        <div class="hero-tag">Outil d&#39;analyse</div>
        <h1>Nomenclature nationale des produits<br>
        pharmaceutiques &#224; usage de la <span>m&#233;decine humaine</span></h1>
        <div class="hero-rep">R&#233;publique Alg&#233;rienne D&#233;mocratique et Populaire</div>
        <div class="hero-min">Minist&#232;re de l&#39;Industrie Pharmaceutique</div>
    </div>
""", unsafe_allow_html=True)

# ── Cards ────────────────────────────────────────────────────────
st.markdown("""
    <div class="cards">
        <div class="card">
            <div class="card-accent"></div>
            <div class="card-number">Onglet 01</div>
            <div class="card-label">Positionnement par laboratoire</div>
            <div class="card-desc">Analyse d&#233;taill&#233;e par dosage et par laboratoire &#8212; une ligne par laboratoire avec le nombre d&#39;enregistrements en fabrication locale et en importation</div>
        </div>
        <div class="card">
            <div class="card-accent"></div>
            <div class="card-number">Onglet 02</div>
            <div class="card-label">Positionnement par code</div>
            <div class="card-desc">Vue synth&#233;tique par code produit &#8212; une ligne par code avec l&#39;ensemble des laboratoires regroup&#233;s, le nombre d&#39;enregistrements en fabrication locale, en importation et le total</div>
        </div>
        <div class="card">
            <div class="card-accent"></div>
            <div class="card-number">Onglet 03</div>
            <div class="card-label">Anomalies codes</div>
            <div class="card-desc">D&#233;tection automatique des codes associ&#233;s &#224; des DCI dont la d&#233;nomination varie &#8212; orthographe diff&#233;rente ou erreur de saisie &#224; v&#233;rifier</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# ── Instructions ─────────────────────────────────────────────────
with st.expander("Comment pr\u00e9parer votre fichier ?"):
    st.markdown("""
1. Ouvrir le fichier Excel relatif à la nomenclature (xlsx)
2. Copier l'ensemble des lignes de l'onglet **Non renouvelés** en dessous de celles de la nomenclature mise à jour
3. Supprimer les 2 onglets **Non renouvelés** et **Retraits**
4. Supprimer le logo, le titre, puis l'ensemble des lignes vides au-dessus du tableau
5. Supprimer les colonnes vides situées à droite du tableau — leur présence génère des colonnes sans nom lors de l'import
6. Supprimer les filtres du tableau (Données → Effacer les filtres)
7. Fichier → Enregistrer sous → CSV UTF-8 (séparateur virgule ou point-virgule)
    """)

# ── Titre upload ─────────────────────────────────────────────────
st.markdown("""
    <div style="
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: white;
        border-radius: 14px 14px 0 0;
        padding: 24px 32px 16px 32px;
        border: 0.5px solid #E5E7EB;
        border-bottom: none;
        font-family: 'Manrope', sans-serif;
        margin-bottom: 0;
    ">
        <div style="font-size:16px; font-weight:800; color:#111827;">
            Importer votre fichier
        </div>
        <div style="background:#ECFDF5; color:#059669; font-size:11px;
        font-weight:700; padding:4px 12px; border-radius:20px;">
            CSV &#8212; UTF-8 &#8212; Virgule ou Point-virgule
        </div>
    </div>
""", unsafe_allow_html=True)

# ── Upload ───────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "",
    type=["csv"],
    label_visibility="collapsed"
)

if uploaded_file is None:
    st.stop()

# ── Traitement ───────────────────────────────────────────────────
from traitement_positionnement import traiter_positionnement, formater_feuille

with st.spinner("Traitement en cours..."):
    par_labo, par_code, anomalies = traiter_positionnement(uploaded_file)

st.success("Traitement termin\u00e9 !")

# ── Métriques ────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.metric("Codes produits", f"{par_code['CODE'].nunique():,}")
with col2:
    st.metric("Anomalies d\u00e9tect\u00e9es", f"{len(anomalies):,}")

# ── Tabs ─────────────────────────────────────────────────────────
tab1, tab2 = st.tabs([
    "Positionnement complet",
    "Analyse par DCI"
])

# ════════════════════════════════════════════════════════════════
# TAB 1 — Positionnement complet + téléchargement Excel
# ════════════════════════════════════════════════════════════════
with tab1:
    # Téléchargement Excel 3 onglets
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        par_labo.to_excel(writer, sheet_name="Par laboratoire", index=False)
        par_code.to_excel(writer, sheet_name="Par code", index=False)
        if len(anomalies) > 0:
            anomalies.to_excel(writer, sheet_name="Anomalies_code", index=False)
        # Formatage
        formater_feuille(writer.sheets["Par laboratoire"], par_labo)
        formater_feuille(writer.sheets["Par code"], par_code)
        if len(anomalies) > 0:
            formater_feuille(writer.sheets["Anomalies_code"], anomalies, col_dci="DCI_1")
    output.seek(0)

    st.download_button(
        label="T\u00e9l\u00e9charger le fichier Positionnement Concurrentiel (Excel)",
        data=output,
        file_name="Positionnement_Concurrentiel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ════════════════════════════════════════════════════════════════
# TAB 2 — Analyse par DCI
# ════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Rechercher et s\u00e9lectionner des DCI")

    liste_dci = sorted(par_code["DCI"].dropna().unique().tolist())

    dci_selectionnees = st.multiselect(
        "Tapez pour rechercher et s\u00e9lectionner",
        options=liste_dci,
        placeholder="Ex: DICLOFENAC, AMOXICILLINE..."
    )

    if not dci_selectionnees:
        st.info("Recherchez et s\u00e9lectionnez au moins une DCI pour g\u00e9n\u00e9rer le tableau.")
        st.stop()

    # Codes correspondant aux DCI sélectionnées
    codes_selectionnes = par_code[par_code["DCI"].isin(dci_selectionnees)]["CODE"].tolist()

    # Filtrer par_labo sur ces codes
    par_labo_filtre = par_labo[
        par_labo["CODE"].isin(codes_selectionnes) |
        par_labo["CODE"].isin([f"Total {c}" for c in codes_selectionnes])
    ]

    st.markdown(f"### R\u00e9sultats pour : {', '.join(dci_selectionnees)}")
    st.dataframe(par_labo_filtre, use_container_width=True)

    # T\u00e9l\u00e9chargement Excel filtré
    output_filtre = io.BytesIO()
    with pd.ExcelWriter(output_filtre, engine="openpyxl") as writer:
        par_labo_filtre.to_excel(writer, sheet_name="Analyse DCI", index=False)
        par_code[par_code["DCI"].isin(dci_selectionnees)].to_excel(
            writer, sheet_name="Par code", index=False
        )
        # Formatage
        formater_feuille(writer.sheets["Analyse DCI"], par_labo_filtre)
        formater_feuille(writer.sheets["Par code"], par_code[par_code["DCI"].isin(dci_selectionnees)])
    output_filtre.seek(0)

    st.download_button(
        label="T\u00e9l\u00e9charger l\u2019analyse filtr\u00e9e (Excel)",
        data=output_filtre,
        file_name="Analyse_DCI_filtre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )