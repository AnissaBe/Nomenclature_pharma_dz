import streamlit as st

st.set_page_config(
    page_title="Nomenclature Pharmaceutique DZ",
    page_icon="",
    layout="wide"
)

password = st.text_input("Mot de passe", type="password")
if password != "tonmotdepasse":
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
        .instructions * {
            font-size: 13px !important;
            font-family: 'Manrope', sans-serif !important;
            color: #111827 !important;
            line-height: 2.2 !important;
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
            <div class="card-number">&#201;tape 01</div>
            <div class="card-label">Normalisation des formes gal&#233;niques</div>
            <div class="card-desc">Standardisation automatique vers les formes gal&#233;niques de base</div>
        </div>
        <div class="card">
            <div class="card-accent"></div>
            <div class="card-number">&#201;tape 02</div>
            <div class="card-label">Harmonisation des dosages</div>
            <div class="card-desc">Suppression des espaces et standardisation des unit&#233;s pharmaceutiques</div>
        </div>
        <div class="card">
            <div class="card-accent"></div>
            <div class="card-number">&#201;tape 03</div>
            <div class="card-label">Analyse des enregistrements</div>
            <div class="card-desc">Nombre d&#39;enregistrements par DCI, dosage, laboratoire et forme gal&#233;nique</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# ── Instructions ─────────────────────────────────────────────────
with st.expander("📋 Comment préparer votre fichier ?"):
    st.markdown("""
    1. Ouvrir le fichier Excel relatif à la nomenclature (xlsx)
    2. Supprimer les 2 onglets **Non renouvelés** et **Retraits**
    3. Supprimer le logo, le titre, puis l'ensemble des lignes vides au-dessus du tableau
    4. Supprimer les filtres du tableau (Données → Effacer les filtres)
    5. Fichier → Enregistrer sous → CSV UTF-8 (séparateur virgule ou point-virgule)
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

# ── Upload natif ─────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "",
    type=["csv"],
    label_visibility="collapsed"
)

if uploaded_file is None:
    st.stop()

    # ── Traitement ───────────────────────────────────────────────────
from traitement import traiter_fichier
import io
 
with st.spinner("Traitement en cours..."):
    df, resume = traiter_fichier(uploaded_file)
 
st.success("Traitement terminé !")
 
# ── Métriques ────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Spécialités traitées", f"{len(df):,}")
with col2:
    st.metric("Formes galéniques", df["FORME_STD"].nunique())
with col3:
    st.metric("Combinaisons uniques", f"{len(resume):,}")
 
# ── Aperçu ───────────────────────────────────────────────────────
st.markdown("### Aperçu du fichier nettoyé")
st.dataframe(
    df[["DENOMINATION COMMUNE INTERNATIONALE", "DOSAGE", "FORME", "FORME_STD"]].head(50),
    use_container_width=True
)
 
# ── Téléchargements ──────────────────────────────────────────────
col1, col2 = st.columns(2)
 
with col1:
    csv_complet = df.to_csv(sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="Télécharger le fichier complet nettoyé",
        data=csv_complet,
        file_name="Nomenclature_clean.csv",
        mime="text/csv"
    )
 
with col2:
    csv_resume = resume.to_csv(sep=";", encoding="utf-8-sig", index=False).encode("utf-8-sig")
    st.download_button(
        label="Télécharger le tableau des enregistrements par DCI",
        data=csv_resume,
        file_name="Nomenclature_resume_DCI.csv",
        mime="text/csv"
    )
