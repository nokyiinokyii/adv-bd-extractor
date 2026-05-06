import streamlit as st
import anthropic
import base64
import json
import io
from pathlib import Path

# ─── Config page ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ADV — Extraction BC",
    page_icon="📋",
    layout="wide",
)

# ─── CSS minimal ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { padding-top: 1.5rem; }
    .stAlert { border-radius: 8px; }
    div[data-testid="metric-container"] {
        background: #f8f8f6;
        border: 0.5px solid #e0dfd8;
        border-radius: 8px;
        padding: 12px 16px;
    }
    .ligne-card {
        background: white;
        border: 0.5px solid #e0dfd8;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .badge-ok {
        background: #e1f5ee; color: #0f6e56;
        padding: 2px 10px; border-radius: 6px;
        font-size: 12px; font-weight: 500;
    }
    .badge-warn {
        background: #faeeda; color: #854f0b;
        padding: 2px 10px; border-radius: 6px;
        font-size: 12px; font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)


# ─── Sidebar : clé API ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    api_key = st.text_input(
        "Clé API Anthropic",
        type="password",
        placeholder="sk-ant-...",
        help="Récupérez votre clé sur console.anthropic.com"
    )
    st.markdown("---")
    st.markdown("### À propos")
    st.markdown(
        "MVP ADV — Extraction automatique de bons de commande.\n\n"
        "**Stack :** Streamlit · Claude API · openpyxl\n\n"
        "**Coût estimé :** ~0,01 € par BC analysé"
    )


# ─── Prompt d'extraction ───────────────────────────────────────────────────────
SYSTEM_PROMPT = """Tu es un assistant spécialisé dans l'analyse de bons de commande pour un ADV (Administration Des Ventes).
Extrait toutes les informations du document et réponds UNIQUEMENT en JSON valide, sans preamble ni backticks Markdown.

Structure exacte attendue :
{
  "client": "Nom complet du client",
  "numero_commande": "Numéro du BC",
  "date_commande": "JJ/MM/AAAA",
  "date_livraison_souhaitee": "JJ/MM/AAAA ou null",
  "adresse_livraison": "Adresse complète ou null",
  "contact_client": "Nom du contact ou null",
  "email_contact": "Email ou null",
  "incoterms": "EXW / DAP / DDP / etc. ou null",
  "devise": "EUR",
  "lignes": [
    {
      "numero_ligne": 1,
      "reference_client": "Référence produit telle qu'écrite dans le BC",
      "designation": "Description du produit",
      "quantite": 10,
      "unite": "pcs / kg / m / etc.",
      "prix_unitaire": 25.50,
      "total_ligne": 255.00,
      "delai_specifique": "Date ou délai particulier pour cette ligne, ou null"
    }
  ],
  "total_ht": 255.00,
  "tva_pct": 20,
  "total_ttc": 306.00,
  "conditions_paiement": "30 jours fin de mois ou null",
  "remarques": "Conditions particulières, urgences, incohérences détectées, ou null",
  "confiance": "haute / moyenne / basse"
}

Si une information est absente du document, utilise null.
Pour les montants, utilise des nombres décimaux (pas de chaînes de caractères).
Le champ 'confiance' indique ta certitude sur l'extraction globale."""


# ─── Fonction d'extraction Claude ─────────────────────────────────────────────
def extraire_bc(fichier_bytes: bytes, nom_fichier: str, cle_api: str) -> dict:
    """Envoie le fichier à Claude et retourne le JSON extrait."""
    client = anthropic.Anthropic(api_key=cle_api)

    # Détecter le type MIME
    ext = Path(nom_fichier).suffix.lower()
    if ext == ".pdf":
        media_type = "application/pdf"
        bloc_fichier = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": base64.standard_b64encode(fichier_bytes).decode("utf-8"),
            },
        }
    elif ext in [".png", ".jpg", ".jpeg", ".webp"]:
        media_type = f"image/{'jpeg' if ext in ['.jpg', '.jpeg'] else ext[1:]}"
        bloc_fichier = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": base64.standard_b64encode(fichier_bytes).decode("utf-8"),
            },
        }
    else:
        raise ValueError(f"Format non supporté : {ext}")

    reponse = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": [
                    bloc_fichier,
                    {"type": "text", "text": "Analyse ce bon de commande et extrais toutes les informations."},
                ],
            }
        ],
    )

    texte = reponse.content[0].text.strip()
    # Nettoyage au cas où le modèle ajoute des backticks
    texte = texte.replace("```json", "").replace("```", "").strip()
    return json.loads(texte)


# ─── Génération Excel ──────────────────────────────────────────────────────────
def generer_excel(donnees: dict) -> bytes:
    """Génère un fichier Excel structuré prêt pour SAP."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        st.error("openpyxl non installé. Lancez : pip install openpyxl")
        return b""

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Commande SAP"

    # Styles
    gris = PatternFill("solid", fgColor="F1EFE8")
    bleu = PatternFill("solid", fgColor="E6F1FB")
    font_titre = Font(bold=True, size=12)
    font_header = Font(bold=True, size=10)
    align_centre = Alignment(horizontal="center", vertical="center")
    bordure = Border(
        bottom=Side(style="thin", color="D3D1C7"),
        top=Side(style="thin", color="D3D1C7"),
    )

    # En-tête
    ws.merge_cells("A1:H1")
    ws["A1"] = "BON DE COMMANDE — SAISIE SAP"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].fill = bleu
    ws["A1"].alignment = align_centre

    # Infos client
    infos = [
        ("Client", donnees.get("client", "")),
        ("N° commande", donnees.get("numero_commande", "")),
        ("Date commande", donnees.get("date_commande", "")),
        ("Date livraison", donnees.get("date_livraison_souhaitee", "")),
        ("Incoterms", donnees.get("incoterms", "")),
        ("Conditions paiement", donnees.get("conditions_paiement", "")),
        ("Devise", donnees.get("devise", "EUR")),
        ("Adresse livraison", donnees.get("adresse_livraison", "")),
    ]

    row = 3
    for label, valeur in infos:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True, size=10)
        ws.cell(row=row, column=1).fill = gris
        ws.cell(row=row, column=2, value=str(valeur) if valeur else "").font = Font(size=10)
        ws.merge_cells(f"B{row}:H{row}")
        row += 1

    # Séparateur
    row += 1

    # En-têtes lignes
    headers = ["#", "Réf. client", "Désignation", "Qté", "Unité", "Prix U. (€)", "Total HT (€)", "Délai"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = font_header
        cell.fill = bleu
        cell.alignment = align_centre
        cell.border = bordure

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 18

    # Lignes produits
    row += 1
    for ligne in donnees.get("lignes", []):
        ws.cell(row=row, column=1, value=ligne.get("numero_ligne", ""))
        ws.cell(row=row, column=2, value=ligne.get("reference_client", ""))
        ws.cell(row=row, column=3, value=ligne.get("designation", ""))
        ws.cell(row=row, column=4, value=ligne.get("quantite", ""))
        ws.cell(row=row, column=5, value=ligne.get("unite", ""))
        ws.cell(row=row, column=6, value=ligne.get("prix_unitaire", ""))
        ws.cell(row=row, column=7, value=ligne.get("total_ligne", ""))
        ws.cell(row=row, column=8, value=ligne.get("delai_specifique", "") or "")
        if row % 2 == 0:
            for col in range(1, 9):
                ws.cell(row=row, column=col).fill = gris
        row += 1

    # Totaux
    row += 1
    ws.cell(row=row, column=6, value="Total HT").font = Font(bold=True)
    ws.cell(row=row, column=7, value=donnees.get("total_ht", "")).font = Font(bold=True)
    row += 1
    ws.cell(row=row, column=6, value=f"TVA {donnees.get('tva_pct', 20)}%").font = Font(bold=True)
    tva = (donnees.get("total_ht") or 0) * (donnees.get("tva_pct", 20) / 100)
    ws.cell(row=row, column=7, value=round(tva, 2)).font = Font(bold=True)
    row += 1
    ws.cell(row=row, column=6, value="Total TTC").font = Font(bold=True)
    ws.cell(row=row, column=7, value=donnees.get("total_ttc", "")).font = Font(bold=True)
    ws.cell(row=row, column=7).fill = bleu

    # Remarques
    if donnees.get("remarques"):
        row += 2
        ws.cell(row=row, column=1, value="⚠ Points d'attention").font = Font(bold=True, color="854F0B")
        row += 1
        ws.cell(row=row, column=1, value=donnees["remarques"])
        ws.merge_cells(f"A{row}:H{row}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── Affichage des résultats ───────────────────────────────────────────────────
def afficher_resultats(d: dict):
    devise = d.get("devise", "EUR")
    symbole = "€" if devise == "EUR" else devise

    def fmt(val):
        if val is None:
            return "—"
        try:
            return f"{float(val):,.2f} {symbole}".replace(",", " ")
        except Exception:
            return str(val)

    # Métriques rapides
    lignes = d.get("lignes", [])
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Client", d.get("client", "—"))
    col2.metric("N° commande", d.get("numero_commande", "—"))
    col3.metric("Lignes", len(lignes))
    col4.metric("Total HT", fmt(d.get("total_ht")))

    st.markdown("---")

    # Deux colonnes : infos + lignes
    left, right = st.columns([1, 2])

    with left:
        st.markdown("##### Informations générales")
        champs = {
            "Date commande": d.get("date_commande"),
            "Date livraison": d.get("date_livraison_souhaitee"),
            "Contact": d.get("contact_client"),
            "Email": d.get("email_contact"),
            "Incoterms": d.get("incoterms"),
            "Paiement": d.get("conditions_paiement"),
            "Adresse livraison": d.get("adresse_livraison"),
        }
        for label, val in champs.items():
            if val:
                st.markdown(f"**{label}** : {val}")

        confiance = d.get("confiance", "moyenne")
        couleur = {"haute": "🟢", "moyenne": "🟡", "basse": "🔴"}.get(confiance, "🟡")
        st.markdown(f"\n**Confiance extraction** : {couleur} {confiance.capitalize()}")

        if d.get("remarques"):
            st.warning(f"**Points d'attention**\n\n{d['remarques']}")

    with right:
        st.markdown("##### Lignes de commande")
        for ligne in lignes:
            with st.container():
                c1, c2, c3, c4 = st.columns([3, 1, 1, 2])
                c1.markdown(f"**{ligne.get('reference_client', '')}**  \n{ligne.get('designation', '')}")
                c2.markdown(f"**{ligne.get('quantite', '')}** {ligne.get('unite', '')}")
                c3.markdown(fmt(ligne.get("prix_unitaire")))
                c4.markdown(f"**{fmt(ligne.get('total_ligne'))}**")
                if ligne.get("delai_specifique"):
                    st.caption(f"⏱ Délai spécifique : {ligne['delai_specifique']}")
                st.markdown("<hr style='margin:6px 0; border-color:#e0dfd8'>", unsafe_allow_html=True)

        # Totaux
        tc1, tc2, tc3 = st.columns([3, 1, 2])
        tc1.markdown(f"**Total HT**")
        tc3.markdown(f"**{fmt(d.get('total_ht'))}**")
        if d.get("tva_pct"):
            tc1b, tc2b, tc3b = st.columns([3, 1, 2])
            tc1b.markdown(f"TVA {d.get('tva_pct', 20)} %")
            tva = (d.get("total_ht") or 0) * (d.get("tva_pct", 20) / 100)
            tc3b.markdown(fmt(tva))
        if d.get("total_ttc"):
            tc1c, tc2c, tc3c = st.columns([3, 1, 2])
            tc1c.markdown(f"**Total TTC**")
            tc3c.markdown(f"**{fmt(d.get('total_ttc'))}**")


# ─── Interface principale ──────────────────────────────────────────────────────
st.title("📋 Extraction de bons de commande")
st.caption("Uploadez un BC (PDF ou image) — l'IA extrait automatiquement toutes les données utiles.")

if not api_key:
    st.info("👈 Entrez votre clé API Anthropic dans la barre latérale pour commencer.")
    st.stop()

fichier = st.file_uploader(
    "Déposez votre bon de commande",
    type=["pdf", "png", "jpg", "jpeg"],
    help="PDF ou image (PNG, JPG). Taille max recommandée : 10 Mo.",
)

if fichier:
    col_btn, col_info = st.columns([1, 3])
    with col_btn:
        lancer = st.button("🔍 Analyser le BC", type="primary", use_container_width=True)
    with col_info:
        st.caption(f"Fichier : **{fichier.name}** · {fichier.size / 1024:.0f} Ko")

    if lancer:
        with st.spinner("Analyse en cours — Claude lit votre bon de commande..."):
            try:
                donnees = extraire_bc(fichier.read(), fichier.name, api_key)
                st.session_state["donnees_bc"] = donnees
                st.session_state["nom_fichier"] = fichier.name
                st.success("Extraction réussie !")
            except json.JSONDecodeError as e:
                st.error(f"Erreur de parsing JSON : {e}\nVérifiez que le fichier est lisible.")
            except anthropic.AuthenticationError:
                st.error("Clé API invalide. Vérifiez votre clé Anthropic.")
            except Exception as e:
                st.error(f"Erreur : {e}")

# Affichage des résultats s'ils existent en session
if "donnees_bc" in st.session_state:
    donnees = st.session_state["donnees_bc"]

    st.markdown("---")
    afficher_resultats(donnees)

    st.markdown("---")

    # Actions
    col_ex, col_json, col_reset = st.columns([2, 2, 1])

    with col_ex:
        excel_bytes = generer_excel(donnees)
        if excel_bytes:
            nom_base = Path(st.session_state.get("nom_fichier", "bc")).stem
            st.download_button(
                label="📥 Télécharger Excel SAP",
                data=excel_bytes,
                file_name=f"{nom_base}_SAP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

    with col_json:
        with st.expander("Voir le JSON brut"):
            st.json(donnees)

    with col_reset:
        if st.button("🔄 Nouveau BC", use_container_width=True):
            del st.session_state["donnees_bc"]
            st.rerun()
