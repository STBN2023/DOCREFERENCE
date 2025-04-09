import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO

def remplacer_placeholder(slide, placeholder, nouvelle_valeur):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if placeholder in paragraph.text:
                    texte_complet = paragraph.text.replace(placeholder, nouvelle_valeur)

                    styles = []
                    for run in paragraph.runs:
                        styles.append({
                            'font_name': run.font.name,
                            'font_size': run.font.size,
                            'bold': run.font.bold,
                            'italic': run.font.italic,
                            'underline': run.font.underline,
                            'color': run.font.color.rgb if run.font.color.type else None,
                        })

                    for _ in range(len(paragraph.runs)):
                        paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)

                    run = paragraph.add_run()
                    run.text = texte_complet
                    
                    if styles:
                        run.font.name = styles[0]['font_name']
                        run.font.size = styles[0]['font_size']
                        run.font.bold = styles[0]['bold']
                        run.font.italic = styles[0]['italic']
                        run.font.underline = styles[0]['underline']
                        if styles[0]['color']:
                            run.font.color.rgb = styles[0]['color']

st.title("🚀 Générateur PPTX Final (Style conservé)")

fichier_excel = st.file_uploader("📥 Charge ton fichier Excel", type=["xlsx"])
fichier_pptx = st.file_uploader("📥 Charge ton modèle PowerPoint", type=["pptx"])

if fichier_excel and fichier_pptx:
    donnees = pd.read_excel(fichier_excel)
    prs = Presentation(fichier_pptx)

    placeholders = {
        '{{Titre_Residence}}': 'Nom de la residence',
        '{{Ville}}': 'Ville',
        '{{Details}}': 'details',
        '{{Maitre_Ouvrage}}': 'Maitre d\'ouvrage',
        '{{Assistant}}': 'Assistant',
        '{{Montant_Travaux}}': 'montant des travaux',
        '{{Type_Travaux}}': 'type de travaux',
        '{{Dates_Realisation}}': 'realisation',
        '{{Surface}}': 'Surface',
        '{{Travaux_Exterieurs}}': 'travaux exterieurs',
        '{{Travaux_Interieurs}}': 'travaux interieurs'
    }

    for slide, (_, projet) in zip(prs.slides, donnees.iterrows()):
        for placeholder, colonne in placeholders.items():
            valeur = projet[colonne]
            if pd.notna(valeur):
                remplacer_placeholder(slide, placeholder, str(valeur))

    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)

    st.download_button(
        label="📤 Télécharger PPTX parfaitement stylisé",
        data=pptx_io,
        file_name="Presentation_Finalisee.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
