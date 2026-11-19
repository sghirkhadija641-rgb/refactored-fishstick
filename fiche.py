"""Génère la fiche A4 des formulaires Ré-Use : un PDF cliquable (et une image PNG) avec les deux accès,
« Saisie d'une réclamation » et « Traitement de réclamation », chacun avec son lien et son bouton cliquables."""
import pymupdf
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white

ACCES = [
    dict(titre="Saisie d'une réclamation", url="https://tinyurl.com/reuse-reclamations",
         texte=["Enregistrer une nouvelle réclamation.", "Un code de suivi s'affiche à l'envoi."],
         bouton="Ouvrir la saisie"),
    dict(titre="Traitement de réclamation", url="https://tinyurl.com/reuse-traitement",
         texte=["Après l'intervention, avec le", "code de suivi de la réclamation."],
         bouton="Ouvrir le traitement"),
]
SORTIE = 'fiche-reclamations-reuse'

pdfmetrics.registerFont(TTFont('B', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('R', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
teal = HexColor('#0f6e86'); ink = HexColor('#14262c'); muted = HexColor('#5b7078'); line = HexColor('#d6e0e3'); soft = HexColor('#f3f6f7')

W, H = A4
c = canvas.Canvas(SORTIE + '.pdf', pagesize=A4); c.setTitle('Réclamations Ré-Use')
c.setFillColor(teal); c.rect(0, H-130, W, 130, stroke=0, fill=1)
c.setFillColor(white); c.setFont('B', 34); c.drawCentredString(W/2, H-70, 'Réclamations Ré-Use')
c.setFont('R', 14); c.setFillColor(HexColor('#dff1f5')); c.drawCentredString(W/2, H-100, 'Choisissez le formulaire : cliquez sur le bouton ou le lien')

boxw = W - 80; boxh = 175; gap = 26; top = H - 170
for i, a in enumerate(ACCES):
    x0 = 40; y0 = top - i*(boxh + gap)
    c.setFillColor(soft); c.setStrokeColor(line); c.roundRect(x0, y0-boxh, boxw, boxh, 12, stroke=1, fill=1)
    c.setFillColor(teal); c.roundRect(x0, y0-boxh, 7, boxh, 3, stroke=0, fill=1)
    c.setFillColor(ink); c.setFont('B', 20); c.drawString(x0+30, y0-42, a['titre'])
    c.setFillColor(muted); c.setFont('R', 12); c.drawString(x0+30, y0-64, ' '.join(a['texte']))
    show = a['url'].replace('https://', '')
    c.setFillColor(teal); c.setFont('B', 13); tw = c.stringWidth(show, 'B', 13); ty = y0 - 92
    c.drawString(x0+30, ty, show); c.setStrokeColor(teal); c.setLineWidth(0.8); c.line(x0+30, ty-3, x0+30+tw, ty-3)
    c.linkURL(a['url'], (x0+30, ty-5, x0+30+tw, ty+13), relative=0)
    bw, bh = 260, 44; bx = x0 + 30; by = y0 - boxh + 22
    c.setFillColor(teal); c.roundRect(bx, by, bw, bh, 8, stroke=0, fill=1)
    c.setFillColor(white); c.setFont('B', 15); c.drawCentredString(bx+bw/2, by+15, a['bouton'] + '  →')
    c.linkURL(a['url'], (bx, by, bx+bw, by+bh), relative=0)
boxh = 2*boxh + gap

c.setFillColor(ink); c.setFont('B', 13); y = top - boxh - 40; c.drawString(40, y, 'Comment faire')
c.setFont('R', 11.5); c.setFillColor(muted)
for k, l in enumerate(["1. Connectez-vous avec votre compte Veolia sur le téléphone.",
                       "2. Saisie : remplissez la réclamation, envoyez-la et notez le code de suivi affiché.",
                       "3. Traitement : après l'intervention, saisissez ce code de suivi puis le traitement."]):
    c.drawString(40, y-22-k*19, l)
c.setFont('R', 9); c.drawCentredString(W/2, 22, 'Les réclamations et traitements sont transmis uniquement au responsable Ré-Use.')
c.showPage(); c.save()
pymupdf.open(SORTIE + '.pdf')[0].get_pixmap(dpi=110).save(SORTIE + '.png')
print('fiche OK')
