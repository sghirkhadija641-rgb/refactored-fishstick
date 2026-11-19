"""Génère la fiche A4 des formulaires Ré-Use : un PDF cliquable (et une image PNG) avec les deux accès,
« Saisie d'une réclamation » et « Traitement de réclamation », chacun avec son QR code, son lien et son bouton."""
import qrcode, pymupdf
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader

ACCES = [
    dict(titre="Saisie d'une réclamation", url="https://tinyurl.com/reuse-reclamations",
         texte=["Enregistrer une nouvelle réclamation.", "Un code de suivi s'affiche à l'envoi."],
         bouton="Ouvrir la saisie", qr="qr-code-seul.png"),
    dict(titre="Traitement de réclamation", url="https://tinyurl.com/reuse-traitement",
         texte=["Après l'intervention, avec le", "code de suivi de la réclamation."],
         bouton="Ouvrir le traitement", qr="qr-code-traitement.png"),
]
SORTIE = 'fiche-reclamations-reuse'

pdfmetrics.registerFont(TTFont('B', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('R', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
teal = HexColor('#0f6e86'); ink = HexColor('#14262c'); muted = HexColor('#5b7078'); line = HexColor('#d6e0e3'); soft = HexColor('#f3f6f7')

W, H = A4
c = canvas.Canvas(SORTIE + '.pdf', pagesize=A4); c.setTitle('Réclamations Ré-Use')
c.setFillColor(teal); c.rect(0, H-130, W, 130, stroke=0, fill=1)
c.setFillColor(white); c.setFont('B', 34); c.drawCentredString(W/2, H-70, 'Réclamations Ré-Use')
c.setFont('R', 14); c.setFillColor(HexColor('#dff1f5')); c.drawCentredString(W/2, H-100, 'Choisissez le formulaire : scannez le QR code ou cliquez dessus')

colw = (W - 3*30) / 2; top = H - 160; boxh = 412
for i, a in enumerate(ACCES):
    x0 = 30 + i*(colw + 30)
    c.setFillColor(soft); c.setStrokeColor(line); c.roundRect(x0, top-boxh, colw, boxh, 12, stroke=1, fill=1)
    c.setFillColor(teal); c.roundRect(x0, top-6, colw, 6, 3, stroke=0, fill=1)
    c.setFillColor(ink); c.setFont('B', 16); c.drawCentredString(x0+colw/2, top-38, a['titre'])
    c.setFillColor(muted); c.setFont('R', 10.5)
    for j, l in enumerate(a['texte']):
        c.drawCentredString(x0+colw/2, top-60-j*14, l)
    q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=20, border=2); q.add_data(a['url']); q.make(fit=True)
    img = q.make_image(fill_color="#0b3f4d", back_color="white").convert("RGB"); img.save(a['qr'])
    s = colw - 50; qx = x0 + 25; qy = top - 100 - s
    c.drawImage(ImageReader(img), qx, qy, s, s); c.linkURL(a['url'], (qx, qy, qx+s, qy+s), relative=0)
    show = a['url'].replace('https://', '')
    c.setFillColor(teal); c.setFont('B', 12.5); tw = c.stringWidth(show, 'B', 12.5); ty = qy - 26
    c.drawCentredString(x0+colw/2, ty, show)
    c.setStrokeColor(teal); c.setLineWidth(0.8); c.line(x0+colw/2-tw/2, ty-3, x0+colw/2+tw/2, ty-3)
    c.linkURL(a['url'], (x0+colw/2-tw/2, ty-5, x0+colw/2+tw/2, ty+12), relative=0)
    bw, bh = colw - 50, 40; bx = x0 + 25; by = ty - 62
    c.setFillColor(teal); c.roundRect(bx, by, bw, bh, 8, stroke=0, fill=1)
    c.setFillColor(white); c.setFont('B', 14); c.drawCentredString(x0+colw/2, by+14, a['bouton'])
    c.linkURL(a['url'], (bx, by, bx+bw, by+bh), relative=0)

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
