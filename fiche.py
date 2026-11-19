"""Génère la fiche A4 (PDF cliquable + PNG) et le QR code du formulaire Ré-Use."""
import qrcode, pymupdf
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader
URL = "https://tinyurl.com/reuse-reclamations"; SHOW = "tinyurl.com/reuse-reclamations"
pdfmetrics.registerFont(TTFont('B', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('R', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=20, border=2); q.add_data(URL); q.make(fit=True)
qi = q.make_image(fill_color="#0b3f4d", back_color="white").convert("RGB"); qi.save('qr-code-seul.png')
W, H = A4; c = canvas.Canvas('fiche-reclamations-reuse.pdf', pagesize=A4); c.setTitle('Réclamations Ré-Use')
teal = HexColor('#0f6e86'); ink = HexColor('#14262c'); muted = HexColor('#5b7078')
c.setFillColor(teal); c.rect(0, H-140, W, 140, stroke=0, fill=1)
c.setFillColor(white); c.setFont('B', 36); c.drawCentredString(W/2, H-75, 'Réclamations Ré-Use')
c.setFont('R', 15); c.setFillColor(HexColor('#dff1f5')); c.drawCentredString(W/2, H-108, 'Saisie terrain des réclamations clients')
s = 290; x = (W-s)/2; y = H-140-30-s
c.drawImage(ImageReader(qi), x, y, s, s); c.linkURL(URL, (x, y, x+s, y+s), relative=0)
c.setFillColor(ink); c.setFont('B', 17); c.drawCentredString(W/2, y-36, 'Scannez le QR code ou cliquez dessus')
c.setFillColor(muted); c.setFont('R', 13); c.drawCentredString(W/2, y-60, 'ou ouvrez :')
c.setFillColor(teal); c.setFont('B', 22); tw = c.stringWidth(SHOW, 'B', 22); c.drawCentredString(W/2, y-88, SHOW)
c.setStrokeColor(teal); c.setLineWidth(1); c.line((W-tw)/2, y-92, (W+tw)/2, y-92); c.linkURL(URL, ((W-tw)/2, y-95, (W+tw)/2, y-68), relative=0)
bw, bh = 260, 46; bx = (W-bw)/2; by = y-160
c.setFillColor(teal); c.roundRect(bx, by, bw, bh, 8, stroke=0, fill=1); c.setFillColor(white); c.setFont('B', 16)
c.drawCentredString(W/2, by+16, 'Ouvrir le formulaire'); c.linkURL(URL, (bx, by, bx+bw, by+bh), relative=0)
c.setFillColor(ink); c.setFont('B', 13); ty = by-38; c.drawString(70, ty, 'Comment faire')
c.setFont('R', 12); c.setFillColor(muted)
for i, l in enumerate(["1. Connectez-vous avec votre compte Veolia sur le téléphone.",
                       "2. Sur le lieu, appuyez sur « Enregistrer ma position ».",
                       "3. Remplissez le formulaire et ajoutez une photo.",
                       "4. Appuyez sur « Envoyer la réclamation » : un message de confirmation s'affiche."]):
    c.drawString(70, ty-22-i*19, l)
c.setFont('R', 9); c.drawCentredString(W/2, 22, 'Les réclamations sont transmises uniquement au responsable Ré-Use.')
c.showPage(); c.save()
pymupdf.open('fiche-reclamations-reuse.pdf')[0].get_pixmap(dpi=110).save('fiche-reclamations-reuse.png')
print('fiche OK')
