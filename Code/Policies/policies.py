#writes the Group Policy templates of the Office add-in: ADMX/ADML (Vista and later) and ADM (Windows 2000 and XP)
import sys, os
from xml.sax.saxutils import escape

root = os.path.join(sys.argv[1], 'Policies')
KEY = 'Software\\Policies\\SystemBanner'

LANGS = ['en-US', 'fr-FR', 'de-DE']

T = {
 'category': ('SystemBanner Office add-in', 'Complément Office SystemBanner', 'SystemBanner Office-Add-in'),
 'category_help': (
  'Settings of the SystemBanner Office add-in. The classification of the computer (Simple) is set by the SystemBanner template.',
  'Réglages du complément Office SystemBanner. La classification du poste (Simple) se règle avec le modèle de SystemBanner.',
  'Einstellungen des SystemBanner Office-Add-ins. Die Einstufung des Rechners (Simple) wird mit der Vorlage von SystemBanner eingestellt.'),
 'supported': ('Office 2000 to 2024 and Microsoft 365, with the SystemBanner Office add-in', 'Office 2000 à 2024 et Microsoft 365, avec le complément Office SystemBanner', 'Office 2000 bis 2024 und Microsoft 365, mit dem SystemBanner Office-Add-in'),
}

#name, value name, title, help
BOOLS = [
 ('RequireClassification', 'OfficeRequireClassification',
  ('Require a classification to save and print', 'Exiger une classification pour enregistrer et imprimer', 'Einstufung vor dem Speichern und Drucken verlangen'),
  ('Enabled: a document with no classification cannot be saved or printed in Word and Excel; the add-in asks to classify it first.\n\nDisabled or not configured: documents are saved and printed as they are.',
   'Activé : un document sans classification ne peut être ni enregistré ni imprimé dans Word et Excel ; le complément demande d\'abord de le classifier.\n\nDésactivé ou non configuré : les documents sont enregistrés et imprimés tels quels.',
   'Aktiviert: Ein Dokument ohne Einstufung kann in Word und Excel weder gespeichert noch gedruckt werden; das Add-in verlangt zuerst eine Einstufung.\n\nDeaktiviert oder nicht konfiguriert: Dokumente werden unverändert gespeichert und gedruckt.')),
 ('SpellOut', 'OfficeSpellOut',
  ('Spell control markings out in IC banner lines', 'Écrire en entier les marquages de contrôle IC', 'IC-Kontrollkennzeichnungen ausschreiben'),
  ('Enabled: IC banner lines spell control markings out (TALENT KEYHOLE) instead of using their abbreviations (TK).\n\nDisabled or not configured: abbreviations, as most IC documents use them.',
   'Activé : les lignes de bannière IC écrivent les marquages de contrôle en entier (TALENT KEYHOLE) au lieu de leurs abréviations (TK).\n\nDésactivé ou non configuré : abréviations, comme dans la plupart des documents IC.',
   'Aktiviert: IC-Bannerzeilen schreiben Kontrollkennzeichnungen aus (TALENT KEYHOLE) statt sie abzukürzen (TK).\n\nDeaktiviert oder nicht konfiguriert: Abkürzungen, wie in den meisten IC-Dokumenten.')),
 ('AnalyticProducts', 'OfficeAnalyticProducts',
  ('Documents are disseminated analytic products', 'Les documents sont des produits d\'analyse diffusés', 'Dokumente sind verbreitete Analyseprodukte'),
  ('Enabled: every classified IC portion must carry a foreign disclosure and release marking (ICD 710 E.5); the roll-up reports the portions without one.\n\nDisabled or not configured: the marking is not required.',
   'Activé : chaque partie IC classifiée doit porter un marquage de communication à l\'étranger (ICD 710 E.5) ; le calcul signale les parties qui n\'en ont pas.\n\nDésactivé ou non configuré : ce marquage n\'est pas exigé.',
   'Aktiviert: Jeder eingestufte IC-Teil muss eine Kennzeichnung zur Weitergabe ins Ausland tragen (ICD 710 E.5); die Ableitung meldet Teile ohne sie.\n\nDeaktiviert oder nicht konfiguriert: Diese Kennzeichnung ist nicht verlangt.')),
 ('BannerBackground', 'OfficeBannerBackground',
  ('IC banner lines on a colored background', 'Lignes de bannière IC sur fond coloré', 'IC-Bannerzeilen auf farbigem Hintergrund'),
  ('Enabled: IC banner lines on the colored background of SystemBanner, as on screen, with white or black text.\n\nDisabled or not configured: bold text in the color of the level, with no background, readable in black and white printing. The IC Markings System Register sets no color.',
   'Activé : lignes de bannière IC sur le fond coloré de SystemBanner, comme à l\'écran, en texte blanc ou noir.\n\nDésactivé ou non configuré : texte gras de la couleur du niveau, sans fond, lisible en impression noir et blanc. Le registre IC n\'impose aucune couleur.',
   'Aktiviert: IC-Bannerzeilen auf dem farbigen Hintergrund von SystemBanner, wie am Bildschirm, mit weißem oder schwarzem Text.\n\nDeaktiviert oder nicht konfiguriert: fetter Text in der Farbe des Geheimhaltungsgrads, ohne Hintergrund, im Schwarzweißdruck lesbar. Das IC-Register schreibt keine Farbe vor.')),
 ('Toolbar', 'OfficeToolbar',
  ('Show the toolbar in Office 2007 and later', 'Afficher la barre d\'outils dans Office 2007 et suivants', 'Symbolleiste in Office 2007 und höher anzeigen'),
  ('Enabled: the SystemBanner toolbar of Office 2000 to 2003 also shows in Office 2007 and later (Add-ins tab), besides the ribbon tab.\n\nDisabled or not configured: the Classification ribbon tab only.',
   'Activé : la barre d\'outils SystemBanner d\'Office 2000 à 2003 s\'affiche aussi dans Office 2007 et suivants (onglet Compléments), en plus de l\'onglet du ruban.\n\nDésactivé ou non configuré : l\'onglet Classification du ruban seul.',
   'Aktiviert: Die SystemBanner-Symbolleiste von Office 2000 bis 2003 erscheint auch in Office 2007 und höher (Registerkarte Add-Ins), neben der Registerkarte im Menüband.\n\nDeaktiviert oder nicht konfiguriert: nur die Registerkarte Klassifizierung.')),
]

LANG_POLICY = ('Language', 'OfficeLanguage',
  ('Language of the add-in', 'Langue du complément', 'Sprache des Add-ins'),
  ('Enabled: the dialogs, buttons and messages of the add-in in the chosen language. The markings themselves are never translated.\n\nDisabled or not configured: the language of the Office user interface (English, French or German, else English). A user may also set OfficeLanguage in HKEY_CURRENT_USER\\Software\\SystemBanner.',
   'Activé : les fenêtres, boutons et messages du complément dans la langue choisie. Les marquages eux-mêmes ne sont jamais traduits.\n\nDésactivé ou non configuré : la langue de l\'interface d\'Office (anglais, français ou allemand, sinon anglais). Un utilisateur peut aussi régler OfficeLanguage dans HKEY_CURRENT_USER\\Software\\SystemBanner.',
   'Aktiviert: Dialoge, Schaltflächen und Meldungen des Add-ins in der gewählten Sprache. Die Kennzeichnungen selbst werden nie übersetzt.\n\nDeaktiviert oder nicht konfiguriert: die Sprache der Office-Oberfläche (Englisch, Französisch oder Deutsch, sonst Englisch). Ein Benutzer kann OfficeLanguage auch in HKEY_CURRENT_USER\\Software\\SystemBanner setzen.'),
  [('en', ('English', 'Anglais', 'Englisch')), ('fr', ('French', 'Français', 'Französisch')), ('de', ('German', 'Allemand', 'Deutsch'))],
  ('Language:', 'Langue :', 'Sprache:'))

def admx():
    out = ['<?xml version="1.0" encoding="utf-8"?>',
           '<!-- SystemBanner Office add-in: Group Policy template, generated with the ADML files of en-US, fr-FR and de-DE -->',
           '<policyDefinitions xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" revision="1.0" schemaVersion="1.0" xmlns="http://schemas.microsoft.com/GroupPolicy/2006/07/PolicyDefinitions">',
           '  <policyNamespaces>',
           '    <target prefix="systembanneroffice" namespace="SystemBanner.Office.Policies" />',
           '    <using prefix="windows" namespace="Microsoft.Policies.Windows" />',
           '  </policyNamespaces>',
           '  <resources minRequiredRevision="1.0" />',
           '  <supportedOn>',
           '    <definitions>',
           '      <definition name="SUPPORTED_SystemBannerOffice" displayName="$(string.SUPPORTED_SystemBannerOffice)" />',
           '    </definitions>',
           '  </supportedOn>',
           '  <categories>',
           '    <category name="SystemBannerOffice" displayName="$(string.SystemBannerOffice)" explainText="$(string.SystemBannerOffice_Help)" />',
           '  </categories>',
           '  <policies>']
    for name, value, _, _ in BOOLS:
        out += ['    <policy name="%s" class="Machine" displayName="$(string.%s)" explainText="$(string.%s_Help)" key="%s" valueName="%s">' % (name, name, name, KEY, value),
                '      <parentCategory ref="SystemBannerOffice" />',
                '      <supportedOn ref="SUPPORTED_SystemBannerOffice" />',
                '      <enabledValue><decimal value="1" /></enabledValue>',
                '      <disabledValue><decimal value="0" /></disabledValue>',
                '    </policy>']
    name, value, _, _, items, _ = LANG_POLICY
    out += ['    <policy name="%s" class="Machine" displayName="$(string.%s)" explainText="$(string.%s_Help)" presentation="$(presentation.%s)" key="%s">' % (name, name, name, name, KEY),
            '      <parentCategory ref="SystemBannerOffice" />',
            '      <supportedOn ref="SUPPORTED_SystemBannerOffice" />',
            '      <elements>',
            '        <enum id="%s" valueName="%s" required="true">' % (name, value)]
    for code, _ in items:
        out += ['          <item displayName="$(string.%s_%s)"><value><string>%s</string></value></item>' % (name, code, code)]
    out += ['        </enum>', '      </elements>', '    </policy>', '  </policies>', '</policyDefinitions>', '']
    return '\r\n'.join(out)

def adml(i):
    s = []
    def add(id, text):
        s.append('      <string id="%s">%s</string>' % (id, escape(text).replace('\n', '&#xD;&#xA;')))
    add('SystemBannerOffice', T['category'][i])
    add('SystemBannerOffice_Help', T['category_help'][i])
    add('SUPPORTED_SystemBannerOffice', T['supported'][i])
    for name, _, title, help in BOOLS:
        add(name, title[i])
        add(name + '_Help', help[i])
    name, _, title, help, items, label = LANG_POLICY
    add(name, title[i])
    add(name + '_Help', help[i])
    for code, names in items:
        add(name + '_' + code, names[i])
    out = ['<?xml version="1.0" encoding="utf-8"?>',
           '<policyDefinitionResources xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" revision="1.0" schemaVersion="1.0" xmlns="http://schemas.microsoft.com/GroupPolicy/2006/07/PolicyDefinitions">',
           '  <displayName>%s</displayName>' % escape(T['category'][i]),
           '  <description>%s</description>' % escape(T['category_help'][i]),
           '  <resources>',
           '    <stringTable>'] + s + [
           '    </stringTable>',
           '    <presentationTable>',
           '      <presentation id="%s">' % name,
           '        <dropdownList refId="%s" noSort="true" defaultItem="0">%s</dropdownList>' % (name, escape(label[i])),
           '      </presentation>',
           '    </presentationTable>',
           '  </resources>',
           '</policyDefinitionResources>', '']
    return '\r\n'.join(out)

def adm(i):
    strings = []
    def sid(id, text):
        strings.append('%s="%s"' % (id, text.replace('"', "'").replace('\n', '\\n')))
    out = ['; SystemBanner Office add-in: Group Policy template for Windows 2000 and XP (Add/Remove Templates).',
           '; Windows Vista and later: use SystemBannerOffice.admx and its language folders instead.',
           '',
           'CLASS MACHINE',
           '',
           'CATEGORY !!Category',
           '  KEYNAME "%s"' % KEY]
    sid('Category', T['category'][i])
    for name, value, title, help in BOOLS:
        out += ['  POLICY !!%s' % name,
                '    EXPLAIN !!%s_Help' % name,
                '    VALUENAME "%s"' % value,
                '    VALUEON NUMERIC 1',
                '    VALUEOFF NUMERIC 0',
                '  END POLICY']
        sid(name, title[i])
        sid(name + '_Help', help[i] + '\n\n' + T['supported'][i] + '.')
    name, value, title, help, items, label = LANG_POLICY
    out += ['  POLICY !!%s' % name,
            '    EXPLAIN !!%s_Help' % name,
            '    PART !!%s_Label DROPDOWNLIST REQUIRED' % name,
            '      VALUENAME "%s"' % value,
            '      ITEMLIST']
    for code, _ in items:
        out += ['        NAME !!%s_%s VALUE "%s"' % (name, code, code)]
    out += ['      END ITEMLIST', '    END PART', '  END POLICY', 'END CATEGORY', '', '[strings]']
    sid(name, title[i])
    sid(name + '_Help', help[i] + '\n\n' + T['supported'][i] + '.')
    sid(name + '_Label', label[i])
    for code, names in items:
        sid(name + '_' + code, names[i])
    return '\r\n'.join(out + strings + [''])

os.makedirs(root, exist_ok=True)
open(os.path.join(root, 'SystemBannerOffice.admx'), 'w', encoding='utf-8', newline='').write(admx())
for i, lang in enumerate(LANGS):
    os.makedirs(os.path.join(root, lang), exist_ok=True)
    open(os.path.join(root, lang, 'SystemBannerOffice.adml'), 'w', encoding='utf-8', newline='').write(adml(i))
    #ADM files with accents are read as Unicode (UTF-16 LE with its byte order mark)
    suffix = '' if i == 0 else '_' + lang[:2].upper()
    open(os.path.join(root, 'SystemBannerOffice%s.adm' % suffix), 'w', encoding='utf-16', newline='').write(adm(i))
print('written to', root)
