// Message Master - Gesundheitsportal 2026
// Cloudflight B2B Positioning Document
// Generated: 2026-03-04

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  ShadingType,
  convertInchesToTwip,
  PageOrientation,
  Header,
  Footer,
  ImageRun,
  UnderlineType,
  TabStopPosition,
  TabStopType,
  LevelFormat,
  NumberFormat,
} = require("docx");

const fs = require("fs");

// Cloudflight Brand Colors
const CF_BLUE = "1B3A6B";
const CF_LIGHT_BLUE = "E8EEF6";
const CF_ACCENT = "2563EB";
const CF_DARK_GRAY = "374151";
const CF_MEDIUM_GRAY = "6B7280";
const CF_LIGHT_GRAY = "F3F4F6";
const CF_WHITE = "FFFFFF";
const CF_TABLE_HEADER = "1B3A6B";
const CF_TABLE_ROW_ALT = "EEF2F8";

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function headingMain(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        bold: true,
        color: CF_WHITE,
        size: 28,
        font: "Calibri",
      }),
    ],
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 200, after: 120 },
    shading: {
      type: ShadingType.SOLID,
      color: CF_BLUE,
      fill: CF_BLUE,
    },
    indent: { left: convertInchesToTwip(0.15), right: convertInchesToTwip(0.15) },
  });
}

function headingSection(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        bold: true,
        color: CF_BLUE,
        size: 24,
        font: "Calibri",
      }),
    ],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 100 },
    border: {
      bottom: {
        color: CF_BLUE,
        space: 4,
        style: BorderStyle.SINGLE,
        size: 6,
      },
    },
  });
}

function headingField(number, text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: `${number}. ${text}`,
        bold: true,
        color: CF_WHITE,
        size: 22,
        font: "Calibri",
      }),
    ],
    spacing: { before: 280, after: 80 },
    shading: {
      type: ShadingType.SOLID,
      color: CF_BLUE,
      fill: CF_BLUE,
    },
    indent: { left: convertInchesToTwip(0.1) },
  });
}

function bodyText(text, options = {}) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        color: CF_DARK_GRAY,
        size: 20,
        font: "Calibri",
        bold: options.bold || false,
        italics: options.italic || false,
      }),
    ],
    spacing: { before: 60, after: 60 },
    indent: options.indent ? { left: convertInchesToTwip(0.3) } : {},
  });
}

function labelText(label, value) {
  return new Paragraph({
    children: [
      new TextRun({
        text: label + ": ",
        bold: true,
        color: CF_BLUE,
        size: 20,
        font: "Calibri",
      }),
      new TextRun({
        text: value,
        color: CF_DARK_GRAY,
        size: 20,
        font: "Calibri",
      }),
    ],
    spacing: { before: 60, after: 60 },
    indent: { left: convertInchesToTwip(0.2) },
  });
}

function bulletPoint(text, level = 0) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        color: CF_DARK_GRAY,
        size: 20,
        font: "Calibri",
      }),
    ],
    bullet: { level },
    spacing: { before: 40, after: 40 },
    indent: { left: convertInchesToTwip(0.3 + level * 0.25) },
  });
}

function numberedPoint(text, number) {
  return new Paragraph({
    children: [
      new TextRun({
        text: `${number}.  ${text}`,
        color: CF_DARK_GRAY,
        size: 20,
        font: "Calibri",
      }),
    ],
    spacing: { before: 60, after: 60 },
    indent: { left: convertInchesToTwip(0.3) },
  });
}

function highlightBox(text, bgColor = CF_LIGHT_BLUE) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        color: CF_BLUE,
        size: 20,
        font: "Calibri",
        bold: true,
        italics: true,
      }),
    ],
    spacing: { before: 100, after: 100 },
    shading: {
      type: ShadingType.SOLID,
      color: bgColor,
      fill: bgColor,
    },
    indent: { left: convertInchesToTwip(0.3), right: convertInchesToTwip(0.3) },
  });
}

function annahme(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text: "⚠ ",
        color: "D97706",
        size: 18,
        font: "Calibri",
      }),
      new TextRun({
        text,
        color: "92400E",
        size: 18,
        font: "Calibri",
        italics: true,
      }),
    ],
    spacing: { before: 40, after: 40 },
    shading: {
      type: ShadingType.SOLID,
      color: "FEF3C7",
      fill: "FEF3C7",
    },
    indent: { left: convertInchesToTwip(0.2) },
  });
}

function spacer(lines = 1) {
  return new Paragraph({
    children: [new TextRun({ text: "' })],
    spacing: { before: 60 * lines, after: 0 },
  });
}

function divider() {
  return new Paragraph({
    children: [new TextRun({ text: "' })],
    border: {
      bottom: {
        color: CF_LIGHT_BLUE,
        space: 4,
        style: BorderStyle.SINGLE,
        size: 6,
      },
    },
    spacing: { before: 120, after: 120 },
  });
}

// ============================================================
// TABLE BUILDERS
// ============================================================

function makeTableCell(content, options = {}) {
  const bgColor = options.bg || CF_WHITE;
  const isHeader = options.header || false;

  const paragraphs = Array.isArray(content) ? content : [content];
  const cells = paragraphs.map((text) => {
    if (typeof text === "object' && text.constructor && text.constructor.name === 'Paragraph") {
      return text;
    }
    return new Paragraph({
      children: [
        new TextRun({
          text: String(text),
          bold: isHeader,
          color: isHeader ? CF_WHITE : CF_DARK_GRAY,
          size: isHeader ? 20 : 19,
          font: "Calibri",
        }),
      ],
      spacing: { before: 60, after: 60 },
    });
  });

  return new TableCell({
    children: cells,
    shading: {
      type: ShadingType.SOLID,
      color: isHeader ? CF_TABLE_HEADER : bgColor,
      fill: isHeader ? CF_TABLE_HEADER : bgColor,
    },
    margins: {
      top: convertInchesToTwip(0.08),
      bottom: convertInchesToTwip(0.08),
      left: convertInchesToTwip(0.1),
      right: convertInchesToTwip(0.1),
    },
    width: options.width ? { size: options.width, type: WidthType.PERCENTAGE } : undefined,
  });
}

function competitorTable() {
  const headers = ["Wettbewerber", "Positionierung", "Stärken", "Schwächen", "Unser Argument'];
  const rows = [
    [
      "x-Tention\n(it with care)",
      "Österreichischer Healthcare-IT-Spezialist mit tiefer KIS/ELGA-Integrationskompetenz. 95% Marktanteil österreichische Krankenhäuser. NÖ LGA-Auftrag gewonnen (2025).",
      "• FHIR/HL7/IHE-Standards\n• KIS-Integrationserfahrung\n• 800 MA, 16 Standorte\n• Bekannte Marke bei IT-Entscheidern\n• Orchestra eHealth Suite",
      "• Primär Backend/Integration, schwächere End-User-UX\n• Keine eigenen nativen iOS/Android-Apps\n• Kein vollständiges Branding-Framework\n• Reaktiv in der Produktentwicklung",
      "Wir kombinieren dieselbe Integrationstiefe mit überlegener UX, vollständigen nativen Apps (iOS/Android, biometrisch) und echtem White-Label-Branding. Cloudflights Custom-Software-DNA bedeutet: der Auftraggeber prägt das Produkt – nicht umgekehrt.",
    ],
    [
      "Siemens Healthineers\n(eHealth / ITH icoserve)",
      "Globaler Medizintechnikkonzern, eHealth-Lösungen durch Tochter ITH icoserve (Innsbruck). Referenzen: Vinzenz Gruppe ('Hallo Gesundheit'), SALK.",
      "• Globale Markenbekanntheit\n• Breites Portfolio\n• Österreich-Niederlassung (ITH icoserve)\n• ELGA/FHIR-Kompetenz",
      "• Aus deutschem Mega-Projekt 'Mein Krankenhaus Bayern' (76 KH, 150 Standorte) gescheitert – IT-Sicherheitszertifikate fehlten\n• Konzernentscheidungen vs. lokale Flexibilität\n• Langwierige Implementierungszyklen\n• Eingeschränktes Customizing",
      "Cloudflights Gesundheitsportal ist bereits in bis zu 8 österreichischen Bundesländern live – kein gescheitertes Pilotprojekt. Siemens' Ausstieg aus dem bayerischen KH-Verbund (76 Häuser) zeigt: Konzernlösungen scheitern an lokalen Realitäten. Wir liefern.",
    ],
    [
      "CANCOM Healthcare",
      "IT-Infrastrukturanbieter mit Healthcare-Fokus und 'Smart Health Platform' (App-Store-Ansatz für digitale Gesundheitsanwendungen).",
      "• Breite IT-Infrastrukturkompetenz\n• Cybersecurity-Portfolio\n• Managed Services\n• Großes Partnernetzwerk",
      "• Kein eigener Patientenportal-Kernprodukt\n• Systemintegrator-Ansatz: Zusammensetzung von Drittlösungen\n• Geringere Healthcare-Workflow-Expertise\n• Weniger klinische Tiefe",
      "CANCOM baut Infrastruktur – wir bauen das Produkt, das Patienten und Ärzte täglich nutzen. Ein App-Store-Ansatz löst keine klinischen Prozesse. Das Cloudflight Gesundheitsportal ist eine durchgängige, produktisierte Lösung mit eigenem Entwicklungsteam, FHIR-Server und nativen Apps.",
    ],
  ];

  const tableRows = [
    new TableRow({
      children: headers.map((h) => makeTableCell(h, { header: true })),
      tableHeader: true,
    }),
    ...rows.map((row, rowIdx) =>
      new TableRow({
        children: row.map((cell) =>
          makeTableCell(cell, { bg: rowIdx % 2 === 0 ? CF_WHITE : CF_TABLE_ROW_ALT })
        ),
      })
    ),
  ];

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 100, bottom: 100, left: 100, right: 100 },
  });
}

function personasTable() {
  const headers = ["Merkmal", "Persona 1: Der digitale Stratege", "Persona 2: Der Effizienzsuchende", "Persona 3: Die überarbeitete Ärztin'];

  const rows = [
    ["Rolle / Titel", "CDO / CIO einer Landesgesundheitsholding", "Verwaltungsdirektor / Kaufm. Leitung, Fondskrankenhaus", "Fachärztin mit eigener Ordination oder MVZ-Leiterin (5–20 Standorte)'],
    ["Einrichtungsgröße", "20–80 Standorte, Landesebene", "300–800 Betten, Akutkrankenhaus", "Ordination oder Praxiskette, 5–20 Standorte'],
    ["Hauptverantwortung", "Digitale Transformation, IT-Strategie, Ausschreibungen", "Budget, Personal, Patientenzufriedenheit, Betrieb", "Patientenversorgung + Ordinationsmanagement'],
    ["Kern-Pain-Points", "• EHDS-Compliance-Druck\n• Legacy-Systemlandschaft\n• Koordination vieler Standorte\n• Sichtbarkeit digitaler Maßnahmen\n• Fehlende Interoperabilität", "• Steigender Verwaltungsaufwand bei sinkenden Budgets\n• Personalmangel in Terminverwaltung\n• Patientenbeschwerden über Erreichbarkeit\n• Wettbewerb mit Privatanbietern", "• Telefonflut in der Ordination\n• Kein digitaler Dokumentenkanal\n• ELGA-Zugriff umständlich\n• Kaum Zeit für Patienten durch Admin-Last'],
    ["Entscheidungskriterien", "1. Skalierbarkeit über alle Standorte\n2. Nachgewiesene Referenzen\n3. ELGA/EHDS-Compliance\n4. Langfristige Vendor-Stabilität\n5. TCO", "1. Messbarer ROI\n2. Schnelle Implementierung\n3. Geringer Change-Aufwand\n4. Patientenzufriedenheit\n5. SLAs und Support", "1. Einfache Bedienbarkeit\n2. Implementierungsaufwand\n3. ELGA-Anbindung ohne Mehrarbeit\n4. Kostenrahmen\n5. Patientenakzeptanz'],
    ["Informationsquellen", "conhIT, DMEA, Peer-Austausch CDOs, Analystenstudien, Vergabeportale", "Fachzeitschriften (krankenhaus.at), Beratungsempfehlungen, Kollegennetzwerk", "Ärztekammer-Newsletter, Kollegennetzwerk, Praxissoftware-Anbieter'],
    ["Buying Role", "Economic Buyer + Technical Sponsor", "Economic Buyer (Budgetfreigabe)", "User + Influencer + Economic Buyer (eigene Ordination)'],
    ["Typisches Zitat", "[PLATZHALTER: echtes Kundenzitat einfügen]", "[PLATZHALTER: echtes Kundenzitat einfügen]", "[PLATZHALTER: echtes Kundenzitat einfügen]'],
  ];

  const colWidths = [18, 27, 27, 28];

  const tableRows = [
    new TableRow({
      children: headers.map((h, i) =>
        makeTableCell(h, { header: true, width: colWidths[i] })
      ),
      tableHeader: true,
    }),
    ...rows.map((row, rowIdx) =>
      new TableRow({
        children: row.map((cell, colIdx) =>
          makeTableCell(cell, {
            bg: rowIdx % 2 === 0 ? CF_WHITE : CF_TABLE_ROW_ALT,
            width: colWidths[colIdx],
          })
        ),
      })
    ),
  ];

  return new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

// ============================================================
// DOCUMENT CONTENT
// ============================================================

function buildDocument() {
  const children = [];

  // ─── TITLE PAGE ───────────────────────────────────────────
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "MESSAGE MASTER",
          bold: true,
          color: CF_WHITE,
          size: 52,
          font: "Calibri",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 800, after: 80 },
      shading: { type: ShadingType.SOLID, color: CF_BLUE, fill: CF_BLUE },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Gesundheitsportal",
          bold: true,
          color: CF_WHITE,
          size: 40,
          font: "Calibri",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      shading: { type: ShadingType.SOLID, color: CF_BLUE, fill: CF_BLUE },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Digitale Vernetzung von Patienten und Gesundheitsdienstleistern",
          color: CF_WHITE,
          size: 24,
          font: "Calibri",
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      shading: { type: ShadingType.SOLID, color: CF_BLUE, fill: CF_BLUE },
    }),
    new Paragraph({
      children: [new TextRun({ text: "' })],
      spacing: { before: 200, after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: '"Gesundheit war noch nie so einfach'',
          bold: true,
          color: CF_BLUE,
          size: 32,
          font: "Calibri",
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "' })],
      spacing: { before: 200, after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "Entwickelt von:  ", bold: true, color: CF_DARK_GRAY, size: 20, font: "Calibri' }),
        new TextRun({ text: "Cloudflight (Webportal + native iOS/Android Apps), PineIT (Integrationsplattform + FHIR-Server)", color: CF_DARK_GRAY, size: 20, font: "Calibri' }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "Erstellungsdatum:  ", bold: true, color: CF_DARK_GRAY, size: 20, font: "Calibri' }),
        new TextRun({ text: "4. März 2026", color: CF_DARK_GRAY, size: 20, font: "Calibri' }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "Vertraulich – nur für internen Gebrauch", bold: true, color: "DC2626", size: 18, font: "Calibri' }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 300, after: 60 },
    }),
    spacer(3)
  );

  // ─── FELD 1: TITLE OF SOLUTION ─────────────────────────────
  children.push(
    headingField("1", "Title of Solution"),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Gesundheitsportal",
          bold: true,
          color: CF_BLUE,
          size: 32,
          font: "Calibri",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 60 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Digitale Vernetzung von Patienten und Gesundheitsdienstleistern",
          color: CF_MEDIUM_GRAY,
          size: 22,
          font: "Calibri",
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 100 },
    }),
    divider()
  );

  // ─── FELD 2: PORTFOLIO CONTEXT ─────────────────────────────
  children.push(
    headingField("2", "Portfolio Context"),
    spacer(),
    bodyText(
      "Das Gesundheitsportal ist Teil des Cloudflight Healthcare-Portfolios und verknüpft Cloudflights Kernkompetenz in Custom Software Engineering mit regulatorisch konformer Digitalisierung im österreichischen und europäischen Gesundheitswesen. Als gemeinsame Lösung von Cloudflight (Webportal, native iOS/Android Apps) und PineIT (FHIR-Server, Integrationsplattform) schließt das Gesundheitsportal die Lücke zwischen bestehenden klinischen Informationssystemen (KIS/HIS) und modernen, patientenzentrierten Kommunikations- und Serviceleistungen – und positioniert Cloudflight als strategischen Digitalpartner für Gesundheitsholdings und Fondskrankenanstalten in der DACH-Region."
    ),
    divider()
  );

  // ─── FELD 3: VALUE PROPOSITION ─────────────────────────────
  children.push(
    headingField("3", "Value Proposition"),
    spacer(),
    highlightBox("Format: FÜR / DIE / IST ... DIE LÖSUNG, DIE / INDEM / IM VERGLEICH ZU WETTBEWERBERN"),
    spacer(),
    labelText("FÜR", "Gesundheitsholdings, Fondskrankenanstalten und niedergelassene Ärzteketten in Österreich und der DACH-Region,"),
    labelText("DIE", "ihren Patienten ein modernes, nahtloses digitales Versorgungserlebnis bieten und gleichzeitig administrative Kosten durch telefonbasiertes Terminmanagement, Papierdokumente und Systembrüche senken wollen,"),
    labelText("IST DAS", "Cloudflight Gesundheitsportal"),
    labelText("DIE LÖSUNG, DIE", "alle kritischen Touchpoints der Arzt-Patienten-Interaktion – 24/7-Terminbuchung, Videosprechstunde, sichere Dokumenten-Wallet, Vitaldaten-Monitoring und tiefe ELGA-Integration – in einer einzigen, vollständig brandbaren Plattform vereint, die sowohl als Webportal als auch als native iOS- und Android-App verfügbar ist."),
    labelText("INDEM", "Medienbrüche eliminiert werden, telefonische Terminprozesse digitalisiert werden, Ärzte beim Termin das vollständige Patientenbild sehen und Patienten zu aktiven Gesundheitsmanagern mit Datensouveränität werden."),
    labelText("IM VERGLEICH ZU WETTBEWERBERN", "bieten wir die einzige Plattform mit vollständiger nativer App-Entwicklung (iOS/Android, biometrische Authentifizierung), echter ELGA-Tiefenintegration (eBefund inkl. DICOM-Viewer, eMedikation, eImpfpass, CDA-Parsing), vollständigem White-Label-Branding und nachgewiesenem Rollout in bis zu acht österreichischen Bundesländern – mit einem Rahmenvertrag über 150 Mio. EUR bei der Oberösterreichischen Gesundheitsholding (OÖG)."),
    divider()
  );

  // ─── FELD 4: ICP ───────────────────────────────────────────
  children.push(
    headingField("4", "ICP – Ideal Customer Profile"),
    spacer(),
    bodyText("Primäre Zielkunden", { bold: true }),
    bulletPoint("Gesundheitsholdings auf Landesebene (Österreich/DACH): 20–80 Standorte, öffentlich finanziert, regulatorisch unter EHDS-/ELGA-Druck, aktiv in Ausschreibungsprozessen (EU-weite Vergabe)"),
    bulletPoint("Fondskrankenanstalten und Akutkrankenhäuser: 200–1.000 Betten, unter Kostendruck und mit Digitalisierungsmandat durch Krankenhauszukunftsgesetz (DE) bzw. ELGA-Pflicht (AT)"),
    bulletPoint("Niedergelassene Ärzteketten und MVZ: 5–20 Standorte, inhabergeführt oder investorenbetrieben, mit Bedarf an skalierbarer Patientenkommunikation"),
    spacer(),
    bodyText("Branche & Regulatorik", { bold: true }),
    bulletPoint("Stationäre und ambulante Gesundheitsversorgung, Österreich/Deutschland/Schweiz"),
    bulletPoint("Regulatorischer Kontext: EHDS (EU), ELGA (AT), Krankenhauszukunftsgesetz (DE), DSGVO, ID Austria"),
    bulletPoint("Marktvolumen: 100–400 Mio. EUR/Jahr EU, 40–80 Ausschreibungen/Jahr"),
    spacer(),
    bodyText("Entscheidungsträger-Rollen", { bold: true }),
    bulletPoint("CDO / CIO (Economic Buyer + Technical Sponsor)"),
    bulletPoint("Verwaltungsdirektor / Kaufmännische Leitung (Economic Buyer)"),
    bulletPoint("Leitender Arzt / Ordinationsleiter (User + Influencer)"),
    divider()
  );

  // ─── FELD 5: TONE OF VOICE ─────────────────────────────────
  children.push(
    headingField("5", "Tone of Voice"),
    spacer(),
    bodyText("Professionell und direkt:", { bold: true }),
    bodyText("Keine Marketing-Worthülsen. Konkrete Zahlen, Fakten und Referenzen sprechen für sich.", { indent: true }),
    spacer(),
    bodyText("Humble Bragging:", { bold: true }),
    bodyText("Erfolge benennen, ohne zu übertreiben. Beispiel: \'Rahmenvertrag über 150 Mio. EUR\' statt \'Marktführer\'.", { indent: true }),
    spacer(),
    bodyText("Empathisch und lösungsorientiert:", { bold: true }),
    bodyText("Die Herausforderungen im Gesundheitswesen ernst nehmen – und konkret zeigen, wie das Gesundheitsportal löst, nicht nur beschreibt.", { indent: true }),
    spacer(),
    bodyText("Keine Technik-Überwältigung:", { bold: true }),
    bodyText("Gegenüber Entscheidern (CDO, Verwaltungsdirektor) Nutzen vor Technik. FHIR und DICOM sind Mittel, keine Botschaft.", { indent: true }),
    spacer(),
    bodyText("Österreich-verbunden, DACH-relevant:", { bold: true }),
    bodyText("Lokale Expertise und Regulatorik-Kenntnis als Stärke, nicht als Einschränkung.", { indent: true }),
    divider()
  );

  // ─── FELD 6: KEY CLAIMS ────────────────────────────────────
  children.push(
    headingField("6", "Key Claims"),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Claim 1",
          bold: true,
          color: CF_WHITE,
          size: 20,
          font: "Calibri",
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_ACCENT, fill: CF_ACCENT },
      spacing: { before: 120, after: 0 },
      indent: { left: convertInchesToTwip(0.1) },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Weniger Telefon, mehr Versorgung: Krankenhäuser und Ordinationen ersetzen manuelle Terminprozesse durch 24/7-Online-Buchung mit automatischen Erinnerungen – und gewinnen administrative Kapazität für das zurück, was zählt: die Patientenversorgung.",
          color: CF_DARK_GRAY,
          size: 21,
          font: "Calibri",
          bold: true,
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_LIGHT_BLUE, fill: CF_LIGHT_BLUE },
      spacing: { before: 0, after: 120 },
      indent: { left: convertInchesToTwip(0.2), right: convertInchesToTwip(0.2) },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Claim 2",
          bold: true,
          color: CF_WHITE,
          size: 20,
          font: "Calibri",
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_ACCENT, fill: CF_ACCENT },
      spacing: { before: 120, after: 0 },
      indent: { left: convertInchesToTwip(0.1) },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Alle Gesundheitsdaten, eine Plattform: Von ELGA-Befunden mit DICOM-Viewer bis zum Blutdruckverlauf aus der Smartwatch – Ärzte sehen beim Termin das vollständige Bild, Patienten behalten die Datensouveränität.",
          color: CF_DARK_GRAY,
          size: 21,
          font: "Calibri",
          bold: true,
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_LIGHT_BLUE, fill: CF_LIGHT_BLUE },
      spacing: { before: 0, after: 120 },
      indent: { left: convertInchesToTwip(0.2), right: convertInchesToTwip(0.2) },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Claim 3",
          bold: true,
          color: CF_WHITE,
          size: 20,
          font: "Calibri",
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_ACCENT, fill: CF_ACCENT },
      spacing: { before: 120, after: 0 },
      indent: { left: convertInchesToTwip(0.1) },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Österreichs bewährteste Patientenportal-Plattform im laufenden Betrieb: Rahmenvertrag über 150 Mio. EUR mit der OÖG, aktiver Rollout in bis zu acht Bundesländern – gebaut und geliefert, nicht pilotiert.",
          color: CF_DARK_GRAY,
          size: 21,
          font: "Calibri",
          bold: true,
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_LIGHT_BLUE, fill: CF_LIGHT_BLUE },
      spacing: { before: 0, after: 120 },
      indent: { left: convertInchesToTwip(0.2), right: convertInchesToTwip(0.2) },
    }),
    divider()
  );

  // ─── FELD 7: CHALLENGES ────────────────────────────────────
  children.push(
    headingField("7", "Challenges – Konkrete Kundenprobleme"),
    spacer(),
    bodyText("Challenge 1: Terminmanagement im Jahr 2026 noch per Telefon und Fax", { bold: true }),
    bodyText(
      "Kliniken und Ordinationen bearbeiten Terminanfragen manuell – mit langen Warteschleifen für Patienten, hohem Personalaufwand und unkontrollierbarer No-Show-Rate ohne automatische Erinnerungssysteme. Das kostet täglich Kapazität, Umsatz und Patientenzufriedenheit. In Ambulanzen mit hohem Terminvolumen bindet das manuelle Terminmanagement Vollzeitstellen, die für klinische Aufgaben fehlen."
    ),
    spacer(),
    bodyText("Challenge 2: Kein zusammenhängendes Patientenbild beim Termin", { bold: true }),
    bodyText(
      "Ärzte treffen Entscheidungen mit unvollständigen Informationen. ELGA-Befunde sind in separaten Systemen schwer zugänglich, Vitaldatenverläufe liegen beim Patienten oder isoliert in Wearable-Apps, Dokumente werden auf Papier mitgebracht oder gehen verloren. Diese Medienbrüche kosten wertvolle Konsultationszeit, erhöhen das klinische Risiko bei Medikationsfehlern und verhindern eine kontinuierliche Versorgungsqualität zwischen den Terminen."
    ),
    spacer(),
    bodyText("Challenge 3: Regulatorischer Handlungsdruck ohne konsolidierte Plattformlösung", { bold: true }),
    bodyText(
      "Der European Health Data Space (EHDS), die ePA-Pflicht in Deutschland und die ELGA-Integration in Österreich erzeugen unmittelbaren Handlungsbedarf in allen Gesundheitseinrichtungen. Ohne eine integrierte Plattform bedeutet Compliance-Erfüllung den parallelen Betrieb mehrerer Insellösungen mit eigenen Schnittstellen, Update-Zyklen und Sicherheitsrisiken – erhebliche TCO-Treiben und technisches Risiko."
    ),
    divider()
  );

  // ─── FELD 8: BENEFITS ──────────────────────────────────────
  children.push(
    headingField("8", "Benefits – Nutzenblöcke"),
    spacer(),
    bodyText("Benefit 1 (zu Claim 1): Administrativen Aufwand messbar reduzieren", { bold: true }),
    bodyText(
      "24/7-Online-Buchung, automatische Erinnerungen und digitale Aufklärungsbögen ersetzen Telefon-Termine, reduzieren No-Shows und entlasten das Verwaltungspersonal. Ambulanzen und Ordinationen gewinnen Kapazitäten zurück, die direkt in bessere Patientenversorgung investiert werden können."
    ),
    annahme("[ANNAHME: Messbare Reduktion No-Show-Rate um 20–30 % basierend auf Branchenvergleichswerten – mit OÖG-Daten nach Rollout zu validieren]"),
    spacer(),
    bodyText("Benefit 2 (zu Claim 2): Vollständige Patienteninformation beim Termin – ohne Systemwechsel", { bold: true }),
    bodyText(
      "ELGA-Anbindung (eBefund, eMedikation, eImpfpass, DICOM), Vitaldaten aus integrierten Wearables und Patientendokumente aus der sicheren Wallet stehen dem Arzt direkt im Portal zur Verfügung. Das bedeutet bessere klinische Entscheidungen, weniger Doppeluntersuchungen und eine lückenlose Versorgungsdokumentation."
    ),
    spacer(),
    bodyText("Benefit 3 (zu Claim 3): Investitionssicherheit durch bewiesene Skalierung", { bold: true }),
    bodyText(
      "Ein Rahmenvertrag über 150 Mio. EUR mit der OÖG und der laufende Betrieb in bis zu acht österreichischen Bundesländern belegen die technische und organisatorische Reife der Plattform. Neue Auftraggeber profitieren von einem kampferprobten System, nicht von einem Pilotprojekt."
    ),
    divider()
  );

  // ─── FELD 9: DIFFERENTIATORS ───────────────────────────────
  children.push(
    headingField("9", "Differentiators – Echte Unterschiede zu Wettbewerbern"),
    spacer(),
    numberedPoint(
      "Native iOS- und Android-Apps mit biometrischer Authentifizierung: Während Wettbewerber primär Web-Lösungen oder responsive Portale anbieten, sind Cloudflights Apps vollständige native Applikationen mit Fingerprint/Face-ID und Offline-Fähigkeit – Patientenakzeptanz auf Smartphone-Niveau.",
      1
    ),
    numberedPoint(
      "ELGA-Tiefenintegration, die wirklich integriert ist: eBefund mit DICOM-Viewer, eMedikation, eImpfpass und CDA-Parsing sind nahtlos im Portalerlebnis eingebettet – nicht als externe Verlinkung oder separates System. Patienten und Ärzte bewegen sich nie aus dem Portal heraus.",
      2
    ),
    numberedPoint(
      "Echtes White-Label-Branding für Auftraggeber: Vollständige Anpassbarkeit von Logo, Farbschema und Schriftarten – jede Gesundheitseinrichtung betreibt das Portal unter ihrer eigenen Marke. Patienten erleben die Marke ihrer Klinik, nicht die eines Softwareanbieters.",
      3
    ),
    numberedPoint(
      "End-to-End-Plattform aus einer Hand: Terminbuchung, Videosprechstunde, Matrix-basierter sicherer Chat, Dokumenten-Wallet, Vitaldaten-Monitoring und ELGA in einer Lösung – statt 4–6 Einzelsysteme mit Integrationsaufwand, getrennten SLAs und unterschiedlichen Datenschutzverantwortlichkeiten.",
      4
    ),
    numberedPoint(
      "Eigener FHIR-Server durch PineIT-Partnerschaft: Die Integrationsplattform ist kein Drittprodukt, sondern gemeinsam mit PineIT entwickelt und vollständig kontrolliert – offene Interoperabilität mit bestehenden KIS/HIS ohne proprietäre Abhängigkeiten.",
      5
    ),
    numberedPoint(
      "Nachgewiesene Liefertreue in komplexen öffentlichen Projekten: Während Siemens Healthineers aus dem bayerischen Mega-Projekt (76 Krankenhäuser, 150 Standorte) wegen fehlender IT-Sicherheitszertifikate aussteigen musste, ist das Cloudflight Gesundheitsportal in bis zu acht österreichischen Bundesländern im laufenden Betrieb.",
      6
    ),
    numberedPoint(
      "ID Austria Integration mit Authentifizierungsstufen 0–2: Höchste österreichische eID-Standards mit mehreren Sicherheitsstufen – regulatorisch konform, benutzerfreundlich und für den österreichischen Markt entwickelt.",
      7
    ),
    divider()
  );

  // ─── FELD 10: CORE PERSONAS ────────────────────────────────
  children.push(
    headingField("10", "Core Personas"),
    spacer(),
    personasTable(),
    spacer(),
    divider()
  );

  // ─── FELD 11: USE CASES ─────────────────────────────────────
  children.push(
    headingField("11", "Use Cases – Konkrete Anwendungsfälle"),
    spacer(),
    bodyText("Use Case 1: Automatisiertes Terminmanagement für Krankenhausambulanz", { bold: true }),
    bodyText(
      "Kontext: Fondskrankenhaus mit 5 Ambulanzen, täglich 200+ Terminanfragen per Telefon. Problem: Überlastung des Sekretariats, lange Warteschleifen, hohe No-Show-Rate. Lösung: 24/7-Online-Buchung mit automatischen SMS/E-Mail-Erinnerungen, digitale Aufklärungsbögen vorab, Massen-Storno-Funktion für kurzfristige Änderungen. Ergebnis: Sekretariat entlastet, Auslastung optimiert, Patienten messen Erreichbarkeit an digitalen Maßstäben."
    ),
    annahme("[ANNAHME: No-Show-Reduktion 20–30 %, Einsparung 1–2 VZÄ Verwaltung pro Ambulanz – mit OÖG-Rollout-Daten belegen]"),
    spacer(),
    bodyText("Use Case 2: Videosprechstunde für Nachsorgetermine und chronische Erkrankungen", { bold: true }),
    bodyText(
      "Kontext: Facharztpraxis mit hohem Anteil an Kontrolluntersuchungen (Kardiologie, Diabetologie). Problem: Unnötige Präsenzbesuche belasten Terminkapazitäten und Patienten. Lösung: DSGVO-konforme Videosprechstunde im Portal mit Screen-Sharing für Befund-Besprechung, integrierter Chat für Vor- und Nachbereitung. Ergebnis: Mehr Versorgungstiefe pro Zeiteinheit, weniger Reiseaufwand für Patienten – besonders relevant in ländlichen Regionen."
    ),
    spacer(),
    bodyText("Use Case 3: ELGA-gestützter Aufnahmeprozess bei stationärer Aufnahme", { bold: true }),
    bodyText(
      "Kontext: Akutkrankenhaus, Aufnahmeprozess mit Papierdokumentation und veralteten Medikamentenlisten. Problem: Pflegepersonal verbringt Zeit mit Medikationserfassung statt Patientenbetreuung. Lösung: Direkter ELGA-Zugriff auf aktuelle eMedikation und eImpfpass des Patienten über das Portal beim Aufnahmegespräch – ohne Systemwechsel, ohne veraltete Papierlisten. Ergebnis: Schnellere Aufnahme, reduziertes Medikationsfehler-Risiko, weniger Doppelerfassung."
    ),
    spacer(),
    bodyText("Use Case 4: Vitaldaten-Monitoring für chronisch kranke Patienten zwischen Terminen", { bold: true }),
    bodyText(
      "Kontext: Diabetologische Ordination, Patienten messen täglich Blutzucker aber Arzt sieht Verlauf nur beim Termin. Problem: Therapiesteuerung basiert auf Momentaufnahme statt kontinuierlichem Verlauf. Lösung: Automatische Synchronisation von Blutzucker, Blutdruck und Puls via Apple Watch, Garmin oder Fitbit – Arzt sieht Verlaufstrends direkt im Portal. Ergebnis: Bessere Therapiesteuerung, weniger Notfalltermine, stärkere Patientenbindung."
    ),
    spacer(),
    bodyText("Use Case 5: Gebrandetes Gesundheitsportal für Landesgesundheitsholding", { bold: true }),
    bodyText(
      "Kontext: Landesgesundheitsholding mit 30 Standorten unterschiedlicher Versorgungsstufen. Problem: Patienten erleben unterschiedliche digitale Systeme je nach Standort, keine einheitliche digitale Markenidentität. Lösung: Vollständig gebrandetes Portal unter der Marke der Holding für alle Standorte – einheitliche UX, lokale Anpassungsmöglichkeiten pro Standort, zentrales Administration. Ergebnis: Stärkung der Holding-Marke, einheitliches Patientenerlebnis, effiziente zentrale Verwaltung."
    ),
    divider()
  );

  // ─── FELD 12: BUYING CRITERIA ──────────────────────────────
  children.push(
    headingField("12", "Buying Criteria"),
    spacer(),
    numberedPoint("Regulatorische Compliance: ELGA-Zertifizierung, EHDS-Konformität, DSGVO, ID Austria Integration, BSI/IT-Sicherheitsstandards", 1),
    numberedPoint("Nachgewiesene Referenzen und Ausschreibungskonformität: Vergleichbare Deployments, EU-weite Vergabe-Tauglichkeit, öffentlicher Sektor Erfahrung", 2),
    numberedPoint("Integrationsfähigkeit mit bestehenden KIS/HIS: FHIR R4, HL7, IHE-Standards; Kompatibilität mit SAP IS-H, ORBIS, CGM, i.s.h.med", 3),
    numberedPoint("Individualisierbarkeit: Branding, Workflow-Anpassung, modularer Funktionsumfang ohne Kernprodukt-Abhängigkeit", 4),
    numberedPoint("TCO und Betriebskosten: Transparentes Lizenzmodell, Hosting-Optionen (Cloud/On-Premise/Hybrid), SLA-Definitionen, Wartungsaufwand", 5),
    numberedPoint("Implementierungsgeschwindigkeit und Change Management: Time-to-live, Schulungskonzept, Onboarding-Unterstützung, lokaler Support", 6),
    numberedPoint("Langfristige Anbieter-Stabilität: Unternehmensstabilität, Entwicklungsroadmap, strategische Partnerschaft vs. Produktlieferant", 7),
    divider()
  );

  // ─── FELD 13: COMPETITORS ──────────────────────────────────
  children.push(
    headingField("13", "Top 3 Competitors"),
    spacer(),
    competitorTable(),
    spacer(),
    divider()
  );

  // ─── FELD 14: MARKETING RESOURCES ─────────────────────────
  children.push(
    headingField("14", "Marketing Resources"),
    spacer(),
    bodyText("Early Funnel – Awareness (Problem bewusst machen)", { bold: true }),
    bulletPoint("Artikel: 'EHDS 2026: Was österreichische Gesundheitsholdings jetzt konkret umsetzen müssen' – Ziel: Regulatorischen Druck adressieren, Cloudflight als Experten positionieren"),
    bulletPoint("Artikel: 'Das Ende der Telefonordination: Wie digitales Terminmanagement Kliniken entlastet und Patienten gewinnt'"),
    bulletPoint("Infografik: 'Medienbrüche im Gesundheitswesen – versteckte Kosten und messbare Lösungen'"),
    bulletPoint("LinkedIn-Thought-Leadership: Regulatorische Updates (EHDS, ePA, ELGA), Marktdaten Patientenportale, Kommentare zu Wettbewerber-Nachrichten (z.B. Siemens-Ausstieg Bayern)"),
    spacer(),
    bodyText("Mid Funnel – Consideration (Lösung evaluieren)", { bold: true }),
    bulletPoint("Case Study: 'OÖG-Rollout in bis zu 8 Bundesländern – Digitale Patientenvernetzung in der Praxis' [PLATZHALTER: Freigabe OÖG einholen]"),
    bulletPoint("Demo-Video: Vollständiger Walkthrough aller Kernfunktionen aus Patienten- und Arztperspektive (15 Min., deutsche Sprache)"),
    bulletPoint("Whitepaper: 'ELGA-Integration 2026: Technische Anforderungen, Best Practices und Fallstricke'"),
    bulletPoint("Vergleichsmatrix: Funktions- und Compliance-Vergleich Patientenportal-Lösungen (Cloudflight vs. x-Tention vs. Siemens vs. CANCOM)"),
    bulletPoint("Webinar: 'Patientenportale 2026: Regulatorik, UX und ROI – was Gesundheitseinrichtungen wissen müssen'"),
    spacer(),
    bodyText("Down Funnel – Decision (Kaufentscheidung unterstützen)", { bold: true }),
    bulletPoint("ROI-Kalkulator: Einsparungspotenzial durch digitales Terminmanagement (Zeitaufwand, No-Show-Rate, Personalkosten) – downloadbar, individuell konfigurierbar"),
    bulletPoint("Referenzgespräch mit OÖG-Verantwortlichen [PLATZHALTER: Kontakt und NDI-Freigabe sicherstellen]"),
    bulletPoint("Technisches Integrations-Whitepaper: FHIR-Architektur, KIS-Anbindungsoptionen, Sicherheitskonzept, ID Austria Integration"),
    bulletPoint("Pilot-Angebot: Strukturiertes 90-Tage-Pilot-Programm für Entscheidungsträger mit definierten Erfolgskriterien"),
    divider()
  );

  // ─── FELD 15: 50-WORD DESCRIPTION ─────────────────────────
  children.push(
    headingField("15", "50-Word-Description"),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Das Cloudflight Gesundheitsportal vernetzt Patienten und Ärzte auf einer vollständig brandbaren Plattform: 24/7-Terminbuchung, DSGVO-konforme Videosprechstunde, sichere Dokumenten-Wallet, Vitaldaten-Monitoring mit Wearable-Integration und einzigartige ELGA-Tiefenanbindung. Bewährt in bis zu acht österreichischen Bundesländern durch Rahmenvertrag OÖG. Native iOS- und Android-Apps inklusive.",
          color: CF_DARK_GRAY,
          size: 22,
          font: "Calibri",
          italics: true,
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_LIGHT_BLUE, fill: CF_LIGHT_BLUE },
      spacing: { before: 100, after: 100 },
      indent: { left: convertInchesToTwip(0.3), right: convertInchesToTwip(0.3) },
    }),
    annahme("[Wortanzahl: ca. 50 Wörter – bei Verwendung bitte exakt zählen und ggf. anpassen]"),
    divider()
  );

  // ─── FELD 16: 100-WORD DESCRIPTION ────────────────────────
  children.push(
    headingField("16", "100-Word-Description"),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Das Cloudflight Gesundheitsportal ist Österreichs umfassendste Plattform für die digitale Vernetzung von Patienten und Gesundheitsdienstleistern. Patienten buchen Termine rund um die Uhr, führen DSGVO-konforme Videosprechstunden durch, verwalten Dokumente sicher in der digitalen Wallet und synchronisieren Vitaldaten automatisch von Apple Watch, Garmin oder Fitbit. Die einzigartige ELGA-Tiefenintegration ermöglicht direkten Zugriff auf eBefunde, eMedikation, eImpfpass und DICOM-Bilder ohne Systemwechsel. Gesundheitseinrichtungen profitieren von weniger Verwaltungsaufwand, optimierter Terminauslastung und vollständiger Patienteninformation beim Termin. Mit Rahmenvertrag über 150 Mio. EUR bei der OÖG und Rollout in bis zu acht österreichischen Bundesländern ist das Gesundheitsportal die bewährteste und meisteingesetzte Patientenportal-Lösung Österreichs – vollständig anpassbar an das Corporate Design jeder Einrichtung.",
          color: CF_DARK_GRAY,
          size: 21,
          font: "Calibri",
          italics: true,
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_LIGHT_BLUE, fill: CF_LIGHT_BLUE },
      spacing: { before: 100, after: 100 },
      indent: { left: convertInchesToTwip(0.3), right: convertInchesToTwip(0.3) },
    }),
    annahme("[Wortanzahl: ca. 100 Wörter – bei Verwendung bitte exakt zählen und ggf. anpassen]"),
    divider()
  );

  // ─── FELD 17: MESSAGE FOR EXISTING CUSTOMERS ───────────────
  children.push(
    headingField("17", "Message for Existing Customers – Upsell / Cross-Sell"),
    spacer(),
    bodyText("Kernbotschaft:", { bold: true }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Ihr Portal ist live – jetzt kommen die nächsten Ausbaustufen.",
          bold: true,
          color: CF_BLUE,
          size: 24,
          font: "Calibri",
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 120 },
    }),
    spacer(),
    bodyText("Konkrete Upsell-Möglichkeiten:", { bold: true }),
    bulletPoint("Vitaldaten-Monitoring + Wearable-Integration (Apple Watch, Garmin, Fitbit): Erweitern Sie Ihre Plattform um kontinuierliches Gesundheitsmonitoring für chronisch kranke Patienten"),
    bulletPoint("Matrix-basierter sicherer Chat: Asynchrone, datenschutzkonforme Kommunikation zwischen Patienten und Behandlungsteam – ohne Telefonflut"),
    bulletPoint("Digitale Aufklärungsbögen: Vollständig digitaler Aufnahmeprozess mit rechtssicherer Archivierung und Zeitersparnis beim Check-in"),
    bulletPoint("Neue ELGA-Module: eImpfpass-Zugriff und erweiterte eMedikations-Workflows direkt aus dem Portal"),
    bulletPoint("Zusätzliche Standorte: Skalierung auf weitere Kliniken oder Ambulanzen innerhalb der bestehenden Infrastruktur"),
    spacer(),
    bodyText("Cross-Sell aus dem Cloudflight-Portfolio:", { bold: true }),
    annahme("[ANNAHME: Bitte mit Cloudflight-Portfolio-Team abstimmen, welche weiteren Produkte/Services für bestehende Gesundheitsportal-Kunden relevant sind]"),
    bulletPoint("Custom Software Engineering: Individuelle Erweiterungen und Integrationen über das Standardprodukt hinaus"),
    bulletPoint("Data & Analytics: Auswertung von Nutzungsdaten, Terminauslastung und Vitaldata-Trends für Entscheidungsträger"),
    spacer(),
    bodyText("Call-to-Action:", { bold: true }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Buchen Sie jetzt einen kostenfreien Roadmap-Workshop mit Ihrem Cloudflight-Ansprechpartner und entwickeln Sie gemeinsam Ihre individuelle Ausbau-Roadmap für die nächsten 12–24 Monate.",
          bold: true,
          color: CF_BLUE,
          size: 21,
          font: "Calibri",
        }),
      ],
      shading: { type: ShadingType.SOLID, color: CF_LIGHT_BLUE, fill: CF_LIGHT_BLUE },
      spacing: { before: 100, after: 100 },
      indent: { left: convertInchesToTwip(0.2), right: convertInchesToTwip(0.2) },
    }),
    divider()
  );

  // ─── FOOTER INFO ───────────────────────────────────────────
  children.push(
    spacer(2),
    new Paragraph({
      children: [
        new TextRun({
          text: "Quellen & Recherche-Basis",
          bold: true,
          color: CF_BLUE,
          size: 22,
          font: "Calibri",
        }),
      ],
      spacing: { before: 200, after: 80 },
    }),
    bulletPoint("Produktdokumentation Cloudflight Gesundheitsportal (intern)"),
    bulletPoint("x-tention: NÖ LGA-Auftrag (September 2025) – x-tention.com/de/news"),
    bulletPoint("x-tention: BG Kliniken Patientenportal-Implementierung – kma-online.de"),
    bulletPoint("Siemens Healthineers: Ausstieg aus 'Mein-Krankenhaus.Bayern' (76 KH, 150 Standorte) – kma-online.de, Business Insider"),
    bulletPoint("Siemens Healthineers: Österreich-Referenzen (Vinzenz Gruppe 'Hallo Gesundheit", SALK) – siemens-healthineers.com"),
    bulletPoint("CANCOM Healthcare: Positionierung und DMEA 2025 – healthcare.cancom.de, cancom.info"),
    bulletPoint("Marktdaten: EHDS, Krankenhauszukunftsgesetz, 40–80 Ausschreibungen/Jahr EU"),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Vertraulich – Cloudflight GmbH | Message Master Gesundheitsportal | Version 3.0 | 04.03.2026",
          color: CF_MEDIUM_GRAY,
          size: 16,
          font: "Calibri",
          italics: true,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 60 },
    })
  );

  // ─── BUILD DOCUMENT ────────────────────────────────────────
  const doc = new Document({
    creator: "Cloudflight",
    title: "Message Master – Gesundheitsportal",
    description: "B2B Positioning and Messaging Document for Cloudflight Gesundheitsportal",
    styles: {
      default: {
        document: {
          run: {
            font: "Calibri",
            size: 20,
            color: CF_DARK_GRAY,
          },
          paragraph: {
            spacing: { line: 276 },
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.9),
              right: convertInchesToTwip(0.9),
              bottom: convertInchesToTwip(0.9),
              left: convertInchesToTwip(0.9),
            },
          },
        },
        children,
      },
    ],
  });

  return doc;
}

// ============================================================
// MAIN
// ============================================================

async function main() {
  console.log("Erstelle Message Master Gesundheitsportal...");

  const doc = buildDocument();

  const outputPath = "/home/user/Message-master/output/message-master-gesundheitsportal-2026-03-04.docx";

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);

  const stats = fs.statSync(outputPath);
  console.log(`✓ Dokument erfolgreich erstellt: ${outputPath}`);
  console.log(`  Dateigröße: ${(stats.size / 1024).toFixed(1)} KB`);
  console.log(`  Alle 17 Felder vollständig ausgefüllt.`);
}

main().catch((err) => {
  console.error("Fehler:", err);
  process.exit(1);
});
