const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require('docx');

// ===== Style helpers =====
const RED  = "C0392B";
const NAVY = "112337";
const GREY = "6B7280";
const LINE = "D8DBE2";
const SOFT = "F5F6F8";

const border = { style: BorderStyle.SINGLE, size: 6, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };

const blank = (h = 100) => new Paragraph({ spacing: { after: h }, children: [new TextRun("")] });

const h1 = (txt) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 200 },
  children: [new TextRun({ text: txt, bold: true, color: NAVY, size: 36, font: "Arial" })],
});

const h2 = (txt) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 140 },
  children: [new TextRun({ text: txt, bold: true, color: NAVY, size: 28, font: "Arial" })],
});

const h3 = (txt) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text: txt, bold: true, color: RED, size: 22, font: "Arial" })],
});

const eyebrow = (txt) => new Paragraph({
  spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: txt.toUpperCase(), bold: true, color: RED, size: 18, font: "Arial", characterSpacing: 80 })],
});

const para = (txt, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 120 },
  children: [new TextRun({ text: txt, color: opts.color ?? "333333", size: opts.size ?? 22, font: "Arial", bold: opts.bold, italics: opts.italics })],
});

const bullet = (txt) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80 },
  children: [new TextRun({ text: txt, color: "333333", size: 22, font: "Arial" })],
});

// Answer-box: a single-cell table with a light border and a few empty lines inside
const answerBox = (lines = 3, label) => {
  const inner = [];
  if (label) {
    inner.push(new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: label, italics: true, color: GREY, size: 18, font: "Arial" })],
    }));
  }
  for (let i = 0; i < lines; i++) {
    inner.push(new Paragraph({
      spacing: { after: 60 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 6 } },
      children: [new TextRun({ text: " ", font: "Arial", size: 22 })],
    }));
  }
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 6, color: LINE },
    },
    rows: [
      new TableRow({
        children: [new TableCell({
          width: { size: 9360, type: WidthType.DXA },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
          children: inner,
        })],
      }),
    ],
  });
};

// Checkbox option line
const checkbox = (txt) =>
  new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: "☐  ", font: "Arial", size: 24, color: NAVY }),
      new TextRun({ text: txt, font: "Arial", size: 22, color: "333333" }),
    ],
  });

// Field row: "Label: ____________________"
const field = (label, hint) =>
  new Paragraph({
    spacing: { after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
    children: [
      new TextRun({ text: `${label}: `, bold: true, font: "Arial", size: 22, color: NAVY }),
      new TextRun({ text: hint ? ` (${hint})` : "", italics: true, font: "Arial", size: 18, color: GREY }),
    ],
  });

// Colored section divider
const divider = () => new Paragraph({
  spacing: { before: 240, after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: RED, space: 1 } },
  children: [new TextRun("")],
});

// Section header banner (shaded cell with title)
const banner = (label, subtitle) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  borders: {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  },
  rows: [
    new TableRow({
      children: [new TableCell({
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 240, bottom: 240, left: 280, right: 280 },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            children: [new TextRun({ text: label.toUpperCase(), bold: true, color: "FFFFFF", size: 30, font: "Arial", characterSpacing: 60 })],
          }),
          subtitle ? new Paragraph({
            spacing: { before: 80 },
            children: [new TextRun({ text: subtitle, color: "BFC9DA", size: 20, font: "Arial", italics: true })],
          }) : null,
        ].filter(Boolean),
      })],
    }),
  ],
});

// ===== Build content =====
const content = [];

// ---------- COVER ----------
content.push(
  blank(200),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "FULL ATTACK MOTORSPORTS", bold: true, color: NAVY, size: 56, font: "Arial", characterSpacing: 80 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: "WEBSITE CONTENT & FEEDBACK QUESTIONNAIRE", bold: true, color: RED, size: 24, font: "Arial", characterSpacing: 100 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
    children: [new TextRun({ text: "Prepared for Lane Alexander — FAM Detailing · Chico, CA", italics: true, color: GREY, size: 22, font: "Arial" })],
  }),
);

// How to use box
content.push(
  banner("How to Use This Document", "Fill in what you want — leave anything blank you're unsure about"),
  blank(100),
  para("This questionnaire walks through every section of your new Full Attack Motorsports website. For each section, you'll find:"),
  bullet("A short summary of what's currently on that section of the site"),
  bullet("Checkboxes for quick approvals or change requests"),
  bullet("Open boxes where you can type or write longer feedback"),
  bullet("Image / asset requests so we know what photos to swap in"),
  blank(100),
  para("Tip: You can fill this out in Word, Google Docs, or print it and write on it — whatever's easiest. Skip any question that doesn't apply.", { italics: true, color: GREY, size: 20 }),
  blank(200),
);

// Client info
content.push(
  h3("Your Information"),
  field("Name", "person filling this out"),
  field("Role", "owner / manager / etc."),
  field("Best contact method", "phone / text / email"),
  field("Date completed"),
);

content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- BRAND IDENTITY ----------
content.push(
  banner("01 · Brand Identity & Logo", "Foundational visual choices that apply to every section"),
  blank(100),
  para("The site currently uses a stylized rotary mark + the wordmark FULL ATTACK MOTORSPORTS in a condensed industrial font (Oswald), paired with race-red (#E10600) on a clean white/light-gray background with navy headings."),
);

content.push(h3("Logo"));
content.push(checkbox("Use the official FAM Detailing logo (rotary shield) — please send high-res PNG/SVG"));
content.push(checkbox("Use the placeholder rotary SVG that's there now"));
content.push(checkbox("Want a new logo designed — describe direction below"));
content.push(answerBox(3, "Logo notes / direction"));

content.push(h3("Brand Colors"));
content.push(para("Currently using:  Race Red #E10600  ·  Navy #112337  ·  White / light gray backgrounds", { italics: true, color: GREY, size: 20 }));
content.push(checkbox("Keep current colors"));
content.push(checkbox("Adjust the red (specify shade below)"));
content.push(checkbox("Add a secondary accent color"));
content.push(answerBox(2, "Color preferences / hex values"));

content.push(h3("Tagline / Voice"));
content.push(para("Current tagline pulled from your site: \"Attack your dreams full on with Full Attack Motorsports.\""));
content.push(checkbox("Keep current tagline"));
content.push(checkbox("Want a different tagline (write below)"));
content.push(answerBox(2, "New tagline or voice notes"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- HERO ----------
content.push(
  banner("02 · Hero Section (Top of Page)", "The first thing visitors see — biggest impact"),
  blank(100),
  para("Currently shows: \"Chico's Premier Auto Detailing & Ceramic Coating Shop\" with a subheading mentioning Lane Alexander, Butte College ASE training, two CTA buttons (Get a Quote / View Services), and three trust badges (5-Star, System X, Climate-Controlled)."),
);

content.push(h3("Headline"));
content.push(checkbox("Approved as-is"));
content.push(checkbox("Tweak — write new version below"));
content.push(answerBox(2, "New headline"));

content.push(h3("Subheading / Lead Paragraph"));
content.push(checkbox("Approved as-is"));
content.push(checkbox("Tweak — write new version below"));
content.push(answerBox(4, "New subheading"));

content.push(h3("Hero Image"));
content.push(para("Currently using the Toyota Supra detail photo from your site.", { italics: true, color: GREY, size: 20 }));
content.push(checkbox("Keep the Supra"));
content.push(checkbox("Use a different vehicle — describe / send photo"));
content.push(checkbox("Use a shot of the shop itself"));
content.push(answerBox(2, "Hero image preference"));

content.push(h3("Call-to-Action Buttons"));
content.push(para("Currently: \"Get a Quote\" and \"View Services\".", { italics: true, color: GREY, size: 20 }));
content.push(field("Primary button label", "default: Get a Quote"));
content.push(field("Primary button links to", "default: contact form"));
content.push(field("Secondary button label", "default: View Services"));
content.push(field("Secondary button links to", "default: services section"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- SERVICES ----------
content.push(
  banner("03 · Services Section", "The 6 service cards in the grid"),
  blank(100),
  para("Currently showing 6 services in this order: Detail Services · System X Ceramic Coating · New Car Protection · Paint Correction · Aftermarket Upgrades · Boats & RV Detailing."),
);

const services = [
  { name: "Detail Services", tag: "Most Popular", img: "Acura full detail photo" },
  { name: "System X Ceramic Coating", tag: "System X Certified", img: "Supra ceramic photo" },
  { name: "New Car Protection", tag: "(no tag)", img: "GMC truck photo" },
  { name: "Paint Correction", tag: "(no tag)", img: "Acura paint correction photo" },
  { name: "Aftermarket Upgrades", tag: "(no tag)", img: "Dodge Ram photo" },
  { name: "Boats & RV Detailing", tag: "Boats & RVs", img: "Boat detail photo" },
];

services.forEach((s, i) => {
  content.push(h3(`Service ${i + 1}: ${s.name}`));
  content.push(para(`Current tag: ${s.tag}  ·  Current image: ${s.img}`, { italics: true, color: GREY, size: 20 }));
  content.push(checkbox("Keep this service in the lineup"));
  content.push(checkbox("Rename this service (write new name below)"));
  content.push(checkbox("Remove this service entirely"));
  content.push(checkbox("Replace the image (describe / attach new photo)"));
  content.push(checkbox("Rewrite the description (write new copy below)"));
  content.push(answerBox(3, `Changes for ${s.name}`));
});

content.push(h3("Additional Services to Add"));
content.push(para("Are there services we missed? (e.g., Window Tinting, PPF, Headlight Restoration, Engine Bay Detail, etc.)"));
content.push(answerBox(5, "New services to add — name + 1-sentence description each"));

content.push(h3("Pricing"));
content.push(para("Currently: no specific prices shown (matching your existing site's approach). Disclaimer about \"estimates for average vehicle\" is included."));
content.push(checkbox("Keep prices hidden — quote on request"));
content.push(checkbox("Show starting-at prices on cards (e.g., \"From $XXX\")"));
content.push(checkbox("Add a full pricing/packages page"));
content.push(answerBox(4, "Pricing notes — what should be public vs. quote-only?"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- WHY US ----------
content.push(
  banner("04 · Why Choose Us", "Four reasons drivers should pick Full Attack"),
  blank(100),
  para("Currently lists: System X Certified Installer · ASE/CORE Trained Founder · Climate-Controlled Shop · Honest, Up-Front Pricing — with two race-track photos of Lane on the side."),
);

content.push(h3("The Four Reasons"));
content.push(checkbox("All four reasons are accurate and approved"));
content.push(checkbox("Need to edit one or more — describe below"));
content.push(checkbox("Want to add a 5th reason (gets reformatted to fit)"));
content.push(answerBox(5, "Why-us edits / additions"));

content.push(h3("Side Photos"));
content.push(para("Currently uses 2 race-track photos of Lane.", { italics: true, color: GREY, size: 20 }));
content.push(checkbox("Keep both race photos"));
content.push(checkbox("Swap to before/after detail shots"));
content.push(checkbox("Swap to shop interior photos"));
content.push(answerBox(2, "Photo direction"));

content.push(divider());

// ---------- PROCESS ----------
content.push(
  banner("05 · Our Process (4 Steps)", "How Full Attack treats a vehicle, start to finish"),
  blank(100),
  para("Currently: 1) Consultation  2) Decontamination  3) Correction  4) Protection."),
);

content.push(h3("Step Accuracy"));
content.push(checkbox("These four steps accurately describe what we do"));
content.push(checkbox("Want to rename one or more steps"));
content.push(checkbox("The actual process has more/fewer steps — list below"));
content.push(answerBox(6, "Your actual process, step-by-step"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- STATS BAND ----------
content.push(
  banner("06 · Stats Band (Dark Navy)", "Quick credibility numbers under Process"),
  blank(100),
  para("Currently shows: 5★ Verified Reviews · System X Certified Installer · Lifetime Coating on New Cars · Up to 8 yr Coating on Used Cars."),
);

content.push(checkbox("All four stats approved"));
content.push(checkbox("Want to add real review count (e.g., \"200+ Reviews\") — provide number"));
content.push(checkbox("Want to add years-in-business (provide year founded)"));
content.push(checkbox("Want to add vehicles detailed count (provide number)"));
content.push(answerBox(3, "Real numbers we should plug in"));

content.push(divider());

// ---------- GALLERY ----------
content.push(
  banner("07 · Gallery Section", "Mosaic grid of recent work"),
  blank(100),
  para("Currently shows 7 tiles using photos pulled from your existing site (Supra, Acura ×2, Boat, Dodge Ram, GMC, Lane racing)."),
);

content.push(h3("Photos"));
content.push(checkbox("Keep current gallery photos"));
content.push(checkbox("Replace with newer photos (send to Lane to upload)"));
content.push(checkbox("Pull live from Instagram @fullattackmotorsports automatically (requires API setup)"));
content.push(answerBox(3, "Gallery preferences"));

content.push(h3("Before / After"));
content.push(checkbox("Add a dedicated Before/After slider section"));
content.push(checkbox("Not needed — gallery is enough"));
content.push(answerBox(2, "Notes"));

content.push(divider());

// ---------- REVIEWS ----------
content.push(
  banner("08 · Reviews Section", "6 customer testimonial cards"),
  blank(100),
  para("Currently shows 6 placeholder reviews with first names from nearby towns. THESE NEED TO BE REPLACED with real customer reviews before going live."),
  para("⚠️  Important: please paste real reviews below (Google, Facebook, Yelp, texts — all fine).", { bold: true, color: RED, size: 22 }),
);

for (let i = 1; i <= 6; i++) {
  content.push(h3(`Real Review #${i}`));
  content.push(field("Customer name", "first name + last initial is fine"));
  content.push(field("Town"));
  content.push(field("Star rating", "1–5"));
  content.push(answerBox(4, "Paste the review text"));
}

content.push(h3("Where Reviews Come From"));
content.push(checkbox("Embed Google Reviews automatically (we set this up)"));
content.push(checkbox("Manually managed (you send us new ones over time)"));
content.push(answerBox(2, "Notes"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- ABOUT ----------
content.push(
  banner("09 · About Section", "Lane's story + race chips + disclaimer"),
  blank(100),
  para("Currently includes Lane's bio (Butte College ASE/CORE Program), race venues (Thunder Hill, Sonoma, Buttonwillow, Laguna Seca), race series (Drift Apocalypse, Speed SF, SCCA), and the \"not a repair shop / no paint applied\" disclaimer."),
);

content.push(h3("Lane's Bio"));
content.push(checkbox("Approved as-is (pulled from your existing About page)"));
content.push(checkbox("Want to add more personal story / years of experience"));
content.push(checkbox("Want to add team members (other detailers on staff)"));
content.push(answerBox(6, "Bio updates / additions"));

content.push(h3("Race Venues & Series"));
content.push(checkbox("Current list is complete and accurate"));
content.push(checkbox("Add more venues / series (list below)"));
content.push(checkbox("Remove some venues / series (list below)"));
content.push(answerBox(3, "Race list updates"));

content.push(h3("About Photo"));
content.push(para("Currently uses Lane's race-car photo as the portrait.", { italics: true, color: GREY, size: 20 }));
content.push(checkbox("Keep the race-car photo"));
content.push(checkbox("Swap to a portrait of Lane (send photo)"));
content.push(checkbox("Swap to a shop / team photo"));
content.push(answerBox(2, "Photo direction"));

content.push(h3("Disclaimer Text"));
content.push(para("Current: \"This shop is not an auto repair shop, and we are not an auto paint shop. No paint is applied. We do detailing.\""));
content.push(checkbox("Keep exactly as-is"));
content.push(checkbox("Reword (write below)"));
content.push(answerBox(3, "Updated disclaimer text"));

content.push(divider());

// ---------- SERVICE AREA ----------
content.push(
  banner("10 · Service Area", "The location chips list"),
  blank(100),
  para("Currently: Chico, Paradise, Oroville, Durham, Magalia, Red Bluff, Orland, Willows, Butte/Tehama/Glenn Counties."),
);

content.push(checkbox("List is accurate"));
content.push(checkbox("Add more towns / regions (list below)"));
content.push(checkbox("Do mobile detailing? — should we add a mobile-service-area note?"));
content.push(answerBox(4, "Service area changes"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- FAQ ----------
content.push(
  banner("11 · FAQ Section", "Currently 6 questions in an accordion"),
  blank(100),
  para("Current FAQs cover: ceramic coating lifespan, paint correction before coating, no-repair/no-paint, boats & RVs, pricing, and how to book."),
);

content.push(h3("Existing FAQ Answers"));
content.push(checkbox("All 6 answers are accurate"));
content.push(checkbox("Need to edit one or more (list below)"));
content.push(answerBox(5, "FAQ corrections"));

content.push(h3("Add More FAQs"));
content.push(para("What questions do customers ask you all the time? Add as many as you want."));
for (let i = 1; i <= 4; i++) {
  content.push(field(`Q${i}`));
  content.push(answerBox(3, `A${i}`));
}

content.push(divider());

// ---------- CONTACT ----------
content.push(
  banner("12 · Contact Section", "Address, phone, hours, quote form"),
  blank(100),
);

content.push(h3("Verify Contact Info"));
content.push(field("Shop address", "currently: 1377 E. 9th St., Chico, CA 95928"));
content.push(field("Phone", "currently: (530) 624-7110"));
content.push(field("Email", "currently not displayed — add one?"));
content.push(field("Business hours", "currently: Mon–Fri 8 AM – 5 PM, Sat by appt"));
content.push(field("Instagram", "currently: @fullattackmotorsports"));
content.push(field("Facebook", "currently: lane.alexander.73 — should this be a business page?"));
content.push(field("Google Business listing URL", "for reviews / map"));
content.push(field("Yelp listing URL"));

content.push(h3("Quote Form Fields"));
content.push(para("Currently asks: Name, Phone, Email, Year/Make/Model, Service, Message."));
content.push(checkbox("Form fields look good"));
content.push(checkbox("Add a \"How did you hear about us?\" field"));
content.push(checkbox("Add a photo upload (so customers can show vehicle condition)"));
content.push(checkbox("Add preferred date/time field"));
content.push(answerBox(3, "Other form changes"));

content.push(h3("Where Should Form Submissions Go?"));
content.push(field("Email address to receive quote requests"));
content.push(checkbox("Also text Lane when a form comes in (set up SMS)"));
content.push(checkbox("Send to a Google Sheet / CRM (specify which)"));
content.push(answerBox(2, "Form-routing notes"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- FINAL CTA + FOOTER ----------
content.push(
  banner("13 · Final CTA & Footer", "Last call-to-action + bottom of page"),
  blank(100),
  para("Currently: \"Attack Your Dreams Full On\" headline over a darkened photo of Lane racing, with Get a Quote / Call buttons. Footer has quick links, services list, address, and social."),
);

content.push(h3("Final CTA Headline"));
content.push(checkbox("\"Attack Your Dreams Full On\" is perfect"));
content.push(checkbox("Write a different headline (below)"));
content.push(answerBox(2, "New CTA headline"));

content.push(h3("Background Image"));
content.push(checkbox("Keep race-track photo of Lane"));
content.push(checkbox("Use the shop exterior"));
content.push(checkbox("Use a hero detail shot"));
content.push(answerBox(2, "Background preference"));

content.push(h3("Footer Tagline"));
content.push(para("Currently: \"Attack your dreams full on with Full Attack Motorsports.\""));
content.push(checkbox("Approved"));
content.push(checkbox("Different (below)"));
content.push(answerBox(2, "Footer tagline"));

content.push(divider());

// ---------- NEW PAGES / SECTIONS ----------
content.push(
  banner("14 · New Pages & Sections to Add", "What's missing that should be there?"),
  blank(100),
);

content.push(h3("Common Add-On Pages — check what you want"));
content.push(checkbox("Dedicated Ceramic Coating page (with package tiers + details)"));
content.push(checkbox("Dedicated Paint Correction page"));
content.push(checkbox("Dedicated Detail Services page with packages"));
content.push(checkbox("Pricing page with full packages and tiers"));
content.push(checkbox("Before/After portfolio with project details"));
content.push(checkbox("Blog / Journal (was on your old site)"));
checkboxesGroup1 = [
  "Booking calendar (customers pick a date/time directly)",
  "Online store for products (was on your old site — c2/c3/c7/c8 shop pages)",
  "Gift cards page",
  "Maintenance / wash plans for ceramic coating customers",
  "Mobile detailing booking flow",
  "Race team / sponsorship page (Lane's racing connection)",
];
checkboxesGroup1.forEach(c => content.push(checkbox(c)));

content.push(h3("Anything Else?"));
content.push(answerBox(8, "New pages, sections, or features you want"));

content.push(divider());
content.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- TONE / VOICE ----------
content.push(
  banner("15 · Tone, Voice & Personality", "How should the site sound?"),
  blank(100),
);

content.push(h3("Pick the closest fit"));
content.push(checkbox("Professional & polished (like Alpha Auto)"));
content.push(checkbox("Gritty / car-enthusiast / racing-rooted (lean into Lane's race story)"));
content.push(checkbox("Friendly local shop / approachable (\"we're your neighbors\")"));
content.push(checkbox("Premium / boutique (System X is the lead — luxury focus)"));
content.push(checkbox("Mix of the above — describe below"));
content.push(answerBox(4, "Tone direction"));

content.push(h3("Words / Phrases You Love"));
content.push(answerBox(3, "Words, phrases, taglines you want sprinkled in"));

content.push(h3("Words / Phrases You Hate"));
content.push(answerBox(3, "Words, phrases, clichés to AVOID"));

content.push(divider());

// ---------- SEO ----------
content.push(
  banner("16 · SEO & Local Search", "How customers find you on Google"),
  blank(100),
  para("Site is already set up with Chico-focused SEO: city in title, geo meta tags, LocalBusiness schema, service-area chips, and keyword-rich copy."),
);

content.push(h3("Keywords You Want to Rank For"));
content.push(para("E.g., \"ceramic coating Chico\", \"auto detailing Paradise CA\". List as many as you can think of."));
content.push(answerBox(5, "Target keywords / phrases"));

content.push(h3("Competitors You Want to Beat in Search"));
content.push(answerBox(3, "Names + websites of competing shops"));

content.push(h3("Google Business / Reviews"));
content.push(field("Google Business Profile URL"));
content.push(field("Approximate # of Google reviews"));
content.push(field("Average star rating"));

content.push(divider());

// ---------- LEGAL / EXTRA ----------
content.push(
  banner("17 · Legal & Misc", "Optional housekeeping"),
  blank(100),
);

content.push(checkbox("Need a Privacy Policy page (we can generate one)"));
content.push(checkbox("Need a Terms of Service page"));
content.push(checkbox("Add an ADA / accessibility statement"));
content.push(checkbox("Add a tracking/analytics tool (e.g., Google Analytics 4)"));
content.push(field("Google Analytics ID", "if you already have one"));
content.push(field("Facebook Pixel ID", "for ads, optional"));
content.push(answerBox(4, "Other legal / privacy / tracking notes"));

content.push(divider());

// ---------- ANYTHING ELSE ----------
content.push(
  banner("18 · Open Floor", "Anything else on your mind"),
  blank(100),
  para("Big or small — features you wish you had on your old site, things you saw on competitor sites, dream features, brand story you want told. Doesn't have to be polished."),
);
content.push(answerBox(20, "Final thoughts"));

content.push(blank(200));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "— END OF QUESTIONNAIRE —", italics: true, color: GREY, size: 20, font: "Arial" })],
}));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 100 },
  children: [new TextRun({ text: "Thanks Lane! Send this back however's easiest and we'll build it in.", color: NAVY, size: 22, font: "Arial", bold: true })],
}));

// ===== Document =====
const doc = new Document({
  creator: "Full Attack Motorsports Website Team",
  title: "Full Attack Motorsports — Website Content & Feedback Questionnaire",
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: RED },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1440, bottom: 1080, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          spacing: { after: 0 },
          children: [
            new TextRun({ text: "FULL ATTACK MOTORSPORTS", bold: true, font: "Arial", size: 18, color: NAVY, characterSpacing: 80 }),
            new TextRun({ text: "   ·   Website Questionnaire", italics: true, font: "Arial", size: 18, color: GREY }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", font: "Arial", size: 18, color: GREY }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: GREY }),
            new TextRun({ text: " · Full Attack Motorsports · Chico, CA", font: "Arial", size: 18, color: GREY }),
          ],
        })],
      }),
    },
    children: content,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("FullAttack-Client-Questionnaire.docx", buffer);
  console.log("✓ Wrote FullAttack-Client-Questionnaire.docx (" + (buffer.length / 1024).toFixed(1) + " KB)");
});
