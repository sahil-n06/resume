const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, LevelFormat, BorderStyle, WidthType,
    ExternalHyperlink, TabStopType, UnderlineType
  } = require('docx');
  const fs = require('fs');
  const path = require('path');
  
  const SERIF    = "Georgia";
  const SANS     = "Arial";
  const C_BLACK  = "0A0A0A";
  const C_DARK   = "1A1A2E";
  const C_MID    = "2D3748";
  const C_MUTED  = "6B7280";
  const C_ACCENT = "1A56A0";
  const C_RULE   = "CBD5E0";
  const CW       = 11000;
  const TW       = 10920; // text area width (12240 - 2 × 660 margin)
  
  function lnk(label, url, sz) {
    return new ExternalHyperlink({ link: url, children: [
      new TextRun({ text: label, size: sz||14, font: SANS, color: C_ACCENT,
        underline: { type: UnderlineType.SINGLE } })
    ]});
  }
  function dot() { return new TextRun({ text: "   ·   ", size: 14, font: SANS, color: C_RULE }); }
  
  function sec(label) {
    return new Paragraph({
      spacing: { before: 54, after: 20 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C_ACCENT, space: 2 } },
      children: [new TextRun({ text: label, bold: true, allCaps: true,
        characterSpacing: 50, size: 16, font: SANS, color: C_ACCENT })]
    });
  }
  
  // Left content with right-aligned content on the same line. A borderless
  // two-column table, because docx-preview ignores right tab stops.
  const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  function split(spacing, left, right, rightW) {
    const rw = rightW || 2600;
    const cell = (w, children, align) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Paragraph({ spacing, alignment: align, children })]
    });
    return new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: [TW - rw, rw],
      borders: { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER,
                 insideHorizontal: NO_BORDER, insideVertical: NO_BORDER },
      rows: [new TableRow({ children: [
        cell(TW - rw, left, AlignmentType.LEFT),
        cell(rw, right, AlignmentType.RIGHT),
      ]})]
    });
  }

  function role(title, org, date) {
    return split({ before: 50, after: 4 }, [
        new TextRun({ text: title, bold: true, size: 18, font: SANS, color: C_DARK }),
        new TextRun({ text: "  ·  " + org, size: 16, font: SANS, color: C_MUTED }),
      ], [
        new TextRun({ text: date, size: 15, font: SANS, color: C_MUTED, italics: true }),
      ]);
  }
  
  function proj(name, stack, date, url) {
    const nameEl = url
      ? new ExternalHyperlink({ link: url, children: [new TextRun({
          text: name, bold: true, size: 17, font: SANS, color: C_ACCENT,
          underline: { type: UnderlineType.SINGLE } })] })
      : new TextRun({ text: name, bold: true, size: 17, font: SANS, color: C_DARK });
    return split({ before: 40, after: 4 }, [
        nameEl,
        new TextRun({ text: "  —  " + stack, size: 13, font: SANS, color: C_MUTED }),
      ], [
        new TextRun({ text: date, size: 14, font: SANS, color: C_MUTED, italics: true }),
      ], 900);
  }
  
  // Parse **bold** but deliberately limit — only one bold item per bullet max
  function r(text, sz, base) {
    return text.split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map(p => {
      const b = p.startsWith('**') && p.endsWith('**');
      return new TextRun({ text: b ? p.slice(2,-2) : p, bold: b,
        size: sz||16, font: SANS, color: b ? C_DARK : (base||C_MID) });
    });
  }
  
  function bRich(children) {
    return new Paragraph({
      numbering: { reference: "bul", level: 0 },
      spacing: { before: 7, after: 0, line: 240, lineRule: "auto" },
      children
    });
  }

  function b(text) {
    return bRich(r(text));
  }

  // Inline hyperlink sized to match bullet body text
  function ilnk(label, url) {
    return lnk(label, url, 16);
  }

  function run(text, bold) {
    return new TextRun({ text, bold: !!bold, size: 16, font: SANS, color: bold ? C_DARK : C_MID });
  }
  
  function skRow(label, val) {
    return new TableRow({ children: [
      new TableCell({
        width: { size: 1600, type: WidthType.DXA },
        borders: { top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
                   left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE} },
        margins: { top: 5, bottom: 5, left: 0, right: 70 },
        children: [new Paragraph({ children: [
          new TextRun({ text: label, bold: true, size: 16, font: SANS, color: C_ACCENT })
        ]})]
      }),
      new TableCell({
        width: { size: CW-1600, type: WidthType.DXA },
        borders: { top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
                   left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE} },
        margins: { top: 5, bottom: 5, left: 0, right: 0 },
        children: [new Paragraph({ children: [
          new TextRun({ text: val, size: 16, font: SANS, color: C_MID })
        ]})]
      })
    ]});
  }
  
  const doc = new Document({
    numbering: { config: [{ reference: "bul", levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 260, hanging: 160 } } }
    }]}]},
    styles: { default: { document: { run: { font: SANS, size: 16, color: C_MID } } } },
    sections: [{
      properties: { page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 280, right: 660, bottom: 60, left: 660 }
      }},
      children: [
  
        // NAME
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [
          new TextRun({ text: "Sahil Nawaz", bold: true, size: 42, font: SERIF, color: C_BLACK })
        ]}),
  
        // CONTACT
        new Paragraph({ spacing: { before: 0, after: 0 }, children: [
          new TextRun({ text: "+91 9990874079", size: 14, font: SANS, color: C_MUTED }), dot(),
          lnk("mdsahil622004@gmail.com","mailto:mdsahil622004@gmail.com"), dot(),
          lnk("linkedin.com/in/sahil-nawaz","https://linkedin.com/in/sahil-nawaz"), dot(),
          lnk("github.com/sahil-n06","https://github.com/sahil-n06"), dot(),
          lnk("sahilnawaz.vercel.app","https://sahilnawaz.vercel.app"), dot(),
          new TextRun({ text: "DOB: 06 Feb 2004", size: 14, font: SANS, color: C_MUTED }),
        ]}),
  
        sec("Summary"),
        new Paragraph({ spacing: { before: 16, after: 0 }, children: r(
          "Software Development Engineer at Nikah Forever. I work across the full stack, from REST APIs and distributed data pipelines to AI integrations and production web platforms. B.Tech CSE, First Division, 2026. Looking for software engineering roles in India or remotely.",
          16, C_MID
        )}),

        sec("Experience"),

        role("Software Development Engineer I", "Nikah Forever (Goalwit Technologies Pvt. Ltd.) — Delhi, India", "Dec 2025 – Present"),
        new Paragraph({ spacing: { before: 0, after: 2 }, children: [
          new TextRun({ text: "SDE I, Jun 2026 – Present  ·  SDE Intern, Dec 2025 – May 2026 (converted to full-time)",
            size: 14, font: SANS, color: C_MUTED, italics: true })
        ]}),

        b("Cut a high-traffic API from 6–7 s to **200–300 ms** with query rewrites, composite indexes and an app-level cache, and shipped an account-lifecycle cron that archives deletions across MySQL, Firebase, S3 and Zoho CRM and restores them if the user returns within the retention window."),

        b("Built real-time chat moderation where TypeSafe's Jev model scores messages and flagged ones reach the receiver **blurred with tap-to-reveal**. It fails open on timeout, is gated by app version and stays in sync across MySQL, the Node outbox and Firebase."),

        b("Building **Coco**, an AI triage agent in Slack and Jira that verifies escalations against read-only production MySQL and the PHP, Node and Flutter code before filing tickets, and dedupes them by error signature."),

        b("Own the Auto Web Dialer **end to end**, a lead-prioritisation engine on Exotel WebRTC that scores and queues 50 contacts per agent from ML signals, payment intent and plan-visit data, feeding a multi-gateway payments layer across Razorpay, Cashfree and PayGlocal that picks the gateway at runtime by success rate for that country and currency."),

        b("Architected the AI hiring platform at career.nikahforever.com, where candidates go through **GPT-4o** resume screening, MCQ and subjective rounds, an optional coding challenge and a 15 to 25 minute voice interview, with anti-cheat running GPT-4o Vision and WPM analysis and support for 12 Indian languages through ElevenLabs and Sarvam AI."),

        bRich([
          run("Rebuilt the NikahForever web app ("),
          ilnk("web.nikahforever.com", "https://web.nikahforever.com"),
          run(") from scratch on Next.js 16 and React 19, with OTP authentication, a 4-stage identity check from selfie to document to number to profile, real-time chat and a subscription plus profile-boost funnel, running on "),
          run("Cloudflare edge", true),
          run(" behind a backend gateway that proxies every API call."),
        ]),

        b("Built a GitHub Actions cache-warming pipeline that keeps tens of thousands of SEO pages fresh across 3 regions so Cloudflare serves most traffic from the edge, at a **98%+ success rate** with no always-on server."),

        b("Built GoSalam.com, a travel agency site fully configurable from an admin portal with no code changes, and a feedback intelligence platform on React, Node.js and PostgreSQL that scores call recordings, app-store reviews and support tickets through a 5-stage AI pipeline, cutting manual processing effort by **85%**."),

        role("Web Developer Intern", "Tussle Digital — Gurugram, India", "Jun 2025 – Nov 2025"),
        b("At Tussle Digital, built an AI Chat API on Node.js and Azure OpenAI with short and long-term memory through embeddings, then architected Kanban, Gantt and Workload modules with WebSocket real-time sync that the team used to track 30+ tickets a day, and ran HubSpot EDM campaigns for 5 clients that cut manual email effort by **30%**."),

        role("Freelance Shopify Developer", "Remote", "Nov 2024 – Jan 2025"),
        b("Delivered a Shopify store with custom Liquid templates that, within one month of launch, lifted SEO score by 35%, cut page load time by 28% and grew organic traffic by **40%**."),

        sec("Personal Projects"),

        proj("NEXUS — Sovereign AI Desktop App",
             "Electron · React · TypeScript · node-llama-cpp · MCP", "2025–26",
             "https://github.com/sahil-n06/NEXUS"),
        b("A local LLM desktop workspace that runs entirely on the user's hardware. A hardware-aware pre-warm pipeline probes VRAM at startup and loads the likely model ahead of time, cutting first-token latency by about **2.5 seconds**, and it routes across Llama.cpp, Ollama and the cloud through one interface, with MCP integration adding real-time web search without any external pings."),

        proj("Nudge Dev — AI Developer Tooling Suite",
             "TypeScript · VS Code API · DeepSeek R1 · Groq · Gemini · SQLite", "2025–26",
             "https://open-vsx.org/extension/nudge-dev/nudge-dev"),
        bRich([
          run("Published three VS Code extensions for developer workflows, "),
          ilnk("Nudge Dev", "https://open-vsx.org/extension/nudge-dev/nudge-dev"),
          run(" where six AI personas debate your plan before you write a line, "),
          ilnk("Nudge Junior", "https://open-vsx.org/extension/nudge-dev/nudge-junior"),
          run(" for lightweight in-editor AI nudges, and "),
          ilnk("BranchMap File Tree", "https://open-vsx.org/extension/nudge-dev/branchmap-file-tree"),
          run(" for generating a clean, copy-ready file tree. Combined they have pulled in "),
          run("5k+ downloads", true),
          run(" on Open VSX."),
        ]),

        proj("MailSetu — Email & SMS Delivery SaaS",
             "Next.js · Node.js · PostgreSQL · BullMQ · Redis · Stalwart SMTP", "2026",
             "https://mailssetu.in"),
        b("A transactional email and SMS platform with a REST API supporting sandbox and live modes, SPF, DKIM and DMARC domain verification, HMAC-signed webhook delivery, BullMQ async workers with retry logic, and **Razorpay billing** with GST invoices."),

        sec("Skills"),
        new Table({
          width: { size: CW, type: WidthType.DXA },
          columnWidths: [1760, CW-1760],
          borders: { top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
                     left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
                     insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE} },
          rows: [
            skRow("Languages", "JavaScript, TypeScript, Python, Java, C, HTML/CSS, Bash"),
            skRow("Frameworks", "React.js, Next.js 16, Node.js, Express.js, Tailwind CSS, Prisma ORM, Electron.js"),
            skRow("Databases", "PostgreSQL, MongoDB, MySQL, Firebase, Redis, Supabase, SQLite"),
            skRow("Tools & AI", "Git, GitHub Actions, Cloudflare Workers, OpenAI API, Azure OpenAI, BullMQ, JWT, REST APIs, WebSockets"),
            skRow("DSA", "302 problems solved on LeetCode · CodeChef 1414 rated (Div 3, 2★) · 12 contests"),
          ]
        }),

        sec("Education"),
        split({ before: 26, after: 3 }, [
            new TextRun({ text: "B.Tech, Computer Science and Engineering", bold: true, size: 18, font: SANS, color: C_DARK }),
            new TextRun({ text: "  ·  2022 – 2026", size: 15, font: SANS, color: C_MUTED, italics: true }),
          ], [
            new TextRun({ text: "Lloyd Institute of Engineering and Technology, AKTU  ·  ", size: 15, font: SANS, color: C_MUTED }),
            new TextRun({ text: "First Division", bold: true, size: 15, font: SANS, color: C_DARK }),
          ], 5000),
        split({ before: 14, after: 2 }, [
            new TextRun({ text: "AI & Machine Learning, Credit-Linked Program", bold: true, size: 18, font: SANS, color: C_DARK }),
            new TextRun({ text: "  ·  2025 – 2026", size: 15, font: SANS, color: C_MUTED, italics: true }),
          ], [
            new TextRun({ text: "IIT Guwahati × Masai School  ·  ", size: 15, font: SANS, color: C_MUTED }),
            new TextRun({ text: "CGPA 8.90 / 10", bold: true, size: 15, font: SANS, color: C_DARK }),
          ], 5000),
      ]
    }]
  });
  
  const outDir = path.join(__dirname, 'output');
  const outFile = path.join(outDir, 'Sahil_Nawaz_Resume.docx');

  Packer.toBuffer(doc).then(buf => {
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(outFile, buf);
    console.log('Wrote ' + outFile);
  }).catch(err => {
    console.error(err);
    process.exit(1);
  });