import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "C:/Users/Lenovo/OneDrive - Indian Institute of Technology Indian School of Mines Dhanbad/Desktop/CampusParking_system";
const OUT_DIR = path.join(ROOT, "outputs", "campus-parking-presentation");
const PREVIEW_DIR = path.join(ROOT, "tmp", "slides", "campus-parking-presentation", "preview");
const OUTPUT_PPTX = path.join(OUT_DIR, "Campus_Parking_Management_System_Presentation_Safe.pptx");
const HERO_IMAGE = path.join(ROOT, "public", "heritage.jpeg");

const COLORS = {
  bg: "#F5F7F3",
  surface: "#FFFFFF",
  surfaceAlt: "#EEF3EC",
  text: "#173224",
  muted: "#607368",
  green: "#1F7A54",
  greenDark: "#14543A",
  amber: "#D9A441",
  border: "#D6E0D8",
  danger: "#C65A4A",
  white: "#FFFFFF",
};

const FONT = {
  title: "Aptos Display",
  body: "Aptos",
};

const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
});

presentation.theme.colorScheme = {
  name: "Campus Parking",
  themeColors: {
    accent1: COLORS.green,
    accent2: COLORS.amber,
    bg1: COLORS.bg,
    bg2: COLORS.surfaceAlt,
    tx1: COLORS.text,
    tx2: COLORS.muted,
  },
};

function readArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function imageBlob(filePath) {
  const bytes = await fs.readFile(filePath);
  return readArrayBuffer(bytes);
}

function addBase(slide, options = {}) {
  slide.background.fill = options.background || COLORS.bg;

  if (options.footer !== false) {
    const footer = slide.shapes.add({
      geometry: "rect",
      position: { left: 48, top: 678, width: 1184, height: 1 },
      fill: COLORS.border,
      line: { width: 0, fill: COLORS.border },
    });
    footer.text = "";

    addText(slide, {
      left: 56,
      top: 686,
      width: 340,
      height: 20,
      text: "Campus Parking Management System",
      fontSize: 12,
      color: COLORS.muted,
    });
  }
}

function addText(slide, options) {
  const shape = slide.shapes.add({
    geometry: "rect",
    position: {
      left: options.left,
      top: options.top,
      width: options.width,
      height: options.height,
    },
    fill: options.fill || "transparent",
    line: { width: 0, fill: options.fill || "transparent" },
  });
  shape.text = options.text;
  shape.text.typeface = options.typeface || FONT.body;
  shape.text.fontSize = options.fontSize || 22;
  shape.text.color = options.color || COLORS.text;
  shape.text.bold = options.bold || false;
  shape.text.alignment = options.alignment || "left";
  shape.text.verticalAlignment = options.verticalAlignment || "top";
  shape.text.insets = options.insets || { left: 0, right: 0, top: 0, bottom: 0 };
  if (options.italic) shape.text.italic = true;
  return shape;
}

function addPanel(slide, options) {
  const panel = slide.shapes.add({
    geometry: options.geometry || "roundRect",
    position: {
      left: options.left,
      top: options.top,
      width: options.width,
      height: options.height,
    },
    fill: options.fill || COLORS.surface,
    line: {
      width: options.lineWidth ?? 1,
      fill: options.lineFill || COLORS.border,
      style: options.lineStyle || "solid",
    },
    adjustmentList: [{ name: "adj", formula: "val 12000" }],
  });
  if (options.rotation) {
    panel.position = {
      left: options.left,
      top: options.top,
      width: options.width,
      height: options.height,
      rotation: options.rotation,
    };
  }
  panel.text = "";
  return panel;
}

function addTitle(slide, kicker, title, subtitle) {
  addText(slide, {
    left: 64,
    top: 42,
    width: 220,
    height: 20,
    text: kicker,
    fontSize: 13,
    color: COLORS.green,
    bold: true,
    typeface: FONT.body,
  });
  addText(slide, {
    left: 64,
    top: 72,
    width: 780,
    height: 82,
    text: title,
    fontSize: 30,
    color: COLORS.text,
    bold: true,
    typeface: FONT.title,
  });
  if (subtitle) {
    addText(slide, {
      left: 64,
      top: 152,
      width: 900,
      height: 40,
      text: subtitle,
      fontSize: 16,
      color: COLORS.muted,
      typeface: FONT.body,
    });
  }
}

function addBulletList(slide, left, top, width, bullets, color = COLORS.text, size = 22) {
  const lines = bullets.map((item) => `• ${item}`).join("\n");
  addText(slide, {
    left,
    top,
    width,
    height: bullets.length * (size + 18),
    text: lines,
    fontSize: size,
    color,
    typeface: FONT.body,
  });
}

function addRoleCard(slide, left, top, title, body, accent) {
  addPanel(slide, { left, top, width: 350, height: 184, fill: COLORS.surface });
  addPanel(slide, { left: left + 24, top: top + 24, width: 44, height: 44, fill: accent, lineWidth: 0 });
  addText(slide, {
    left: left + 84,
    top: top + 24,
    width: 220,
    height: 32,
    text: title,
    fontSize: 24,
    bold: true,
  });
  addText(slide, {
    left: left + 24,
    top: top + 82,
    width: 296,
    height: 90,
    text: body,
    fontSize: 18,
    color: COLORS.muted,
  });
}

function addScreenshotPlaceholder(slide, left, top, width, height, title, lines) {
  addPanel(slide, {
    left,
    top,
    width,
    height,
    fill: "#FBFCFA",
    lineFill: COLORS.green,
    lineStyle: "dashed",
    lineWidth: 2,
  });
  addText(slide, {
    left: left + 20,
    top: top + 20,
    width: width - 40,
    height: 28,
    text: title,
    fontSize: 20,
    bold: true,
    color: COLORS.greenDark,
  });
  addText(slide, {
    left: left + 20,
    top: top + 60,
    width: width - 40,
    height: height - 80,
    text: lines.join("\n"),
    fontSize: 16,
    color: COLORS.muted,
  });
}

function addFlowBox(slide, left, top, width, height, title, body, fill = COLORS.surface) {
  addPanel(slide, { left, top, width, height, fill });
  addText(slide, {
    left: left + 16,
    top: top + 14,
    width: width - 32,
    height: 28,
    text: title,
    fontSize: 20,
    bold: true,
  });
  addText(slide, {
    left: left + 16,
    top: top + 50,
    width: width - 32,
    height: height - 64,
    text: body,
    fontSize: 16,
    color: COLORS.muted,
  });
}

function addArrow(slide, left, top, width, height) {
  const arrow = slide.shapes.add({
    geometry: "rightArrow",
    position: { left, top, width, height },
    fill: COLORS.amber,
    line: { width: 0, fill: COLORS.amber },
  });
  arrow.text = "";
}

function addTable(slide, left, top, width, headers, rows) {
  const totalRows = rows.length + 1;
  const colWidth = width / headers.length;
  const rowHeight = 46;

  headers.forEach((header, column) => {
    addPanel(slide, {
      left: left + column * colWidth,
      top,
      width: colWidth,
      height: rowHeight,
      fill: COLORS.greenDark,
      lineFill: COLORS.greenDark,
      geometry: "rect",
      lineWidth: 1,
    });
    addText(slide, {
      left: left + column * colWidth + 12,
      top: top + 10,
      width: colWidth - 24,
      height: 24,
      text: header,
      fontSize: 16,
      color: COLORS.white,
      bold: true,
    });
  });

  rows.forEach((row, rowIndex) => {
    row.forEach((cell, column) => {
      addPanel(slide, {
        left: left + column * colWidth,
        top: top + rowHeight * (rowIndex + 1),
        width: colWidth,
        height: rowHeight,
        fill: rowIndex % 2 === 0 ? COLORS.surface : "#F7FAF6",
        lineFill: COLORS.border,
        geometry: "rect",
        lineWidth: 1,
      });
      addText(slide, {
        left: left + column * colWidth + 12,
        top: top + rowHeight * (rowIndex + 1) + 10,
        width: colWidth - 24,
        height: 24,
        text: cell,
        fontSize: 15,
        color: column === row.length - 1 && cell === "PASS" ? COLORS.greenDark : COLORS.text,
        bold: column === row.length - 1,
      });
    });
  });

  return totalRows * rowHeight;
}

async function buildSlides() {
  const heroBlob = await imageBlob(HERO_IMAGE);

  const slide1 = presentation.slides.add();
  slide1.background.fill = "#F3F6F2";
  addPanel(slide1, { left: 0, top: 0, width: 1280, height: 720, fill: "#F3F6F2", geometry: "rect", lineWidth: 0 });
  addPanel(slide1, { left: 58, top: 56, width: 650, height: 608, fill: COLORS.greenDark, lineWidth: 0 });
  addText(slide1, {
    left: 76,
    top: 72,
    width: 240,
    height: 24,
    text: "FINAL YEAR ENGINEERING PROJECT",
    fontSize: 16,
    color: "#E5EFE8",
    bold: true,
  });
  addText(slide1, {
    left: 76,
    top: 118,
    width: 690,
    height: 180,
    text: "Campus Parking\nManagement System",
    fontSize: 48,
    color: COLORS.white,
    bold: true,
    typeface: FONT.title,
  });
  addText(slide1, {
    left: 78,
    top: 318,
    width: 620,
    height: 70,
    text: "Smart, role-based parking reservation and monitoring platform for campus mobility, security, and administrative control.",
    fontSize: 22,
    color: "#EAF3ED",
  });
  const hero = slide1.images.add({ blob: heroBlob, fit: "cover", alt: "Campus hero image" });
  hero.position = { left: 760, top: 72, width: 438, height: 240 };
  addPanel(slide1, { left: 760, top: 72, width: 438, height: 240, fill: COLORS.surface, lineFill: COLORS.border, lineWidth: 1 });
  addPanel(slide1, { left: 760, top: 336, width: 438, height: 188, fill: COLORS.surface, lineFill: COLORS.border, lineWidth: 1 });
  addText(slide1, {
    left: 790,
    top: 362,
    width: 250,
    height: 28,
    text: "Presented By",
    fontSize: 22,
    bold: true,
  });
  addBulletList(slide1, 790, 404, 270, [
    "Team Member 1",
    "Team Member 2",
    "Team Member 3",
    "Team Member 4",
  ], COLORS.text, 18);
  addText(slide1, {
    left: 790,
    top: 514,
    width: 260,
    height: 54,
    text: "Department / College Name\nGuided by: Faculty Mentor",
    fontSize: 17,
    color: COLORS.muted,
  });
  addPanel(slide1, { left: 760, top: 548, width: 438, height: 116, fill: COLORS.amber, lineWidth: 0 });
  addText(slide1, {
    left: 790,
    top: 572,
    width: 290,
    height: 28,
    text: "Demo Credentials",
    fontSize: 22,
    color: "#4B3200",
    bold: true,
  });
  addText(slide1, {
    left: 790,
    top: 610,
    width: 360,
    height: 46,
    text: "User: user@campus.com / password123\nSecurity: security@campus.com / password123\nAdmin: admin@campus.com / password123",
    fontSize: 15,
    color: "#4B3200",
  });
  slide1.speakerNotes.setText("Speaker: Team Member 1\nScript: Good morning. Our project is Campus Parking Management System, a full-stack web application designed to manage campus vehicle parking in a smart and structured way. The system supports user, security, and admin roles, and provides slot reservation, live availability tracking, permit verification, reservation history, and email-based alerts. Our goal was to solve real campus parking inefficiencies using a scalable, secure, and deployment-ready architecture.");

  const slide2 = presentation.slides.add();
  addBase(slide2);
  addTitle(slide2, "SLIDE 2", "Problem Statement", "Why a digital campus parking platform is necessary");
  addBulletList(slide2, 78, 224, 540, [
    "Manual parking handling causes congestion, delay, and inefficient slot utilization.",
    "Users cannot identify zone-wise slot availability in real time.",
    "Security staff need faster permit validation at entry and exit points.",
    "Administrators lack centralized control, analytics, and reservation history.",
    "Campus-specific needs require role-based and zone-based parking workflow.",
  ], COLORS.text, 21);
  addPanel(slide2, { left: 708, top: 214, width: 484, height: 384, fill: "#FFFFFF" });
  addFlowBox(slide2, 742, 250, 128, 92, "Manual Process", "No live visibility\nNo digital record", "#F9FBF8");
  addArrow(slide2, 882, 278, 74, 36);
  addFlowBox(slide2, 966, 250, 128, 92, "Pain Points", "Delay\nConfusion\nMisuse", "#FFF8EF");
  addArrow(slide2, 844, 396, 246, 34);
  addFlowBox(slide2, 742, 462, 352, 98, "Operational Impact", "Wasted time, user dissatisfaction, limited traceability, and poor resource management.", "#F7FAF6");
  slide2.speakerNotes.setText("Speaker: Team Member 2\nScript: The key problem we identified is that campus parking is often handled manually or semi-manually. This creates issues such as uncertain slot availability, inefficient space utilization, difficulty in permit verification, and no central reporting mechanism for administrators. In a large campus environment, these issues directly affect mobility, time efficiency, and security operations. Our project addresses these gaps with a centralized digital system.");

  const slide3 = presentation.slides.add();
  addBase(slide3);
  addTitle(slide3, "SLIDE 3", "Proposed Solution", "A centralized role-based parking platform tailored for campus operations");
  addRoleCard(slide3, 70, 214, "User", "Register, log in, view live slots, reserve parking, and track active permits with custom duration.", "#DFF2E8");
  addRoleCard(slide3, 464, 214, "Security", "Monitor active reservations, verify permits, and update slot occupancy during gate operations.", "#FFF1D7");
  addRoleCard(slide3, 858, 214, "Admin", "Manage zones and slots, monitor usage statistics, and generate reservation reports.", "#E7EEF6");
  addPanel(slide3, { left: 70, top: 442, width: 1138, height: 172, fill: "#F9FBF8" });
  addBulletList(slide3, 96, 472, 1040, [
    "Preconfigured with 14 campus zones and 10 slots per zone for realistic deployment.",
    "Supports live availability tracking, email reminders, and reservation history.",
    "Built as a responsive premium UI with backend APIs and deployment-ready configuration.",
  ], COLORS.text, 20);
  slide3.speakerNotes.setText("Speaker: Team Member 1\nScript: Our solution is a centralized campus parking platform that digitizes the full lifecycle of parking operations. Users can reserve slots based on live availability, security personnel can validate permits quickly, and administrators can manage zones and analyze usage from a single system. We also customized the system for campus needs by creating 14 parking zones with structured slot allocation and role-wise dashboards.");

  const slide4 = presentation.slides.add();
  addBase(slide4);
  addTitle(slide4, "SLIDE 4", "System Architecture", "Integrated full-stack architecture with secure role-based access");
  addFlowBox(slide4, 72, 232, 238, 120, "Frontend Layer", "React-based modular UI\nReusable components\nResponsive layout", "#F7FAF6");
  addFlowBox(slide4, 368, 232, 238, 120, "Backend Layer", "Node.js + Express APIs\nControllers and routes\nBusiness logic", "#F7FAF6");
  addFlowBox(slide4, 664, 232, 238, 120, "Security Layer", "JWT authentication\nRole authorization\nHelmet + CORS", "#F7FAF6");
  addFlowBox(slide4, 960, 232, 238, 120, "Data Layer", "MongoDB + Mongoose\nUsers, zones, slots,\nreservations", "#F7FAF6");
  addArrow(slide4, 314, 274, 42, 24);
  addArrow(slide4, 610, 274, 42, 24);
  addArrow(slide4, 906, 274, 42, 24);
  addPanel(slide4, { left: 132, top: 430, width: 968, height: 150, fill: "#163427" });
  addText(slide4, {
    left: 162,
    top: 454,
    width: 908,
    height: 28,
    text: "Event and Notification Flow",
    fontSize: 22,
    color: COLORS.white,
    bold: true,
  });
  addText(slide4, {
    left: 162,
    top: 498,
    width: 900,
    height: 56,
    text: "Reservation request -> validation -> database write -> slot status update -> background alert service -> reminder / expiry email via Nodemailer",
    fontSize: 19,
    color: "#E4EEE8",
  });
  slide4.speakerNotes.setText("Speaker: Team Member 3\nScript: This slide shows our system architecture. The frontend is implemented using a React-based modular UI, while the backend uses Express APIs for all business operations. Authentication is handled through JWT, and MongoDB stores users, zones, slots, and reservations. We also integrated a background alert service using Nodemailer to send reminder and expiry emails. Since the frontend is served through Express, the application is easy to deploy as one integrated system.");

  const slide5 = presentation.slides.add();
  addBase(slide5);
  addTitle(slide5, "SLIDE 5", "Core Features and Functional Modules", "Role-wise separation improves usability, security, and maintainability");
  addBulletList(slide5, 72, 220, 470, [
    "User Module: registration, login, slot booking, reservation history, active permit access",
    "Security Module: active reservation tracking, permit verification, occupancy updates",
    "Admin Module: zone creation, slot management, reports, and system statistics",
    "Extra Features: email reminders, custom duration booking, premium responsive UI",
  ], COLORS.text, 19);
  addScreenshotPlaceholder(slide5, 592, 214, 190, 200, "User Dashboard Screenshot", [
    "Place booking screen here",
    "Show slot card list and",
    "reservation action",
  ]);
  addScreenshotPlaceholder(slide5, 808, 214, 190, 200, "Security Screenshot", [
    "Place permit verification",
    "or active reservation",
    "screen here",
  ]);
  addScreenshotPlaceholder(slide5, 1024, 214, 190, 200, "Admin Screenshot", [
    "Place zone management",
    "or statistics screen",
    "here",
  ]);
  addPanel(slide5, { left: 592, top: 446, width: 622, height: 162, fill: "#F9FBF8" });
  addText(slide5, {
    left: 620,
    top: 470,
    width: 560,
    height: 28,
    text: "Recommended Visual Layout",
    fontSize: 22,
    bold: true,
  });
  addText(slide5, {
    left: 620,
    top: 512,
    width: 552,
    height: 60,
    text: "Use real screenshots from your application in these three boxes during the final evaluation to make the feature discussion stronger and more credible.",
    fontSize: 18,
    color: COLORS.muted,
  });
  slide5.speakerNotes.setText("Speaker: Team Member 2\nScript: The system is divided into functional modules according to role. The user module focuses on booking and tracking reservations. The security module is designed for permit verification and live parking operations. The admin module manages master data such as zones and slots and also provides statistical reporting. In addition, we included reminder emails, reservation history, and a responsive interface to improve usability across devices.");

  const slide6 = presentation.slides.add();
  addBase(slide6);
  addTitle(slide6, "SLIDE 6", "Technology Stack and Implementation Choices", "Selected for modularity, security, fast development, and deployment readiness");
  addRoleCard(slide6, 76, 224, "Frontend", "React-based modular ES modules, HTM templating, custom CSS, responsive user interface", "#E4F3EC");
  addRoleCard(slide6, 464, 224, "Backend", "Node.js, Express.js, route/controller separation, middleware-based API design", "#FDF0D8");
  addRoleCard(slide6, 852, 224, "Data + Services", "MongoDB, Mongoose, JWT, bcrypt, Nodemailer, Helmet, CORS, QR support", "#E8EEF6");
  addPanel(slide6, { left: 76, top: 454, width: 1132, height: 150, fill: "#FFFFFF" });
  addText(slide6, {
    left: 104,
    top: 478,
    width: 1080,
    height: 28,
    text: "Deployment Status",
    fontSize: 22,
    bold: true,
  });
  addText(slide6, {
    left: 104,
    top: 520,
    width: 1070,
    height: 50,
    text: "The repository includes environment configuration, static asset hosting through Express, and a deployment-ready structure suitable for platforms such as Render.",
    fontSize: 19,
    color: COLORS.muted,
  });
  slide6.speakerNotes.setText("Speaker: Team Member 3\nScript: We selected technologies that are practical, scalable, and aligned with modern web development. On the frontend, we use a React-based modular architecture with reusable components and responsive styling. The backend uses Node.js and Express for API design, while MongoDB with Mongoose handles flexible schema-based data storage. For security, we use JWT authentication, password hashing, and middleware like Helmet and CORS. The project is also deployment-ready for cloud hosting.");

  const slide7 = presentation.slides.add();
  addBase(slide7);
  addTitle(slide7, "SLIDE 7", "Database Design and Data Model", "Schema-based design with clear relationships and validation constraints");
  addFlowBox(slide7, 82, 232, 220, 122, "User", "name\nemail\npassword\nrole\nvehicleNumber", "#F7FAF6");
  addFlowBox(slide7, 356, 232, 220, 122, "ParkingZone", "name\nlocation\ndescription\nisActive", "#F7FAF6");
  addFlowBox(slide7, 630, 232, 220, 122, "ParkingSlot", "zone\nslotNumber\nvehicleType\nisAvailable", "#F7FAF6");
  addFlowBox(slide7, 904, 232, 278, 122, "Reservation", "user, slot, zone\nstartsAt, expiresAt\nstatus, permitCode\nverifiedBy", "#F7FAF6");
  addArrow(slide7, 304, 278, 42, 24);
  addArrow(slide7, 578, 278, 42, 24);
  addArrow(slide7, 852, 278, 42, 24);
  addPanel(slide7, { left: 82, top: 416, width: 1100, height: 182, fill: "#FFFFFF" });
  addBulletList(slide7, 106, 444, 1048, [
    "One user can create multiple reservations, while each reservation maps to one slot and one zone.",
    "Unique constraints exist for user email, permit code, and slot number within a zone.",
    "Business validation limits reservation duration from 10 minutes to 12 hours.",
    "Seed data initializes 14 zones and 140 parking slots for realistic campus usage.",
  ], COLORS.text, 18);
  slide7.speakerNotes.setText("Speaker: Team Member 4\nScript: Our database design follows a clear relationship model. Users, parking zones, slots, and reservations are stored as separate collections. A reservation acts as the transactional entity linking a user to a specific slot in a zone for a defined time interval. We enforced uniqueness for critical fields like email, permit code, and slot number per zone. We also implemented input constraints such as valid reservation durations to improve consistency and reliability.");

  const slide8 = presentation.slides.add();
  addBase(slide8);
  addTitle(slide8, "SLIDE 8", "Testing, Validation, and Reliability Checks", "Functional verification completed across core workflows and backend rules");
  addTable(slide8, 72, 210, 1136, ["Validation Area", "Representative Check", "Status"], [
    ["Authentication", "Register, login, invalid credentials, duplicate email rejection", "PASS"],
    ["Authorization", "JWT-protected routes and role-based access control", "PASS"],
    ["Reservation Logic", "Duration validation, slot availability, booking workflow", "PASS"],
    ["Admin Operations", "Zone creation, slot creation, statistics, reports", "PASS"],
    ["Alert Handling", "Expiry processing and email reminder workflow", "PASS"],
  ]);
  addPanel(slide8, { left: 72, top: 524, width: 1136, height: 88, fill: "#FFF8EF", lineFill: "#F0D298" });
  addText(slide8, {
    left: 98,
    top: 548,
    width: 1080,
    height: 36,
    text: "Current repository status: a formal automated unit/integration test suite is not yet included; validation has been completed through API-level and end-to-end functional testing.",
    fontSize: 18,
    color: "#6B4A0B",
  });
  slide8.speakerNotes.setText("Speaker: Team Member 4\nScript: We validated the project primarily through functional and API-level testing. We checked the full flow from registration to reservation, permit verification, and administrative control. We also validated backend logic such as duplicate account prevention, JWT-based authorization, reservation duration constraints, and slot status updates. At this stage, the repository does not contain a formal automated testing suite, so we present this honestly as a current limitation and future enhancement area.");

  const slide9 = presentation.slides.add();
  addBase(slide9);
  addTitle(slide9, "SLIDE 9", "Live Demonstration Flow", "Recommended sequence for a smooth 2-3 minute faculty demo");
  const demoSteps = [
    ["1. Open App", "Show login and registration page"],
    ["2. User Login", "Open user dashboard and available slots"],
    ["3. Reserve Slot", "Select duration and confirm booking"],
    ["4. Permit View", "Show reservation history and permit details"],
    ["5. Security Login", "Verify active reservation / permit"],
    ["6. Admin Login", "Show zones, slots, and statistics"],
  ];
  demoSteps.forEach(([step, body], index) => {
    const left = 72 + (index % 3) * 376;
    const top = index < 3 ? 232 : 410;
    addFlowBox(slide9, left, top, 332, 126, step, body, index % 2 === 0 ? "#F7FAF6" : "#FFF9F0");
  });
  addPanel(slide9, { left: 72, top: 576, width: 1136, height: 48, fill: "#163427" });
  addText(slide9, {
    left: 96,
    top: 588,
    width: 1088,
    height: 20,
    text: "Demo tip: keep all three role credentials open in notes to avoid delay during evaluation.",
    fontSize: 17,
    color: COLORS.white,
  });
  slide9.speakerNotes.setText("Speaker: Team Member 1\nScript: In the live demo, we will first show the authentication page and then login as a user to demonstrate slot booking. Next, we will show how the reservation appears in history along with permit information. We will then switch to the security role to demonstrate permit verification and finally log in as admin to show zone management and system statistics. This flow demonstrates that the system supports complete end-to-end campus parking operations.\n\nShort demo script: First, we open the application and show the login and registration interface. We log in as a user and select an available parking slot from one of the 14 campus zones. We reserve the slot for a chosen duration and show the reservation details in the dashboard. Next, we switch to the security role and verify the permit from the active reservation list. Finally, we log in as admin to demonstrate zone management, slot creation, and usage statistics.");

  const slide10 = presentation.slides.add();
  addBase(slide10);
  addTitle(slide10, "SLIDE 10", "Conclusion and Future Scope", "Campus-focused parking digitization with strong scope for extension");
  addPanel(slide10, { left: 72, top: 218, width: 540, height: 360, fill: "#FFFFFF" });
  addText(slide10, {
    left: 98,
    top: 246,
    width: 420,
    height: 28,
    text: "Conclusion",
    fontSize: 24,
    bold: true,
  });
  addBulletList(slide10, 98, 290, 468, [
    "Delivered a functional full-stack web application for campus parking management.",
    "Integrated booking, verification, reporting, and notification workflows in one platform.",
    "Improved visibility, control, and operational efficiency of parking resources.",
  ], COLORS.text, 19);
  addPanel(slide10, { left: 668, top: 218, width: 540, height: 360, fill: "#F9FBF8" });
  addText(slide10, {
    left: 694,
    top: 246,
    width: 340,
    height: 28,
    text: "Future Scope",
    fontSize: 24,
    bold: true,
  });
  addBulletList(slide10, 694, 290, 470, [
    "Automated unit and integration testing",
    "QR scanner integration for real gate validation",
    "Payment workflow for visitor parking",
    "RFID / IoT-based live occupancy detection",
    "Predictive analytics for peak-time demand",
    "Mobile application support",
  ], COLORS.text, 18);
  addPanel(slide10, { left: 72, top: 610, width: 1136, height: 32, fill: COLORS.greenDark, lineWidth: 0 });
  addText(slide10, {
    left: 410,
    top: 616,
    width: 480,
    height: 16,
    text: "Thank you",
    fontSize: 18,
    color: COLORS.white,
    bold: true,
    alignment: "center",
  });
  slide10.speakerNotes.setText("Speaker: Team Member 2 or Team Leader\nScript: To conclude, our project delivers a practical and technically sound solution for campus parking digitization. It combines user convenience, security operations, and administrative control into one integrated platform. The project is already functional and deployment-ready, and it can be extended further with IoT-based occupancy sensing, QR gate automation, mobile support, and predictive analytics to make campus mobility even smarter.\n\nPossible faculty questions and answers:\n1. Why MongoDB? MongoDB fit our document-oriented entities and worked well with Mongoose validation and references.\n2. How is security handled? JWT authentication, bcrypt hashing, middleware-based role authorization, Helmet, and CORS.\n3. How do you avoid double booking? The backend checks slot availability before reservation and marks the slot unavailable immediately after booking.\n4. How are expired reservations handled? A background alert service updates status and frees the slot when expiry time is reached.\n5. What is the current campus scale? Seed data includes 14 zones and 140 total slots.\n6. Is the system scalable? Yes, due to modular APIs, separate schemas, and deployment-ready architecture.\n7. Do you have automated tests? Not yet as a formal suite; current validation is functional and API-level.\n8. Why is this better than manual parking? It improves visibility, traceability, access control, and resource efficiency.");
}

async function exportDeck() {
  await fs.mkdir(OUT_DIR, { recursive: true });
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  await buildSlides();

  for (let i = 0; i < presentation.slides.count; i += 1) {
    const slide = presentation.slides.getItem(i);
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    const previewPath = path.join(PREVIEW_DIR, `slide-${String(i + 1).padStart(2, "0")}.png`);

    if (png?.save) {
      await png.save(previewPath);
      continue;
    }

    if (png instanceof ArrayBuffer) {
      await fs.writeFile(previewPath, Buffer.from(png));
      continue;
    }

    if (ArrayBuffer.isView(png)) {
      await fs.writeFile(
        previewPath,
        Buffer.from(png.buffer, png.byteOffset, png.byteLength)
      );
      continue;
    }

    if (png?.arrayBuffer) {
      const buffer = await png.arrayBuffer();
      await fs.writeFile(previewPath, Buffer.from(buffer));
      continue;
    }

    if (png?.data) {
      await fs.writeFile(previewPath, Buffer.from(png.data));
      continue;
    }

    throw new Error(`Unsupported PNG export result: ${Object.prototype.toString.call(png)}`);
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PPTX);
  console.log(`Saved deck to ${OUTPUT_PPTX}`);
}

await exportDeck();
