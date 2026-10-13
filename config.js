/*
  Van Reservations configuration
  ------------------------------
  1. Paste the two Power Automate "HTTP URL" values below (see README.md).
  2. Leave them blank to run in demo mode (bookings save only in your browser).
*/
window.VAN_CONFIG = {
  // Power Automate flow that returns all active reservations
  listUrl: "",

  // Power Automate flow that checks for conflicts and saves a new reservation
  bookUrl: "",

  // W&L Microsoft 365 (Outlook) sign-in. See README.md, "Set up W&L sign-in".
  // Leave clientId blank to skip sign-in (demo or testing only).
  auth: {
    clientId: "",          // Application (client) ID from Microsoft Entra
    tenantId: "",          // Directory (tenant) ID from Microsoft Entra
    redirectUri: "",       // Leave blank to use this page's address
    flowScope: ""          // Optional, advanced: see README "Locking the booking flow"
  },

  // Only accounts at this domain can book
  emailDomain: "wlu.edu",

  // Optional: longest reservation allowed, in days. Use null for no limit.
  maxDays: null,

  // The fleet. "id" must match the Van column in the SharePoint list.
  vans: [
    { id: "97",  name: "Van 97" },
    { id: "140", name: "Van 140" },
    { id: "203", name: "Van 203" },
    { id: "206", name: "Van 206" },
    { id: "207", name: "Van 207" },
    { id: "208", name: "Mini Van 208" },
    { id: "209", name: "Mini Van 209" }
  ],

  // Team list for the booking form. Edit freely. "Other" is added automatically.
  teams: [
    "Athletics Administration",
    "Baseball",
    "Men's Basketball",
    "Women's Basketball",
    "Men's Cross Country",
    "Women's Cross Country",
    "Equestrian",
    "Field Hockey",
    "Football",
    "Men's Golf",
    "Women's Golf",
    "Men's Lacrosse",
    "Women's Lacrosse",
    "Men's Soccer",
    "Women's Soccer",
    "Softball",
    "Men's Swimming and Diving",
    "Women's Swimming and Diving",
    "Men's Tennis",
    "Women's Tennis",
    "Men's Track and Field",
    "Women's Track and Field",
    "Volleyball",
    "Wrestling"
  ]
};
