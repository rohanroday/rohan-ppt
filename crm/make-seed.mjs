// Builds crm/seed-data.json: the Phase 1 seed records for the Demo org.
// The records are pushed to CRM through the zoho-crm-demo-dataops MCP server, not by this script.
// Run: node crm/make-seed.mjs
import { writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const TEST_ACCOUNT = {
  Account_Name: "Test Realty Pvt Ltd",
  Industry: "Real Estate",
  Phone: "+91 80 4000 1234",
  Website: "https://testrealty.example.com",
  Billing_City: "Bengaluru",
  Billing_State: "Karnataka",
  Billing_Country: "India",
  Description: "Seed account for the Presentation Builder widget (45 Property Options, 2 Contacts).",
};

const CONTACTS = [
  { First_Name: "Asha", Last_Name: "Menon", Title: "Head of Acquisitions", Email: "asha.menon@testrealty.example.com", Phone: "+91 98450 11111" },
  { First_Name: "Vikram", Last_Name: "Rao", Title: "Investment Manager", Email: "vikram.rao@testrealty.example.com", Phone: "+91 98450 22222" },
];

// 45 rows so a limit of 20 gives 3 slides (20/20/5).
const LOCATIONS = ["Whitefield", "Koramangala", "Indiranagar", "HSR Layout", "Hebbal", "Sarjapur Road", "Electronic City", "Yelahanka", "JP Nagar"];
const TYPES = ["Apartment", "Villa", "Plot", "Commercial"];
const PROPERTIES = Array.from({ length: 45 }, (_, i) => {
  const n = i + 1;
  const type = TYPES[i % TYPES.length];
  const location = LOCATIONS[i % LOCATIONS.length];
  return {
    // Row 7 holds markup so AC-4 (values are escaped, not rendered) can be checked.
    Name: n === 7 ? "Escape test <b>x</b>" : `${location} ${type} ${String(n).padStart(2, "0")}`,
    Location: `${location}, Bengaluru`,
    Price: 2500000 + ((n * 737000) % 9000000),
    Area: 650 + ((n * 113) % 2400),
    Type: type,
  };
});

// Slide HTML. Inline styles are included on purpose, to see what the Rich Text field keeps.
const T1_INTRO = `<h1 style="color:#1f4e79;text-align:center">#Account_Name$</h1>
<p style="text-align:center"><i>Property portfolio prepared by #Owner.name$</i></p>
<ul>
<li><b>Industry:</b> #Industry$</li>
<li><b>City:</b> #Billing_City$</li>
<li><b>Phone:</b> #Phone$</li>
<li><b>Website:</b> #Website$</li>
</ul>`;

const T1_LISTINGS = `<h2 style="color:#1f4e79">Property Listings for #Account.Account_Name$</h2>
<table border="1" style="width:100%;border-collapse:collapse">
<tr style="background-color:#1f4e79;color:#ffffff"><th>S.No</th><th>Property</th><th>Location</th><th>Type</th><th>Area (sq ft)</th><th>Price</th></tr>
#RowCountIterator
<tr><td>#Sno$</td><td>#Name$</td><td>#Location$</td><td>#Type$</td><td>#Area$</td><td>#Price$</td></tr>
$RowCountIterator
</table>`;

const T1_CONTACTS = `<h2 style="color:#1f4e79">Contact Details</h2>
<p>Your contacts at #Account.Account_Name$:</p>
<table border="1" style="width:100%;border-collapse:collapse">
<tr style="background-color:#1f4e79;color:#ffffff"><th>S.No</th><th>Name</th><th>Title</th><th>Email</th><th>Phone</th></tr>
#RowCountIterator
<tr><td>#Sno$</td><td>#Full_Name$</td><td>#Title$</td><td>#Email$</td><td>#Phone$</td></tr>
$RowCountIterator
</table>`;

const T2_SUMMARY = `<h1 style="color:#7a3e00">#Account_Name$ at a glance</h1>
<p>#Description$</p>
<p><b>Owner:</b> #Owner.name$ &nbsp; <b>City:</b> #Billing_City$</p>`;

const T2_SHORTLIST = `<h2 style="color:#7a3e00">Shortlist</h2>
<table border="1" style="width:100%;border-collapse:collapse">
<tr style="background-color:#7a3e00;color:#ffffff"><th>#</th><th>Property</th><th>Price</th></tr>
#RowCountIterator
<tr><td>#Sno$</td><td>#Name$</td><td>#Price$</td></tr>
$RowCountIterator
</table>`;

const TEMPLATES = [
  { Name: "Temp1", Slide_Title: "Introduction", Slide_HTML: T1_INTRO, Module_Name: "Accounts", Slide_Order: 1 },
  { Name: "Temp1", Slide_Title: "Property Listings", Slide_HTML: T1_LISTINGS, Limit_of_Rows: 20, Module_Name: "Property_Options", Slide_Order: 2 },
  { Name: "Temp1", Slide_Title: "Contact Details", Slide_HTML: T1_CONTACTS, Module_Name: "Contacts", Slide_Order: 3 },
  { Name: "Temp2", Slide_Title: "Account Summary", Slide_HTML: T2_SUMMARY, Module_Name: "Accounts", Slide_Order: 1 },
  { Name: "Temp2", Slide_Title: "Shortlist", Slide_HTML: T2_SHORTLIST, Limit_of_Rows: 10, Module_Name: "Property_Options", Slide_Order: 2 },
];

// Same data feeds the widget's mock mode (local dev outside CRM).
const json = JSON.stringify({ account: TEST_ACCOUNT, contacts: CONTACTS, properties: PROPERTIES, templates: TEMPLATES }, null, 2);
for (const rel of ["./seed-data.json", "../ppt_widget/web/src/lib/mock/seed-data.json"]) {
  const out = fileURLToPath(new URL(rel, import.meta.url));
  writeFileSync(out, json);
  console.log(`Wrote ${out}`);
}
console.log(`1 account, ${CONTACTS.length} contacts, ${PROPERTIES.length} properties, ${TEMPLATES.length} template rows`);
