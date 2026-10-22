// Mock CRM for local development outside Zoho (plain `next dev`, or ?mock=1).
// Uses the same seed data that was loaded into the Demo org (crm/make-seed.mjs).

import type { CrmRecord, TemplateRow } from "../types";
import type { DataSource } from "../zoho";
import seed from "./seed-data.json";

const OWNER = { name: "Demo User", id: "u1" };
const CREATED = "2026-10-22T16:51:10+05:30";

const TEST_ACCOUNT_ID = "mock-acc-1";

const accounts: CrmRecord[] = [
  { id: TEST_ACCOUNT_ID, Owner: OWNER, Created_Time: CREATED, ...seed.account },
  // Sample accounts with no Property Options or Contacts (AC-9).
  ...["Benton (Sample)", "Chanay (Sample)", "Chemel (Sample)", "King (Sample)"].map((name, i) => ({
    id: `mock-acc-${i + 2}`,
    Account_Name: name,
    Owner: OWNER,
    Industry: "Technology",
    Phone: "555-555-5555",
    Website: null,
    Billing_City: ["New Orleans", "Brighton", "Bridgeport", "Baltimore"][i],
    Description: null,
    Created_Time: CREATED,
  })),
];

const related: Record<string, Record<string, CrmRecord[]>> = {
  [TEST_ACCOUNT_ID]: {
    Property_Options: seed.properties.map((p, i) => ({
      id: `mock-po-${i + 1}`,
      Owner: OWNER,
      Account: { id: TEST_ACCOUNT_ID, name: seed.account.Account_Name },
      ...p,
    })),
    Contacts: seed.contacts.map((c, i) => ({
      id: `mock-ct-${i + 1}`,
      Owner: OWNER,
      Full_Name: `${c.First_Name} ${c.Last_Name}`,
      Account_Name: { id: TEST_ACCOUNT_ID, name: seed.account.Account_Name },
      ...c,
    })),
  },
};

const templates: TemplateRow[] = seed.templates.map((t, i) => ({
  id: `mock-tpl-${i + 1}`,
  Name: t.Name,
  Slide_Title: t.Slide_Title,
  Slide_HTML: t.Slide_HTML,
  Limit_of_Rows: "Limit_of_Rows" in t ? (t.Limit_of_Rows as number) : null,
  Module_Name: t.Module_Name,
  Slide_Order: t.Slide_Order,
  Created_Time: CREATED,
}));

const delay = <T,>(value: T, ms = 150) => new Promise<T>((r) => setTimeout(() => r(value), ms));

export const mockSource: DataSource = {
  mode: "mock",
  init: () => delay({ userName: "Demo User (mock data)" }),
  listTemplates: () => delay(templates),
  searchAccounts: (q) => {
    const needle = q.trim().toLowerCase();
    if (needle.length < 2) return delay([]);
    return delay(
      accounts
        .filter((a) => String(a.Account_Name).toLowerCase().includes(needle))
        .slice(0, 20)
        .map((a) => ({ id: String(a.id), name: String(a.Account_Name), secondary: [a.Billing_City, OWNER.name].filter(Boolean).join(" · ") })),
    );
  },
  getAccount: (id) => {
    const a = accounts.find((x) => x.id === id);
    return a ? delay(a) : Promise.reject(new Error("The selected Account no longer exists."));
  },
  getRelated: (accountId, module) => delay(related[accountId]?.[module] ?? []),
  getFormatContext: () =>
    delay({
      dateFormat: "dd/MM/yyyy",
      currencySymbol: "₹",
      locale: "en-IN",
      fieldTypes: {
        Accounts: { Account_Name: "text", Description: "textarea", Website: "website", Billing_City: "text", Industry: "picklist", Phone: "phone" },
        Property_Options: { Price: "currency", Area: "double", Name: "text", Location: "text", Type: "picklist" },
        Contacts: { Full_Name: "text", Title: "text", Email: "email", Phone: "phone" },
      },
    }),
  attachFile: () => delay(undefined, 400),
};
