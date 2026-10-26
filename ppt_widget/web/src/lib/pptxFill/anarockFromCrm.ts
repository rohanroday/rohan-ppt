// CRM records -> Anarock deck. Office_Options records become options; pictures and the template file
// are downloaded through the SDK (the template itself is never stored in this repo).

import { fileIdOf } from "../imageResolver";
import type { CrmRecord, Template } from "../types";
import type { DataSource } from "../zoho";
import { ANAROCK_ROLES, buildAnarockSpecs, type AnarockOption } from "./anarockDeck";
import { fillTemplate } from "./fillTemplate";
import { outlineSlides, type SlideOutline } from "./outline";

export const OFFICE_OPTIONS_MODULE = "Office_Options";
const CONCURRENCY = 4;

const s = (v: unknown): string => (v === null || v === undefined ? "" : typeof v === "object" ? "" : String(v));
const lines = (v: unknown): string[] => s(v).split(/\r?\n/).map((x) => x.trim()).filter(Boolean);

/** The id of the template's file when this template is a real .pptx (an uploaded file), else "". */
export function templateFileId(t: Template): string {
  for (const row of t.slides) {
    const id = fileIdOf(row.Template_File);
    if (id) return id;
  }
  return "";
}

/** One Office_Options record -> option (pictures are added separately). */
export function toAnarockOption(r: CrmRecord): AnarockOption {
  return {
    kind: /^co-?working/i.test(s(r.Option_Type).trim()) ? "coworking" : "conventional",
    name: s(r.Name), operator: s(r.Operator_Name), city: s(r.City), microMarket: s(r.Micro_Market),
    floor: s(r.Proposed_Floor), structure: s(r.Building_Structure), floorPlate: s(r.Typical_Floor_Plate),
    area: s(r.Proposed_Super_Area), rent: s(r.Quoted_Rent), cam: s(r.CAM_Charges), handover: s(r.State_of_Handover),
    availability: s(r.Availability_Timeline), status: s(r.Project_Status), seats: s(r.Billable_Seats), price: s(r.Quoted_Price),
    wsSize: s(r.Workstation_Size), lockin: s(r.Lock_in_Period), escalation: s(r.Annual_Escalation),
    ownership: s(r.Ownership_Structure), developer: s(r.Developer), projectSize: s(r.Total_Project_Size),
    efficiency: s(r.Quoted_Efficiency), hvac: s(r.HVAC_System), year: s(r.Year_of_Construction), power: s(r.Power_Back_Up),
    cert: s(r.Building_Certification), elevators: s(r.No_of_Elevators), grade: s(r.Building_Grade),
    metro: s(r.Metro_Distance), airport: s(r.Airport_Distance), leaseTerm: s(r.Lease_Term), opHours: s(r.Operational_Hours),
    parkingRatio: s(r.Parking_Ratio), parkingCharges: s(r.Parking_Charges), addnParking: s(r.Addn_Parking_Charges),
    deposit: s(r.Security_Deposit), centerArea: s(r.Centre_Area), capacity: s(r.Seating_Capacity),
    density: s(r.Density_Per_Workstation), configuration: s(r.Configuration), arrangement: s(r.Space_Arrangement),
    notice: s(r.Notice_Period), setupCost: s(r.One_Time_Set_Up_Cost), parkingSlots: s(r.Car_Parking_Slots),
    modifications: s(r.Cost_of_Modifications), credits: s(r.No_of_Credits), meetingRoomCharges: s(r.Meeting_Room_Charges),
    highlights: lines(r.Highlights), projectHighlights: lines(r.Project_Highlights), tenants: lines(r.Tenant_Profile),
  };
}

export interface AnarockResult {
  data: Uint8Array;
  /** Every text value and table per slide, for checking the data before export. */
  outline: SlideOutline[];
  warnings: string[];
  slideCount: number;
  conventional: number;
  coworking: number;
}

/** Downloads each distinct picture once, a few at a time. Failures become warnings, never errors. */
async function loadPictures(source: DataSource, ids: string[], warnings: string[]): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const queue = [...new Set(ids.filter(Boolean))];
  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, queue.length) }, async () => {
      for (let id = queue.shift(); id !== undefined; id = queue.shift()) {
        try {
          out.set(id, await source.getImage(id));
        } catch (e) {
          warnings.push(`A picture could not be loaded and was left as the template placeholder${e instanceof Error && e.message ? ` (${e.message.slice(0, 80)})` : ""}.`);
        }
      }
    }),
  );
  return out;
}

export async function generateAnarockDeck(args: {
  source: DataSource;
  account: CrmRecord;
  options: CrmRecord[];
  templateFileId: string;
  date?: Date;
}): Promise<AnarockResult> {
  const { source, account, options } = args;
  const warnings: string[] = [];

  const sorted = [...options].sort((a, b) => (Number(a.Option_Order) || 9999) - (Number(b.Option_Order) || 9999));
  const mapId = fileIdOf(account.Location_Map);
  const buildingIds = sorted.map((r) => fileIdOf(r.Building_Picture));
  const floorIds = sorted.map((r) => fileIdOf(r.Floor_Plan));

  const [template, user, pictures] = await Promise.all([
    source.getFileBytes(args.templateFileId),
    source.getUserDetails().catch(() => ({ name: "", email: "", phone: "" })),
    loadPictures(source, [mapId, ...buildingIds, ...floorIds], warnings),
  ]);

  const opts = sorted.map((r, i): AnarockOption => ({
    ...toAnarockOption(r),
    buildingImage: pictures.get(buildingIds[i]),
    floorImage: pictures.get(floorIds[i]),
  }));
  if (!mapId) warnings.push("The Account has no Location Map picture, so the map slides keep the template's placeholder.");
  opts.forEach((o) => {
    if (!o.buildingImage) warnings.push(`${o.name}: no Building Picture, so the template placeholder is kept.`);
    if (!o.floorImage) warnings.push(`${o.name}: no Floor Plan picture, so the template placeholder is kept.`);
  });
  if (opts.length === 0) warnings.push("This Account has no Office Options, so the deck has only the cover and closing slides.");

  const specs = buildAnarockSpecs({
    client: s(account.Account_Name),
    requirementType: s(account.Requirement_Type),
    date: args.date ?? new Date(),
    user,
    mapImage: pictures.get(mapId),
    options: opts,
  });
  const filled = await fillTemplate(template, ANAROCK_ROLES, specs);
  return {
    data: filled.data,
    outline: await outlineSlides(filled.data),
    warnings: [...warnings, ...filled.warnings],
    slideCount: specs.length,
    conventional: opts.filter((o) => o.kind === "conventional").length,
    coworking: opts.filter((o) => o.kind === "coworking").length,
  };
}
