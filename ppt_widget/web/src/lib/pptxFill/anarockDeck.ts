// Turns Anarock option data into the slide specs for the Anarock template (see scripts/make-anarock-template.mjs).
// Template slides in order: cover, map (conventional), map (co-working), summary (conv), summary (co-working),
// option detail (conv), option detail (co-working), closing.

import type { SlideSpec } from "./fillTemplate";

export const ANAROCK_ROLES = ["cover", "mapConv", "mapCow", "sumConv", "sumCow", "detConv", "detCow", "closing"];

/** Rows on one summary slide of the template. More options continue on a copy of the slide. */
export const SUMMARY_ROWS_PER_SLIDE = 10;

export type OptionKind = "conventional" | "coworking";

/** One proposed option. Values are stored raw (e.g. rent "325"); the deck adds units where the template shows them. */
export interface AnarockOption {
  kind: OptionKind;
  name: string; // building name (conventional) or building the centre is in (co-working)
  operator?: string; // co-working only
  city?: string; // subtitle under the title
  microMarket?: string;
  floor?: string;
  structure?: string;
  floorPlate?: string;
  area?: string;
  rent?: string; // INR/SF/Month
  cam?: string; // INR/SF/Month
  handover?: string;
  availability?: string;
  status?: string;
  seats?: string;
  price?: string; // INR/Seat/Month
  wsSize?: string;
  lockin?: string;
  escalation?: string;
  // detail slides
  ownership?: string; developer?: string; projectSize?: string; efficiency?: string;
  hvac?: string; year?: string; power?: string; cert?: string; elevators?: string; grade?: string;
  metro?: string; airport?: string;
  leaseTerm?: string; opHours?: string; parkingRatio?: string; parkingCharges?: string; addnParking?: string; deposit?: string;
  centerArea?: string; capacity?: string; density?: string; configuration?: string; arrangement?: string;
  notice?: string; setupCost?: string; parkingSlots?: string; modifications?: string; credits?: string; meetingRoomCharges?: string;
  /** Up to 4 lines (conventional project highlights / co-working centre highlights). */
  highlights?: string[];
  /** Up to 3 lines, co-working project highlights. */
  projectHighlights?: string[];
  /** Up to 6 names, conventional tenant profile. */
  tenants?: string[];
  buildingImage?: string; // data URL
  floorImage?: string; // data URL
}

export interface AnarockInput {
  client: string;
  requirementType: string;
  date: Date;
  user: { name: string; email: string; phone: string };
  mapImage?: string; // data URL, used on both map slides
  options: AnarockOption[];
}

const ORD = ["th", "st", "nd", "rd"];
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

/** 26th October 2026, like the template's cover. */
export function formatCoverDate(d: Date): string {
  const day = d.getDate();
  const suffix = day % 100 >= 11 && day % 100 <= 13 ? "th" : ORD[day % 10 < 4 ? day % 10 : 0];
  return `${day}${suffix} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

const v = (s: string | undefined) => s ?? "";
const withUnit = (s: string | undefined, fmt: (x: string) => string) => (s ? fmt(s) : "");
const slot = (list: string[] | undefined, prefix: string, count: number) =>
  Object.fromEntries(Array.from({ length: count }, (_, i) => [`${prefix}${i + 1}`, v(list?.[i])]));

function conventionalTokens(o: AnarockOption, n: number): Record<string, string> {
  return {
    title: `Option ${n}: ${o.name}`, city: v(o.city),
    ...slot(o.highlights, "hl", 4), ...slot(o.tenants, "tp", 6),
    ownership: v(o.ownership), developer: v(o.developer), project_size: v(o.projectSize), floor_plate: v(o.floorPlate),
    structure: v(o.structure), efficiency: v(o.efficiency),
    floor: v(o.floor), area: withUnit(o.area, (x) => `${x} SF`), handover: v(o.handover), availability: v(o.availability),
    metro: v(o.metro), airport: v(o.airport),
    hvac: v(o.hvac), year: v(o.year), power: v(o.power), cert: v(o.cert), elevators: v(o.elevators),
    lease_term: v(o.leaseTerm), lockin: v(o.lockin), rent: withUnit(o.rent, (x) => `INR ${x}/SF/Month`),
    maintenance: withUnit(o.cam, (x) => `INR ${x} /SF/Month`), op_hours: v(o.opHours), escalation: v(o.escalation),
    parking_ratio: v(o.parkingRatio), parking_charges: withUnit(o.parkingCharges, (x) => `INR ${x}/Slot/Month`),
    addn_parking: v(o.addnParking), deposit: v(o.deposit),
  };
}

function coworkingTokens(o: AnarockOption, n: number): Record<string, string> {
  const hours = v(o.opHours).split(/\r?\n/);
  return {
    title: `Option ${n}: ${o.operator ? `${o.operator}, ` : ""}${o.name}`, city: v(o.city),
    ...slot(o.highlights, "ch", 4), ...slot(o.projectHighlights, "ph", 3),
    center_area: v(o.centerArea), capacity: v(o.capacity), floor: v(o.floor), ws_size: v(o.wsSize), density: v(o.density),
    op_hours1: v(hours[0]), op_hours2: v(hours[1]),
    seats: v(o.seats), configuration: v(o.configuration), arrangement: v(o.arrangement), availability: v(o.availability),
    metro: v(o.metro), airport: v(o.airport),
    developer: v(o.developer), year: v(o.year), grade: v(o.grade), cert: v(o.cert), elevators: v(o.elevators),
    lockin: v(o.lockin), notice: v(o.notice), deposit: v(o.deposit), escalation: v(o.escalation),
    price: withUnit(o.price, (x) => `INR ${x}/Seat/Month`), setup_cost: v(o.setupCost), parking_slots: v(o.parkingSlots),
    parking_charges: withUnit(o.parkingCharges, (x) => `INR ${x} /Slot/Month`), modifications: v(o.modifications),
    credits: v(o.credits), meeting_room_charges: v(o.meetingRoomCharges),
  };
}

const images = (pairs: [string, string | undefined][]) =>
  Object.fromEntries(pairs.filter((p): p is [string, string] => !!p[1]));

/** Slide specs for the whole deck: cover, then conventional slides, then co-working slides, then the closing slide. */
export function buildAnarockSpecs(input: AnarockInput): SlideSpec[] {
  const specs: SlideSpec[] = [
    { role: "cover", tokens: { date: formatCoverDate(input.date), client: input.client, type: input.requirementType } },
  ];

  const conv = input.options.filter((o) => o.kind === "conventional");
  const cow = input.options.filter((o) => o.kind === "coworking");

  if (conv.length) {
    specs.push({ role: "mapConv", table: conv.map((o) => [o.name, v(o.microMarket)]), images: images([["Map", input.mapImage]]) });
    for (let i = 0; i < conv.length; i += SUMMARY_ROWS_PER_SLIDE) {
      specs.push({
        role: "sumConv",
        table: conv.slice(i, i + SUMMARY_ROWS_PER_SLIDE).map((o) => [
          o.name, v(o.microMarket), v(o.structure), v(o.floor), v(o.floorPlate), v(o.area), v(o.rent), v(o.cam),
          v(o.handover), v(o.availability), v(o.status),
        ]),
      });
    }
    conv.forEach((o, i) =>
      specs.push({
        role: "detConv",
        tokens: conventionalTokens(o, i + 1),
        images: images([["Project Plan", o.buildingImage], ["Floor Plan", o.floorImage]]),
      }),
    );
  }

  if (cow.length) {
    specs.push({ role: "mapCow", table: cow.map((o) => [v(o.operator), o.name, v(o.microMarket)]), images: images([["Map", input.mapImage]]) });
    for (let i = 0; i < cow.length; i += SUMMARY_ROWS_PER_SLIDE) {
      specs.push({
        role: "sumCow",
        table: cow.slice(i, i + SUMMARY_ROWS_PER_SLIDE).map((o) => [
          v(o.operator), o.name, v(o.microMarket), v(o.floor), v(o.seats), v(o.price), v(o.wsSize), v(o.lockin), v(o.escalation),
        ]),
      });
    }
    cow.forEach((o, i) =>
      specs.push({
        role: "detCow",
        tokens: coworkingTokens(o, i + 1),
        images: images([["Building Picture", o.buildingImage], ["Centre Layout", o.floorImage]]),
      }),
    );
  }

  specs.push({ role: "closing", tokens: { user_name: input.user.name, user_email: input.user.email, user_phone: input.user.phone } });
  return specs;
}
