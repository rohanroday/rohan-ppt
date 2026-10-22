// Shared types for CRM data, templates and generated slides.

export type CrmRecord = Record<string, unknown> & { id?: string };

/** One Presentation_Templates record (one slide definition). */
export interface TemplateRow {
  id: string;
  Name: string;
  Slide_Title: string;
  Slide_HTML: string;
  Limit_of_Rows: number | null;
  Module_Name: string;
  Slide_Order: number | null;
  Created_Time?: string;
}

/** All slide definitions that share a Template Name, in deck order. */
export interface Template {
  name: string;
  slides: TemplateRow[];
}

export interface GeneratedSlide {
  /** Slide title, with " (n/N)" added on overflow slides. */
  title: string;
  /** Filled HTML: every CRM value is already escaped. */
  html: string;
  templateRowId: string;
}

export interface GenerationWarning {
  slide: string;
  message: string;
}

/** Field data types per module, e.g. { Property_Options: { Price: "currency" } }. */
export type FieldTypes = Record<string, Record<string, string>>;

export interface FormatContext {
  /** CRM user date format, e.g. "dd/MM/yyyy". Falls back to ISO. */
  dateFormat?: string;
  /** Org currency symbol, e.g. "₹". */
  currencySymbol?: string;
  /** Locale for number formatting, e.g. "en-IN". */
  locale?: string;
  fieldTypes?: FieldTypes;
}

export interface DeckInput {
  template: Template;
  account: CrmRecord;
  /** Related records of the account, keyed by module name (Property_Options, Contacts). */
  related: Record<string, CrmRecord[]>;
  format?: FormatContext;
}

export interface DeckOutput {
  slides: GeneratedSlide[];
  warnings: GenerationWarning[];
}
