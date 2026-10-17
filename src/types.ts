import type { JSONContent } from "@tiptap/react";
export interface PageSettings {
  size: "A4" | "Letter";
  orientation: "portrait" | "landscape";
  margin: number;
  margins?: { top: number; right: number; bottom: number; left: number };
  header: string;
  footer: string;
  pageNumbers: boolean;
}
export interface Comment {
  id: string;
  quote: string;
  text: string;
  createdAt: string;
  resolved: boolean;
}
export interface StudioDocument {
  id?: string;
  format: "word-studio";
  version: 1;
  title: string;
  content: JSONContent;
  page: PageSettings;
  comments: Comment[];
  updatedAt: string;
}
export const pageMargins = (page: PageSettings) => page.margins ?? {
  top: page.margin, right: page.margin, bottom: page.margin, left: page.margin,
};
export const defaultPage: PageSettings = {
  size: "A4",
  orientation: "portrait",
  margin: 25.4,
  header: "",
  footer: "",
  pageNumbers: true,
};
