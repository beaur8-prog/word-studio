import type { JSONContent } from "@tiptap/react";
export interface PageSettings {
  size: "A4" | "Letter";
  orientation: "portrait" | "landscape";
  margin: number;
  margins?: { top: number; right: number; bottom: number; left: number };
  header: string;
  footer: string;
  pageNumbers: boolean;
  differentFirstPage?: boolean;
  differentOddEven?: boolean;
  firstHeader?: string;
  firstFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstPageNumbers?: boolean;
  evenPageNumbers?: boolean;
  headerDistance?: number;
  footerDistance?: number;
  gutter?: number;
  gutterTop?: boolean;
  breakType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  columns?: { count: number; space: number; widths?: number[]; balance?: boolean };
  headerHtml?: string;
  footerHtml?: string;
  firstHeaderHtml?: string;
  firstFooterHtml?: string;
  evenHeaderHtml?: string;
  evenFooterHtml?: string;
}
export interface Comment {
  id: string;
  quote: string;
  text: string;
  createdAt: string;
  resolved: boolean;
  author?: string;
  initials?: string;
  html?: string;
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
export const basePageMargins = (page: PageSettings) => page.margins ?? {
  top: page.margin, right: page.margin, bottom: page.margin, left: page.margin,
};
export const pageMargins = (page: PageSettings) => {
  const margins = { ...basePageMargins(page) };
  margins[page.gutterTop ? 'top' : 'left'] += page.gutter ?? 0;
  return margins;
};
export function pageFurniture(page: PageSettings, number: number, first = number === 1) {
  const prefix = first && page.differentFirstPage ? 'first' : number % 2 === 0 && page.differentOddEven ? 'even' : null;
  return prefix ? { header: page[`${prefix}Header`] ?? '', footer: page[`${prefix}Footer`] ?? '', pageNumbers: page[`${prefix}PageNumbers`] ?? false,
    headerHtml: page[`${prefix}HeaderHtml`], footerHtml: page[`${prefix}FooterHtml`] }
    : { header: page.header, footer: page.footer, pageNumbers: page.pageNumbers, headerHtml: page.headerHtml, footerHtml: page.footerHtml };
}
export const defaultPage: PageSettings = {
  size: "A4",
  orientation: "portrait",
  margin: 25.4,
  header: "",
  footer: "",
  pageNumbers: true,
};
