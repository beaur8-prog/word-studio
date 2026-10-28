import type { JSONContent } from "@tiptap/react";
export interface PageSettings {
  size: "A4" | "Letter" | "Custom";
  /** Unrotated physical paper dimensions in millimetres. */
  width?: number;
  height?: number;
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
export function paperSize(page: PageSettings): [number, number] {
  return page.size === 'Custom' && Number.isFinite(page.width) && Number.isFinite(page.height)
    ? [page.width!, page.height!] : page.size === 'Letter' ? [215.9, 279.4] : [210, 297];
}
export function paperDimensions(page: PageSettings): [number, number] {
  const [width, height] = paperSize(page);
  return page.orientation === 'landscape' ? [height, width] : [width, height];
}
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
