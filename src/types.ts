export type CellValue = string | number | boolean | null;

export interface PointageData {
  headers: string[];
  rows: CellValue[][];
  fileName: string;
}
